<?php

namespace Kolay\XlsxStream\Sketches;

/**
 * Misra-Gries frequent-items sketch (KXSI "TOPK" payload) — the top-K
 * most frequent values of a column in bounded memory, mergeable across
 * sheets / shards / stitched files (Agarwal-Cormode, *Mergeable
 * Summaries*, 2012).
 *
 * Values are caller-supplied canonical strings (the sketch is
 * encoding-agnostic; canonicalization lives with the producer, as with
 * HyperLogLog). At most `k` (value → count) counters are held. The
 * guarantee: a reported count never OVER-estimates the true count and
 * under-estimates it by at most N/k, and every value with true frequency
 * above N/k is retained — so the true heavy hitters always survive.
 *
 * **Saturated bit.** If no counter was ever evicted (cardinality ≤ k),
 * the surviving counts are EXACT and `saturated()` is false — the sketch
 * is then the complete, exact categorical distribution. The first
 * eviction (a decrement-all, or a merge that must prune) flips it true,
 * after which counts carry the N/k error bound. Merge rule:
 * `saturated_out = a.saturated OR b.saturated OR (the merge prunes)`.
 */
class MisraGries
{
    public const FORMAT_VERSION = 1;

    public const DEFAULT_K = 64;

    private const MIN_K = 1;

    private const MAX_K = 65535; // fits the uint16 wire field

    private int $k;

    /** @var array<string, int> value => stored count */
    private array $counters = [];

    private bool $saturated = false;

    public function __construct(int $k = self::DEFAULT_K)
    {
        if ($k < self::MIN_K || $k > self::MAX_K) {
            throw new \InvalidArgumentException("Misra-Gries k must be within [1, 65535]; got {$k}");
        }
        $this->k = $k;
    }

    /**
     * Fold one value's occurrence into the sketch.
     */
    public function add(string $value): void
    {
        if (isset($this->counters[$value])) {
            $this->counters[$value]++;

            return;
        }
        if (\count($this->counters) < $this->k) {
            $this->counters[$value] = 1;

            return;
        }
        // All k slots full and this value unmonitored: decrement every
        // counter (the eviction event), dropping any that reach zero. The
        // new value is not stored — its unit cancels one decrement.
        $this->saturated = true;
        foreach ($this->counters as $key => $count) {
            if ($count <= 1) {
                unset($this->counters[$key]);
            } else {
                $this->counters[$key] = $count - 1;
            }
        }
    }

    /**
     * @param  iterable<string>  $values
     */
    public function addMany(iterable $values): void
    {
        foreach ($values as $value) {
            $this->add($value);
        }
    }

    /**
     * Merge another sketch into this one (must share k). Sums common
     * counters, then — if the union exceeds k — subtracts the (k+1)-th
     * largest count from every counter and drops the non-positive ones,
     * which preserves the N/k bound over the combined stream.
     */
    public function merge(MisraGries $other): void
    {
        if ($other->k !== $this->k) {
            throw new \InvalidArgumentException("cannot merge Misra-Gries sketches with different k ({$this->k} vs {$other->k})");
        }

        foreach ($other->counters as $value => $count) {
            $this->counters[$value] = ($this->counters[$value] ?? 0) + $count;
        }
        $this->saturated = $this->saturated || $other->saturated;

        if (\count($this->counters) > $this->k) {
            $counts = array_values($this->counters);
            rsort($counts);
            $threshold = $counts[$this->k]; // (k+1)-th largest, 0-indexed
            foreach ($this->counters as $value => $count) {
                $reduced = $count - $threshold;
                if ($reduced > 0) {
                    $this->counters[$value] = $reduced;
                } else {
                    unset($this->counters[$value]);
                }
            }
            $this->saturated = true;
        }
    }

    /**
     * The tracked values with their (approximate unless saturated is
     * false) counts, ordered by count descending then value ascending —
     * a total order, so the output is deterministic across runs.
     *
     * @return list<array{value: string, count: int}>
     */
    public function topValues(): array
    {
        $pairs = [];
        foreach ($this->counters as $value => $count) {
            // PHP coerces decimal-integer-looking array keys to int; cast
            // back so a numeric value ('5') is always exposed as a string.
            $pairs[] = ['value' => (string) $value, 'count' => $count];
        }
        usort($pairs, static function (array $a, array $b): int {
            return $b['count'] <=> $a['count'] ?: strcmp($a['value'], $b['value']);
        });

        return $pairs;
    }

    public function saturated(): bool
    {
        return $this->saturated;
    }

    public function k(): int
    {
        return $this->k;
    }

    /**
     * Serialized form: format version, k, a flags byte (bit0 = saturated),
     * the counter count, then each counter as count + length-prefixed
     * value, in `topValues()` order so the bytes are deterministic.
     */
    public function serialize(): string
    {
        $pairs = $this->topValues();
        $out = pack('vv', self::FORMAT_VERSION, $this->k);
        $out .= \chr($this->saturated ? 1 : 0);
        $out .= pack('V', \count($pairs));
        foreach ($pairs as $pair) {
            $out .= pack('V', $pair['count']);
            $out .= pack('v', \strlen($pair['value'])).$pair['value'];
        }

        return $out;
    }

    /**
     * Inverse of serialize(). The payload is untrusted (it rides in a
     * sidecar an attacker can rewrite with a valid CRC), so every property
     * the sketch relies on is verified before an instance exists: version,
     * k in range, counter count ≤ k, every length inside the payload, and
     * counts positive.
     */
    public static function deserialize(string $payload): self
    {
        $len = \strlen($payload);
        if ($len < 7) {
            throw new \InvalidArgumentException('Misra-Gries payload too short');
        }

        $head = unpack('vversion/vk/Cflags', substr($payload, 0, 5));
        if ($head['version'] !== self::FORMAT_VERSION) {
            throw new \InvalidArgumentException("unsupported Misra-Gries format version {$head['version']}");
        }
        if (($head['flags'] & ~0x01) !== 0) {
            throw new \InvalidArgumentException('Misra-Gries reserved flag bits must be zero');
        }

        $sketch = new self($head['k']);
        $sketch->saturated = ($head['flags'] & 0x01) === 0x01;

        $counterCount = unpack('V', substr($payload, 5, 4))[1];
        if ($counterCount > $head['k']) {
            throw new \InvalidArgumentException("Misra-Gries counter count {$counterCount} exceeds k {$head['k']}");
        }

        $cursor = 9;
        for ($i = 0; $i < $counterCount; $i++) {
            if ($cursor + 6 > $len) {
                throw new \InvalidArgumentException('Misra-Gries payload truncated in counter header');
            }
            $count = unpack('V', substr($payload, $cursor, 4))[1];
            $valueLen = unpack('v', substr($payload, $cursor + 4, 2))[1];
            $cursor += 6;
            if ($count < 1) {
                throw new \InvalidArgumentException('Misra-Gries counter must be positive');
            }
            if ($cursor + $valueLen > $len) {
                throw new \InvalidArgumentException('Misra-Gries payload truncated in counter value');
            }
            $value = substr($payload, $cursor, $valueLen);
            $cursor += $valueLen;
            // A duplicate key in a crafted payload would silently overwrite;
            // reject it so the decoded state matches a real sketch.
            if (isset($sketch->counters[$value])) {
                throw new \InvalidArgumentException('Misra-Gries payload has duplicate value');
            }
            $sketch->counters[$value] = $count;
        }

        if ($cursor !== $len) {
            throw new \InvalidArgumentException('Misra-Gries payload has trailing bytes');
        }

        return $sketch;
    }
}

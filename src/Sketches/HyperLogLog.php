<?php

namespace Kolay\XlsxStream\Sketches;

/**
 * HyperLogLog (Flajolet et al.) — a fixed-size sketch that answers
 * approximate distinct counts over a stream of strings.
 *
 * Configuration is the classic estimator with p=11 by default:
 * m = 2^11 = 2048 single-byte registers (~2 KB), standard error
 * 1.04/√m ≈ ±2.3 %. Hashing is xxh64 (ext-hash, always available on
 * PHP 8.1+) over the caller-supplied canonical string — the sketch
 * itself is encoding-agnostic; canonicalization rules live with the
 * producer (SPEC.md §4.4 for the KXSI "CHLL" section).
 *
 * The small-range (linear counting) correction is applied below
 * E ≤ 2.5m — that is what keeps estimates at cardinality 10 near-exact.
 * The classic LARGE-range correction (for E approaching 2^64) is
 * deliberately omitted: these sketches count values inside one
 * spreadsheet, bounded by its row count (< 2^31 per the format), which
 * is astronomically below the ~2^57 threshold where hash-collision bias
 * appears. Omitting it keeps the estimator branch-free and exactly
 * reproducible across implementations.
 *
 * Two HLLs with equal p merge losslessly by taking the register-wise
 * max — merge(a, b) is EXACTLY the sketch of the union stream (a
 * stronger property than the t-digest's statistical mergeability).
 *
 * Serialized layout (KXSI "CHLL" payload, little-endian, SPEC.md §4.4):
 *
 *     2   format version   uint16, = 1
 *     1   p                uint8, register-count exponent (m = 2^p)
 *     m   registers        raw bytes, register j at offset 3 + j
 */
class HyperLogLog
{
    public const FORMAT_VERSION = 1;

    public const DEFAULT_P = 11;

    /** Sane precision bounds: below 4 the estimator is meaningless, above 18 the payload (256 KB+) defeats the sidecar's size budget. */
    private const MIN_P = 4;

    private const MAX_P = 18;

    private int $p;

    private int $m;

    /** Register file as a mutable byte string — 1 byte per register, no PHP-array overhead. */
    private string $registers;

    public function __construct(int $p = self::DEFAULT_P)
    {
        if ($p < self::MIN_P || $p > self::MAX_P) {
            throw new \InvalidArgumentException(
                'HyperLogLog precision must be in ['.self::MIN_P.', '.self::MAX_P."]; got {$p}"
            );
        }
        $this->p = $p;
        $this->m = 1 << $p;
        $this->registers = str_repeat("\0", $this->m);
    }

    public function add(string $value): void
    {
        /** @var int $hash */
        $hash = unpack('J', hash('xxh64', $value, binary: true))[1];

        // Top p bits select the register; the remaining 64−p bits feed
        // the rank (position of the first 1-bit, 1-based). PHP's >> is
        // arithmetic, so mask after shifting to discard sign extension.
        $index = ($hash >> (64 - $this->p)) & ($this->m - 1);

        $w = $hash << $this->p;
        if ($w === 0) {
            $rank = 65 - $this->p; // all remaining bits zero
        } else {
            $rank = 1;
            while ($w > 0) { // MSB still 0 (and w non-zero) => one more leading zero
                $rank++;
                $w <<= 1;
            }
        }

        if (ord($this->registers[$index]) < $rank) {
            $this->registers[$index] = chr($rank);
        }
    }

    /**
     * Bulk add — identical to add() in a loop, minus the per-value
     * method-call overhead (the hash itself is ~20% of a single add()'s
     * cost; the rest is call/unpack plumbing this loop amortizes). Fed
     * by the writer's per-column row buffer.
     *
     * @param  list<string>  $values
     */
    public function addMany(array $values): void
    {
        $p = $this->p;
        $shift = 64 - $p;
        $mask = $this->m - 1;
        $zeroRank = 65 - $p;
        // Work on a local copy (one COW memcpy of the 2^p-byte string
        // per batch) so the per-value register reads skip the property
        // indirection; written back once below.
        $registers = $this->registers;

        foreach ($values as $value) {
            /** @var int $hash */
            $hash = unpack('J', hash('xxh64', $value, binary: true))[1];
            $index = ($hash >> $shift) & $mask;

            $w = $hash << $p;
            if ($w === 0) {
                $rank = $zeroRank;
            } else {
                $rank = 1;
                while ($w > 0) {
                    $rank++;
                    $w <<= 1;
                }
            }

            if (ord($registers[$index]) < $rank) {
                $registers[$index] = chr($rank);
            }
        }

        $this->registers = $registers;
    }

    /** Estimated number of distinct values added. */
    public function count(): int
    {
        $sum = 0.0;
        $zeros = 0;
        for ($j = 0; $j < $this->m; $j++) {
            $r = ord($this->registers[$j]);
            if ($r === 0) {
                $zeros++;
                $sum += 1.0;
            } else {
                $sum += 1.0 / (1 << $r);
            }
        }

        $alpha = 0.7213 / (1.0 + 1.079 / $this->m);
        $estimate = $alpha * $this->m * $this->m / $sum;

        // Small-range correction: linear counting on the empty-register
        // fraction is more accurate until ~2.5m distinct values.
        if ($estimate <= 2.5 * $this->m && $zeros > 0) {
            $estimate = $this->m * log($this->m / $zeros);
        }

        return (int) round($estimate);
    }

    /**
     * Register-wise max — after this call the sketch equals the one a
     * single pass over both streams would have produced. Precisions
     * must match; there is no sound way to fold registers across
     * different m.
     */
    public function merge(HyperLogLog $other): void
    {
        if ($other->p !== $this->p) {
            throw new \InvalidArgumentException(
                "cannot merge HyperLogLog sketches of different precision (p={$this->p} vs p={$other->p})"
            );
        }

        for ($j = 0; $j < $this->m; $j++) {
            if (ord($other->registers[$j]) > ord($this->registers[$j])) {
                $this->registers[$j] = $other->registers[$j];
            }
        }
    }

    public function precision(): int
    {
        return $this->p;
    }

    public function serialize(): string
    {
        return pack('v', self::FORMAT_VERSION).chr($this->p).$this->registers;
    }

    /**
     * Inverse of serialize(). Untrusted-input guards mirror TDigest:
     * exact length for the declared p, and every register within the
     * maximum rank a 64-bit hash can produce (64 − p + 1) — a larger
     * value would skew count() while looking structurally healthy.
     */
    public static function deserialize(string $payload): self
    {
        if (strlen($payload) < 3) {
            throw new \InvalidArgumentException('HyperLogLog payload too short');
        }

        $version = unpack('v', substr($payload, 0, 2))[1];
        if ($version !== self::FORMAT_VERSION) {
            throw new \InvalidArgumentException("unsupported HyperLogLog format version {$version}");
        }

        $p = ord($payload[2]);
        if ($p < self::MIN_P || $p > self::MAX_P) {
            throw new \InvalidArgumentException("HyperLogLog precision {$p} out of range");
        }

        $m = 1 << $p;
        if (strlen($payload) !== 3 + $m) {
            throw new \InvalidArgumentException(
                'HyperLogLog payload length '.strlen($payload)." does not match precision {$p} (expected ".(3 + $m).')'
            );
        }

        $registers = substr($payload, 3);
        $maxRank = 64 - $p + 1;
        for ($j = 0; $j < $m; $j++) {
            if (ord($registers[$j]) > $maxRank) {
                throw new \InvalidArgumentException("HyperLogLog register {$j} exceeds the maximum rank {$maxRank}");
            }
        }

        $hll = new self($p);
        $hll->registers = $registers;

        return $hll;
    }
}

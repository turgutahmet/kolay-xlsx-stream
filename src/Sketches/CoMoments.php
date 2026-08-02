<?php

namespace Kolay\XlsxStream\Sketches;

/**
 * Co-moment accumulators for exact Pearson correlation of one column pair
 * (KXSI "CORR" payload). Six running sums — n, Σx, Σy, Σxy, Σx², Σy² — are
 * everything Pearson's r needs, so a reader answers a pairwise correlation
 * from the sidecar alone with no row scan.
 *
 * Only rows where BOTH cells are numeric feed the pair (the producer's
 * job), so every accumulator counts the same n aligned observations. The
 * sums are plainly ADDITIVE, so the accumulator is mergeable by
 * construction: it composes across auto-split members, shards and future
 * stitched files exactly like the sketch family.
 *
 * Precision is the textbook single-pass formula's: r is computed from
 * n·Σxy − Σx·Σy over √((n·Σx² − (Σx)²)(n·Σy² − (Σy)²)). This is exact to
 * a rounding whisker on real-world magnitudes (the PoC measured < 1e-9),
 * but it does subtract large like-sized quantities, so pathologically huge
 * values with tiny variance can lose precision — the classic trade of the
 * moment form for O(1) mergeable state.
 */
class CoMoments
{
    /**
     * Serialized payload width: uint64 n + five little-endian doubles.
     * Fixed, so the CORR section frames pairs by count, not length.
     */
    public const PAYLOAD_BYTES = 8 + 5 * 8;

    private float $sumX = 0.0;

    private float $sumY = 0.0;

    private float $sumXY = 0.0;

    private float $sumX2 = 0.0;

    private float $sumY2 = 0.0;

    private int $n = 0;

    /** Fold one aligned observation (both cells numeric) into the sums. */
    public function add(float $x, float $y): void
    {
        $this->sumX += $x;
        $this->sumY += $y;
        $this->sumXY += $x * $y;
        $this->sumX2 += $x * $x;
        $this->sumY2 += $y * $y;
        $this->n++;
    }

    /** Absorb another accumulator's sums (additive merge). */
    public function merge(self $other): void
    {
        $this->sumX += $other->sumX;
        $this->sumY += $other->sumY;
        $this->sumXY += $other->sumXY;
        $this->sumX2 += $other->sumX2;
        $this->sumY2 += $other->sumY2;
        $this->n += $other->n;
    }

    public function n(): int
    {
        return $this->n;
    }

    /**
     * Pearson correlation coefficient, or null when it is undefined:
     * fewer than two observations, or either variable has zero variance
     * (a constant column — no correlation is defined). The result is
     * clamped to [-1, 1] so floating-point drift on a perfect relationship
     * cannot report a value fractionally outside the valid range.
     */
    public function pearson(): ?float
    {
        if ($this->n < 2) {
            return null;
        }

        $n = $this->n;
        $varX = $n * $this->sumX2 - $this->sumX ** 2;
        $varY = $n * $this->sumY2 - $this->sumY ** 2;
        if ($varX <= 0.0 || $varY <= 0.0) {
            return null;
        }

        $r = ($n * $this->sumXY - $this->sumX * $this->sumY) / sqrt($varX * $varY);

        return max(-1.0, min(1.0, $r));
    }

    /**
     * Fixed 48-byte payload: uint64 n then Σx, Σy, Σxy, Σx², Σy² as
     * little-endian doubles, in that order. Deterministic, so equal
     * accumulators serialize to equal bytes.
     */
    public function serialize(): string
    {
        return pack('P', $this->n)
            .pack('eeeee', $this->sumX, $this->sumY, $this->sumXY, $this->sumX2, $this->sumY2);
    }

    /**
     * Rebuild an accumulator from serialize()'s payload. Rejects a wrong
     * length or non-finite sums — a corrupt or crafted section must not
     * yield a NaN/Inf correlation downstream.
     */
    public static function deserialize(string $payload): self
    {
        if (\strlen($payload) !== self::PAYLOAD_BYTES) {
            throw new \InvalidArgumentException(
                'CoMoments payload must be '.self::PAYLOAD_BYTES.' bytes; got '.\strlen($payload)
            );
        }

        $fields = unpack('Pn/esumX/esumY/esumXY/esumX2/esumY2', $payload);
        if ($fields === false) {
            throw new \InvalidArgumentException('CoMoments payload failed to unpack');
        }
        foreach (['sumX', 'sumY', 'sumXY', 'sumX2', 'sumY2'] as $key) {
            if (! is_finite($fields[$key])) {
                throw new \InvalidArgumentException("CoMoments payload carries a non-finite {$key}");
            }
        }

        $acc = new self();
        $acc->n = $fields['n'];
        $acc->sumX = $fields['sumX'];
        $acc->sumY = $fields['sumY'];
        $acc->sumXY = $fields['sumXY'];
        $acc->sumX2 = $fields['sumX2'];
        $acc->sumY2 = $fields['sumY2'];

        return $acc;
    }
}

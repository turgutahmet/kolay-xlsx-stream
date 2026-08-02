<?php

namespace Kolay\XlsxStream\Sketches;

/**
 * Co-moment accumulators for exact Pearson correlation of one column pair
 * (KXSI "CORR" payload). Six values — n and the running MEANS and CENTRED
 * moments (mean_x, mean_y, M2_x = Σ(x−x̄)², M2_y, C_xy = Σ(x−x̄)(y−ȳ)) — are
 * everything Pearson's r needs, so a reader answers a pairwise correlation
 * from the sidecar alone with no row scan.
 *
 * Only rows where BOTH cells are numeric feed the pair (the producer's
 * job), so every moment counts the same n aligned observations.
 *
 * The centred (Welford) form is used deliberately, NOT the textbook
 * n·Σxy − Σx·Σy sums. That sum form subtracts two large like-sized
 * quantities and loses precision exactly on the most ordinary export
 * column — a timestamp or date serial, whose values are huge relative to
 * their spread (measured: r wrong from the third digit on a one-hour epoch
 * window). Welford never forms those sums, so it stays accurate there, and
 * it is STILL mergeable by construction via Chan's parallel algorithm — the
 * merge of two centred states is exact algebra on (n, means, moments), so
 * it composes across auto-split members, shards and stitched files like the
 * sketch family. The payload is unchanged in shape (uint64 n + five
 * doubles), so this stays a fixed 48-byte accumulator.
 */
class CoMoments
{
    /**
     * Serialized payload width: uint64 n + five little-endian doubles.
     * Fixed, so the CORR section frames pairs by count, not length.
     */
    public const PAYLOAD_BYTES = 8 + 5 * 8;

    private float $meanX = 0.0;

    private float $meanY = 0.0;

    private float $m2X = 0.0;

    private float $m2Y = 0.0;

    private float $cXY = 0.0;

    private int $n = 0;

    /**
     * Fold one aligned observation (both cells numeric) with the Welford
     * update: the means move by 1/n of the residual, and the centred
     * moments accumulate against the pre- and post-update means.
     */
    public function add(float $x, float $y): void
    {
        $this->n++;
        $dx = $x - $this->meanX;
        $dy = $y - $this->meanY;
        $this->meanX += $dx / $this->n;
        $this->meanY += $dy / $this->n;
        $this->m2X += $dx * ($x - $this->meanX);
        $this->m2Y += $dy * ($y - $this->meanY);
        $this->cXY += $dx * ($y - $this->meanY);
    }

    /**
     * Absorb another accumulator via Chan's parallel merge — exact
     * combination of two centred states, the property that keeps the
     * Welford form mergeable across shards / chains.
     */
    public function merge(self $other): void
    {
        if ($other->n === 0) {
            return;
        }
        if ($this->n === 0) {
            $this->n = $other->n;
            $this->meanX = $other->meanX;
            $this->meanY = $other->meanY;
            $this->m2X = $other->m2X;
            $this->m2Y = $other->m2Y;
            $this->cXY = $other->cXY;

            return;
        }

        $na = $this->n;
        $nb = $other->n;
        $n = $na + $nb;
        $dX = $other->meanX - $this->meanX;
        $dY = $other->meanY - $this->meanY;
        $scale = ($na * $nb) / $n;

        $this->m2X += $other->m2X + $dX * $dX * $scale;
        $this->m2Y += $other->m2Y + $dY * $dY * $scale;
        $this->cXY += $other->cXY + $dX * $dY * $scale;
        $this->meanX += $dX * $nb / $n;
        $this->meanY += $dY * $nb / $n;
        $this->n = $n;
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
        if ($this->n < 2 || $this->m2X <= 0.0 || $this->m2Y <= 0.0) {
            return null;
        }

        $r = $this->cXY / sqrt($this->m2X * $this->m2Y);

        return max(-1.0, min(1.0, $r));
    }

    /**
     * Fixed 48-byte payload: uint64 n then mean_x, mean_y, M2_x, M2_y, C_xy
     * as little-endian doubles, in that order. Deterministic, so equal
     * accumulators serialize to equal bytes.
     */
    public function serialize(): string
    {
        return pack('P', $this->n)
            .pack('eeeee', $this->meanX, $this->meanY, $this->m2X, $this->m2Y, $this->cXY);
    }

    /**
     * Rebuild an accumulator from serialize()'s payload. Rejects a wrong
     * length or non-finite moments — a corrupt or crafted section must not
     * yield a NaN/Inf correlation downstream.
     */
    public static function deserialize(string $payload): self
    {
        if (\strlen($payload) !== self::PAYLOAD_BYTES) {
            throw new \InvalidArgumentException(
                'CoMoments payload must be '.self::PAYLOAD_BYTES.' bytes; got '.\strlen($payload)
            );
        }

        $fields = unpack('Pn/emeanX/emeanY/em2X/em2Y/ecXY', $payload);
        if ($fields === false) {
            throw new \InvalidArgumentException('CoMoments payload failed to unpack');
        }
        foreach (['meanX', 'meanY', 'm2X', 'm2Y', 'cXY'] as $key) {
            if (! is_finite($fields[$key])) {
                throw new \InvalidArgumentException("CoMoments payload carries a non-finite {$key}");
            }
        }

        $acc = new self();
        $acc->n = $fields['n'];
        $acc->meanX = $fields['meanX'];
        $acc->meanY = $fields['meanY'];
        $acc->m2X = $fields['m2X'];
        $acc->m2Y = $fields['m2Y'];
        $acc->cXY = $fields['cXY'];

        return $acc;
    }
}

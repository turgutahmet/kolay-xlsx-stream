<?php

namespace Kolay\XlsxStream\Sketches;

/**
 * Merging t-digest (Dunning) — a fixed-size sketch that answers
 * approximate quantiles over a numeric stream of any length.
 *
 * Design points, matching the reference MergingDigest:
 *
 *   - Values are buffered raw and folded into the centroid set in
 *     batches (amortized O(1) per add): the buffer plus the existing
 *     centroids are sorted and re-clustered in one merge pass.
 *   - Cluster sizes are bounded by the k1 scale function
 *     k(q) = δ/(2π)·asin(2q−1), which keeps clusters near q=0 and q=1
 *     tiny (tail quantiles are the accurate ones — rank error is
 *     proportional to q(1−q)) while allowing wide clusters mid-range.
 *     δ (compression) ≈ the number of centroids retained; δ=100 holds
 *     the serialized form around 1–4 KB.
 *   - Digests merge associatively: merging the digests of two halves
 *     of a stream yields (statistically) the digest of the whole — the
 *     property that lets per-sheet/per-segment sketches be combined
 *     without touching row data.
 *
 * quantile() interpolates linearly between centroid means positioned
 * at their cumulative-weight midpoints, anchored at the exact min/max:
 * q=0 and q=1 return the true extremes, and every estimate is clamped
 * inside [min, max] by construction.
 *
 * Serialized layout (KXSI "TDIG" payload, little-endian, SPEC.md §4.3):
 *
 *     2   format version   uint16, = 1
 *     2   compression δ    uint16
 *     4   centroid_count   uint32
 *     16×C centroids       per centroid: float64 mean, float64 weight
 *     8   min              float64 — meaningless when count == 0
 *     8   max              float64 — meaningless when count == 0
 *     8   count            uint64, total values added
 *
 * The instance is NOT thread-safe; non-finite values are rejected
 * loudly (a NAN mean would silently poison every later estimate).
 */
class TDigest
{
    public const FORMAT_VERSION = 1;

    public const DEFAULT_COMPRESSION = 100;

    /**
     * Buffered raw values folded in per merge pass. Larger buffers
     * amortize the O(n log n) sort over more adds; 512 at δ=100 keeps
     * the working set trivial while measuring ~amortized-constant add
     * cost.
     */
    private const BUFFER_LIMIT = 512;

    /** Serialized: 8-byte fixed head + 16 per centroid + 24-byte tail. */
    private const HEAD_BYTES = 8;

    private const TAIL_BYTES = 24;

    private const CENTROID_BYTES = 16;

    private int $compression;

    /** @var list<float> centroid means, ascending */
    private array $means = [];

    /** @var list<float> centroid weights, parallel to $means */
    private array $weights = [];

    /** @var list<float> raw values not yet merged into centroids */
    private array $buffer = [];

    private float $min = 0.0;

    private float $max = 0.0;

    private int $count = 0;

    public function __construct(int $compression = self::DEFAULT_COMPRESSION)
    {
        if ($compression < 1 || $compression > 0xFFFF) {
            throw new \InvalidArgumentException(
                "t-digest compression must be in [1, 65535]; got {$compression}"
            );
        }
        $this->compression = $compression;
    }

    public function add(float $value): void
    {
        if (! is_finite($value)) {
            throw new \InvalidArgumentException(
                't-digest cannot absorb a non-finite value — filter NAN/INF before add()'
            );
        }

        if ($this->count === 0) {
            $this->min = $value;
            $this->max = $value;
        } else {
            if ($value < $this->min) {
                $this->min = $value;
            }
            if ($value > $this->max) {
                $this->max = $value;
            }
        }
        $this->count++;

        $this->buffer[] = $value;
        if (count($this->buffer) >= self::BUFFER_LIMIT) {
            $this->flushBuffer();
        }
    }

    /**
     * Bulk add — one call absorbs a whole batch. Semantically identical
     * to add() in a loop, but the min/max/finite scan runs in a single
     * tight pass and the per-value method-call overhead disappears
     * (measured ~2x on the writer's accumulation path, where this is
     * fed from a per-column row buffer). Values MUST be finite floats;
     * the batch is validated before any state changes, so a rejected
     * batch leaves the digest untouched.
     *
     * @param  list<float>  $values
     */
    public function addMany(array $values): void
    {
        if ($values === []) {
            return;
        }

        // C-speed validation: a NAN/INF anywhere in the batch propagates
        // into array_sum. Only that (vanishingly rare) case — which
        // includes a legitimate overflow of huge finite values — pays
        // for a per-value scan to tell the two apart.
        if (! is_finite(array_sum($values))) {
            foreach ($values as $v) {
                if (! is_finite($v)) {
                    throw new \InvalidArgumentException(
                        't-digest cannot absorb a non-finite value — filter NAN/INF before addMany()'
                    );
                }
            }
        }

        $min = (float) min($values);
        $max = (float) max($values);

        if ($this->count === 0) {
            $this->min = $min;
            $this->max = $max;
        } else {
            if ($min < $this->min) {
                $this->min = $min;
            }
            if ($max > $this->max) {
                $this->max = $max;
            }
        }
        $this->count += count($values);

        $this->buffer = $this->buffer === []
            ? array_values($values)
            : array_merge($this->buffer, $values);
        if (count($this->buffer) >= self::BUFFER_LIMIT) {
            $this->flushBuffer();
        }
    }

    /**
     * Fold another digest into this one (associative up to statistical
     * noise — the stitch primitive for per-segment sketches). The other
     * digest is not modified beyond flushing its internal buffer; this
     * digest keeps its own compression setting.
     */
    public function merge(TDigest $other): void
    {
        $other->flushBuffer();
        if ($other->count === 0) {
            return;
        }

        $this->flushBuffer();

        if ($this->count === 0) {
            $this->min = $other->min;
            $this->max = $other->max;
        } else {
            $this->min = min($this->min, $other->min);
            $this->max = max($this->max, $other->max);
        }
        $this->count += $other->count;

        $means = array_merge($this->means, $other->means);
        $weights = array_merge($this->weights, $other->weights);
        array_multisort($means, SORT_ASC, SORT_NUMERIC, $weights);

        $this->mergeCentroids($means, $weights);
    }

    /**
     * Approximate value at quantile $q ∈ [0, 1]; null when the digest
     * holds no values. q=0 and q=1 return the exact min/max; single-
     * value and all-equal streams are answered exactly.
     */
    public function quantile(float $q): ?float
    {
        if ($q < 0.0 || $q > 1.0) {
            throw new \InvalidArgumentException("quantile must be within [0, 1]; got {$q}");
        }

        $this->flushBuffer();

        if ($this->count === 0) {
            return null;
        }
        if ($q === 0.0 || $this->min === $this->max) {
            return $this->min;
        }
        if ($q === 1.0) {
            return $this->max;
        }

        // Centroid i sits at its cumulative-weight midpoint
        // m_i = Σ_{j<i} w_j + w_i / 2. Interpolate linearly between
        // neighbouring midpoints; the head segment is anchored at
        // (0, min) and the tail segment at (count, max), so estimates
        // can never leave [min, max].
        $index = $q * $this->count;

        $cumBefore = 0.0;
        $prevMid = 0.0;
        $prevValue = $this->min;

        foreach ($this->means as $i => $mean) {
            $mid = $cumBefore + $this->weights[$i] / 2;
            if ($index < $mid) {
                return $prevValue + ($mean - $prevValue) * ($index - $prevMid) / ($mid - $prevMid);
            }
            $cumBefore += $this->weights[$i];
            $prevMid = $mid;
            $prevValue = $mean;
        }

        // Past the last midpoint: tail segment toward the exact max.
        // count − prevMid = w_last / 2 > 0, so the division is safe.
        return $prevValue + ($this->max - $prevValue) * ($index - $prevMid) / ($this->count - $prevMid);
    }

    /**
     * Approximate CDF: fraction of the stream's values <= $x — the inverse
     * of quantile(). null for an empty digest; 0.0 below min, 1.0 above
     * max. A single O(centroids) pass through the same piecewise-linear
     * model quantile() uses, so rank() and quantile() are numerical
     * inverses within the digest's resolution:
     *
     *     rank(quantile(q)) ≈ q    and    quantile(rank(x)) ≈ x
     *
     * This is the direct CDF the reader's selectivity estimator needs;
     * previously it recovered the CDF by bisecting quantile() 40 times
     * (40 × O(centroids) → one O(centroids) pass here, ~80× fewer calls
     * into the centroid walk for a 'between' estimate).
     */
    public function rank(float $x): ?float
    {
        $this->flushBuffer();

        if ($this->count === 0) {
            return null;
        }
        // Boundary conventions mirror quantile()'s early returns: x at/below
        // min → 0.0 (the rank of the first value), x at/above max → 1.0.
        if ($x <= $this->min) {
            return 0.0;
        }
        if ($x >= $this->max) {
            return 1.0;
        }

        // Walk centroid midpoints exactly as quantile() does, but solve for
        // the rank whose quantile would land at x. Centroid i sits at the
        // cumulative-weight midpoint m_i; the value segment between the
        // previous anchor and mean_i maps linearly onto the rank segment
        // [prevMid, m_i]. Head anchor (min, 0), tail (max, count).
        //
        // Whenever $x < $mean triggers we have prevValue <= x < mean, so
        // $mean > $prevValue strictly — the interpolation denominator is
        // always positive (no degenerate-segment guard needed).
        $cumBefore = 0.0;
        $prevMid = 0.0;
        $prevValue = $this->min;

        foreach ($this->means as $i => $mean) {
            $mid = $cumBefore + $this->weights[$i] / 2;
            if ($x < $mean) {
                $index = $prevMid + ($mid - $prevMid) * ($x - $prevValue) / ($mean - $prevValue);

                return $index / $this->count;
            }
            $cumBefore += $this->weights[$i];
            $prevMid = $mid;
            $prevValue = $mean;
        }

        // Tail segment [prevValue, max] → rank [prevMid, count]. We reach
        // here only when x >= every centroid mean; x < max (guarded above),
        // and max > prevValue (the last centroid does not saturate at max
        // unless min === max, which the early return handled).
        $index = $prevMid + ($this->count - $prevMid) * ($x - $prevValue) / ($this->max - $prevValue);

        return $index / $this->count;
    }

    public function count(): int
    {
        return $this->count;
    }

    public function min(): ?float
    {
        return $this->count > 0 ? $this->min : null;
    }

    public function max(): ?float
    {
        return $this->count > 0 ? $this->max : null;
    }

    public function compression(): int
    {
        return $this->compression;
    }

    /** Number of retained centroids (after folding the pending buffer). */
    public function centroidCount(): int
    {
        $this->flushBuffer();

        return count($this->means);
    }

    public function serialize(): string
    {
        $this->flushBuffer();

        $out = pack('vv', self::FORMAT_VERSION, $this->compression);
        $out .= pack('V', count($this->means));
        foreach ($this->means as $i => $mean) {
            $out .= pack('ee', $mean, $this->weights[$i]);
        }
        $out .= pack('ee', $this->min, $this->max);
        $out .= pack('P', $this->count);

        return $out;
    }

    /**
     * Inverse of serialize(). The payload is untrusted input (it rides
     * in a sidecar an attacker can rewrite with a valid CRC), so every
     * structural property the estimator relies on is verified before an
     * instance exists: exact length, positive finite weights, finite
     * ascending means, and weight/count agreement.
     */
    public static function deserialize(string $payload): self
    {
        $len = strlen($payload);
        if ($len < self::HEAD_BYTES + self::TAIL_BYTES) {
            throw new \InvalidArgumentException('t-digest payload too short');
        }

        $head = unpack('vversion/vcompression/Vcentroids', substr($payload, 0, self::HEAD_BYTES));
        if ($head['version'] !== self::FORMAT_VERSION) {
            throw new \InvalidArgumentException(
                "unsupported t-digest format version {$head['version']}"
            );
        }
        if ($head['compression'] < 1) {
            throw new \InvalidArgumentException('t-digest compression must be >= 1');
        }

        $centroidCount = $head['centroids'];
        $expected = self::HEAD_BYTES + self::CENTROID_BYTES * $centroidCount + self::TAIL_BYTES;
        if ($len !== $expected) {
            throw new \InvalidArgumentException(
                "t-digest payload length {$len} does not match centroid count {$centroidCount} (expected {$expected})"
            );
        }

        $digest = new self($head['compression']);

        $cursor = self::HEAD_BYTES;
        $totalWeight = 0.0;
        $prevMean = -INF;
        for ($i = 0; $i < $centroidCount; $i++) {
            $c = unpack('emean/eweight', substr($payload, $cursor, self::CENTROID_BYTES));
            $cursor += self::CENTROID_BYTES;

            if (! is_finite($c['mean']) || ! is_finite($c['weight']) || $c['weight'] <= 0.0) {
                throw new \InvalidArgumentException('t-digest centroid mean/weight out of range');
            }
            if ($c['mean'] < $prevMean) {
                throw new \InvalidArgumentException('t-digest centroid means not ascending');
            }
            $prevMean = $c['mean'];

            $digest->means[] = $c['mean'];
            $digest->weights[] = $c['weight'];
            $totalWeight += $c['weight'];
        }

        $tail = unpack('emin/emax/Pcount', substr($payload, $cursor, self::TAIL_BYTES));
        $digest->min = $tail['min'];
        $digest->max = $tail['max'];
        $digest->count = $tail['count'];

        if ($digest->count < 0) {
            throw new \InvalidArgumentException('t-digest count exceeds the supported integer range');
        }
        // Weights are integral by construction (each add contributes 1;
        // merges sum integers, exact in float64 far past any realistic
        // row count) — so the centroid mass must reproduce the count.
        if (abs($totalWeight - $digest->count) > 1e-6 * max(1.0, (float) $digest->count)) {
            throw new \InvalidArgumentException('t-digest centroid weights do not sum to the recorded count');
        }
        if ($digest->count > 0) {
            if (! is_finite($digest->min) || ! is_finite($digest->max) || $digest->min > $digest->max) {
                throw new \InvalidArgumentException('t-digest min/max out of range');
            }
        }

        return $digest;
    }

    /**
     * Sort the pending raw values, weave them into the centroid list
     * (both already sorted — a linear two-pointer pass), and re-cluster.
     */
    private function flushBuffer(): void
    {
        if ($this->buffer === []) {
            return;
        }

        sort($this->buffer);

        $means = [];
        $weights = [];
        $b = 0;
        $bn = count($this->buffer);
        $c = 0;
        $cn = count($this->means);
        while ($b < $bn || $c < $cn) {
            if ($c >= $cn || ($b < $bn && $this->buffer[$b] <= $this->means[$c])) {
                $means[] = $this->buffer[$b];
                $weights[] = 1.0;
                $b++;
            } else {
                $means[] = $this->means[$c];
                $weights[] = $this->weights[$c];
                $c++;
            }
        }

        $this->buffer = [];
        $this->mergeCentroids($means, $weights);
    }

    /**
     * One merge pass over sorted (mean, weight) pairs: greedily grow a
     * cluster while its cumulative weight stays under the k1 bound for
     * its quantile position, emit it when the next item would burst it.
     *
     * @param  list<float>  $means  ascending
     * @param  list<float>  $weights  parallel, all > 0
     */
    private function mergeCentroids(array $means, array $weights): void
    {
        $n = count($means);
        if ($n === 0) {
            $this->means = [];
            $this->weights = [];

            return;
        }

        $total = 0.0;
        foreach ($weights as $w) {
            $total += $w;
        }

        $outMeans = [];
        $outWeights = [];
        $sigmaMean = $means[0];
        $sigmaWeight = $weights[0];
        $weightSoFar = 0.0;
        $limit = $total * $this->qLimit(0.0);

        for ($i = 1; $i < $n; $i++) {
            if ($weightSoFar + $sigmaWeight + $weights[$i] <= $limit) {
                // Absorb into the open cluster (incremental mean keeps
                // the value inside [sigmaMean, means[i]]).
                $sigmaWeight += $weights[$i];
                $sigmaMean += ($means[$i] - $sigmaMean) * $weights[$i] / $sigmaWeight;
            } else {
                $outMeans[] = $sigmaMean;
                $outWeights[] = $sigmaWeight;
                $weightSoFar += $sigmaWeight;
                $limit = $total * $this->qLimit($weightSoFar / $total);
                $sigmaMean = $means[$i];
                $sigmaWeight = $weights[$i];
            }
        }
        $outMeans[] = $sigmaMean;
        $outWeights[] = $sigmaWeight;

        $this->means = $outMeans;
        $this->weights = $outWeights;
    }

    /**
     * Upper quantile bound for a cluster that starts at quantile $q:
     * k1⁻¹(k1(q) + 1) with k1(q) = δ/(2π)·asin(2q−1). Near the tails
     * one k-unit spans a sliver of q (small clusters, tight tails);
     * mid-range it spans much more (wide clusters, bounded total).
     */
    private function qLimit(float $q): float
    {
        $k = ($this->compression / (2.0 * M_PI)) * asin(2.0 * $q - 1.0) + 1.0;
        $x = 2.0 * M_PI * $k / $this->compression;
        if ($x >= M_PI / 2.0) {
            return 1.0;
        }

        return (sin($x) + 1.0) / 2.0;
    }
}

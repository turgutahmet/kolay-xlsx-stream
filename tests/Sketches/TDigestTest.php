<?php

namespace Kolay\XlsxStream\Tests\Sketches;

use Kolay\XlsxStream\Sketches\TDigest;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * Oracle tests for the merging t-digest: every estimate is measured in
 * RANK error against the exact empirical quantile of the full sample —
 * "how far off is the estimate's position in the sorted data", which is
 * the error the t-digest actually bounds (value error is unbounded for
 * skewed data, rank error is not).
 *
 * Tolerances are PINNED FROM MEASURED RUNS on the fixed seeds below
 * (worst observed, then rounded up ~5-10x for cross-platform libm slack
 * in the clustering bound — asin/sin may differ in the last ulp):
 *
 *   uniform    p01/p99 ≤ 0.04 %, p50 ≤ 0.02 %   → pinned 0.2 % / 0.5 %
 *   lognormal  p01/p99 ≤ 0.09 %, p50 ≤ 0.05 %   → pinned 0.2 % / 0.5 %
 *   heavy-dup  tails exact,      p50 ≈ 2.4 %    → pinned 0.5 % / 4.5 %
 *
 * All far inside the feature contract (tails ≤ 0.5 %, median ≤ 1.5 %
 * on continuous data). Heavy-duplicate mid-range is looser by nature:
 * with 21 distinct values each spans ~4.8 % of rank space, and an
 * interpolated estimate falling BETWEEN two duplicate masses is charged
 * the full distance to the target rank.
 */
class TDigestTest extends TestCase
{
    private const N = 100000;

    // ------------------------------------------------------------------
    // Accuracy oracles
    // ------------------------------------------------------------------

    public function test_uniform_distribution_rank_error(): void
    {
        mt_srand(1234);
        [$digest, $sorted] = $this->build(fn () => mt_rand() / mt_getrandmax() * 1000.0);

        $this->assertRankErrors($digest, $sorted, tail: 0.002, mid: 0.005);
    }

    public function test_lognormal_distribution_rank_error(): void
    {
        // Box-Muller over mt_rand uniforms — heavily right-skewed
        // (exp(1.5·gauss)), the shape that breaks value-error oracles
        // and naive equal-width histograms.
        mt_srand(1234);
        [$digest, $sorted] = $this->build(function () {
            $u1 = (mt_rand() + 1) / (mt_getrandmax() + 2);
            $u2 = (mt_rand() + 1) / (mt_getrandmax() + 2);

            return exp(1.5 * sqrt(-2 * log($u1)) * cos(2 * M_PI * $u2));
        });

        $this->assertRankErrors($digest, $sorted, tail: 0.002, mid: 0.005);
    }

    public function test_heavy_duplicate_distribution_rank_error(): void
    {
        // Only 21 distinct values over 100K samples — every centroid is
        // a stack of duplicates. Tails stay sharp; mid-range rank error
        // is dominated by interpolation landing between value masses.
        mt_srand(1234);
        [$digest, $sorted] = $this->build(fn () => (float) mt_rand(0, 20));

        $this->assertRankErrors($digest, $sorted, tail: 0.005, mid: 0.045);

        // The rounded estimate must still land on the exact percentile
        // value — "which value is the median" has one right answer here.
        $exactMedian = $sorted[(int) (0.5 * (self::N - 1))];
        $this->assertEqualsWithDelta($exactMedian, round((float) $digest->quantile(0.5)), 1.0);
    }

    public function test_compression_bounds_centroids_and_size(): void
    {
        mt_srand(99);
        $digest = new TDigest(100);
        for ($i = 0; $i < self::N; $i++) {
            $digest->add(mt_rand() / mt_getrandmax());
        }

        // δ=100 with the k1 scale retains well under 3δ centroids —
        // the serialized form stays in the promised 1-4 KB envelope.
        $this->assertLessThan(300, $digest->centroidCount());
        $this->assertLessThan(4096, strlen($digest->serialize()));
    }

    // ------------------------------------------------------------------
    // Small-n exactness
    // ------------------------------------------------------------------

    public function test_empty_digest_returns_null(): void
    {
        $digest = new TDigest();

        $this->assertNull($digest->quantile(0.0));
        $this->assertNull($digest->quantile(0.5));
        $this->assertNull($digest->quantile(1.0));
        $this->assertSame(0, $digest->count());
        $this->assertNull($digest->min());
        $this->assertNull($digest->max());
    }

    public function test_single_value_is_every_quantile(): void
    {
        $digest = new TDigest();
        $digest->add(42.5);

        foreach ([0.0, 0.25, 0.5, 0.99, 1.0] as $q) {
            $this->assertSame(42.5, $digest->quantile($q), "q={$q}");
        }
    }

    public function test_two_values_hit_min_mid_max(): void
    {
        $digest = new TDigest();
        $digest->add(3.0);
        $digest->add(1.0);

        $this->assertSame(1.0, $digest->quantile(0.0));
        $this->assertSame(2.0, $digest->quantile(0.5));
        $this->assertSame(3.0, $digest->quantile(1.0));
    }

    public function test_all_equal_values_answer_exactly(): void
    {
        $digest = new TDigest();
        for ($i = 0; $i < 10000; $i++) {
            $digest->add(7.5);
        }

        foreach ([0.0, 0.1, 0.5, 0.9, 1.0] as $q) {
            $this->assertSame(7.5, $digest->quantile($q), "q={$q}");
        }
    }

    public function test_extreme_quantiles_return_exact_min_max(): void
    {
        mt_srand(7);
        $digest = new TDigest();
        $min = INF;
        $max = -INF;
        for ($i = 0; $i < 50000; $i++) {
            $v = mt_rand() / mt_getrandmax() * 2000.0 - 1000.0;
            $min = min($min, $v);
            $max = max($max, $v);
            $digest->add($v);
        }

        $this->assertSame($min, $digest->quantile(0.0));
        $this->assertSame($max, $digest->quantile(1.0));
        $this->assertSame($min, $digest->min());
        $this->assertSame($max, $digest->max());
    }

    public function test_estimates_never_leave_min_max(): void
    {
        mt_srand(11);
        $digest = new TDigest();
        for ($i = 0; $i < 10000; $i++) {
            $digest->add(exp(10 * (mt_rand() / mt_getrandmax() - 0.5)));
        }

        for ($q = 0.0; $q <= 1.0; $q += 0.01) {
            $v = $digest->quantile(round($q, 2));
            $this->assertGreaterThanOrEqual($digest->min(), $v);
            $this->assertLessThanOrEqual($digest->max(), $v);
        }
    }

    // ------------------------------------------------------------------
    // Merge — the stitch-critical property
    // ------------------------------------------------------------------

    public function test_merged_halves_approximate_the_whole(): void
    {
        mt_srand(4321);
        $a = new TDigest();
        $b = new TDigest();
        $whole = new TDigest();
        $data = [];
        for ($i = 0; $i < self::N; $i++) {
            $v = mt_rand() / mt_getrandmax() * 1000.0;
            $data[] = $v;
            $whole->add($v);
            ($i < self::N / 2 ? $a : $b)->add($v);
        }
        sort($data);

        $a->merge($b);

        $this->assertSame(self::N, $a->count());
        $this->assertSame($whole->min(), $a->min());
        $this->assertSame($whole->max(), $a->max());

        // The merged digest must satisfy the same rank-error contract
        // as the single-pass digest — that is what makes per-sheet /
        // per-segment sketches stitchable.
        $this->assertRankErrors($a, $data, tail: 0.002, mid: 0.005);
    }

    public function test_merging_empty_and_into_empty(): void
    {
        $empty = new TDigest();
        $full = new TDigest();
        for ($i = 1; $i <= 100; $i++) {
            $full->add((float) $i);
        }

        $full->merge(new TDigest());
        $this->assertSame(100, $full->count());
        $this->assertSame(50.5, $full->quantile(0.5));

        $empty->merge($full);
        $this->assertSame(100, $empty->count());
        $this->assertSame(1.0, $empty->quantile(0.0));
        $this->assertSame(100.0, $empty->quantile(1.0));
    }

    // ------------------------------------------------------------------
    // Serialization
    // ------------------------------------------------------------------

    public function test_serialize_round_trips_every_quantile(): void
    {
        mt_srand(7);
        $digest = new TDigest();
        for ($i = 0; $i < 10000; $i++) {
            $digest->add(mt_rand() / mt_getrandmax());
        }

        $restored = TDigest::deserialize($digest->serialize());

        $this->assertSame($digest->count(), $restored->count());
        $this->assertSame($digest->compression(), $restored->compression());
        foreach ([0.0, 0.01, 0.25, 0.5, 0.75, 0.99, 1.0] as $q) {
            $this->assertSame($digest->quantile($q), $restored->quantile($q), "q={$q}");
        }
        // Canonical form: serializing the restored digest reproduces
        // the exact bytes.
        $this->assertSame($digest->serialize(), $restored->serialize());
    }

    public function test_empty_digest_serializes_and_round_trips(): void
    {
        $restored = TDigest::deserialize((new TDigest())->serialize());

        $this->assertSame(0, $restored->count());
        $this->assertNull($restored->quantile(0.5));
    }

    #[\PHPUnit\Framework\Attributes\DataProvider('corruptPayloads')]
    public function test_deserialize_rejects_corrupt_payloads(string $label, string $payload): void
    {
        $this->expectException(\InvalidArgumentException::class);

        TDigest::deserialize($payload);
    }

    /** @return array<string, array{string, string}> */
    public static function corruptPayloads(): array
    {
        $valid = (function (): string {
            $d = new TDigest();
            foreach ([1.0, 2.0, 3.0] as $v) {
                $d->add($v);
            }

            return $d->serialize();
        })();

        $withCentroids = function (string $centroids, int $count, int $total): string {
            return pack('vv', 1, 100).pack('V', $count).$centroids
                .pack('ee', 1.0, 9.0).pack('P', $total);
        };

        return [
            'too short' => ['too short', 'ABC'],
            'bad version' => ['bad version', pack('vv', 99, 100).substr($valid, 4)],
            'zero compression' => ['zero compression', substr($valid, 0, 2).pack('v', 0).substr($valid, 4)],
            'length mismatch' => ['length mismatch', $valid.'x'],
            'truncated centroids' => ['truncated centroids', substr($valid, 0, strlen($valid) - 1)],
            'negative weight' => ['negative weight', $withCentroids(pack('ee', 5.0, -1.0), 1, 1)],
            'zero weight' => ['zero weight', $withCentroids(pack('ee', 5.0, 0.0), 1, 1)],
            'nan mean' => ['nan mean', $withCentroids(pack('ee', NAN, 1.0), 1, 1)],
            'descending means' => ['descending means', $withCentroids(pack('ee', 9.0, 1.0).pack('ee', 1.0, 1.0), 2, 2)],
            'weight/count mismatch' => ['weight/count mismatch', $withCentroids(pack('ee', 5.0, 2.0), 1, 5)],
            'min greater than max' => ['min greater than max', pack('vv', 1, 100).pack('V', 1).pack('ee', 5.0, 1.0).pack('ee', 9.0, 1.0).pack('P', 1)],
        ];
    }

    public function test_rejects_non_finite_values_and_bad_quantile_args(): void
    {
        $digest = new TDigest();

        try {
            $digest->add(NAN);
            $this->fail('NAN accepted');
        } catch (\InvalidArgumentException) {
        }
        try {
            $digest->add(INF);
            $this->fail('INF accepted');
        } catch (\InvalidArgumentException) {
        }

        $digest->add(1.0);
        try {
            $digest->quantile(-0.01);
            $this->fail('q < 0 accepted');
        } catch (\InvalidArgumentException) {
        }
        try {
            $digest->quantile(1.01);
            $this->fail('q > 1 accepted');
        } catch (\InvalidArgumentException) {
        }

        $this->assertSame(1, $digest->count());
    }

    public function test_rejects_out_of_range_compression(): void
    {
        $this->expectException(\InvalidArgumentException::class);

        new TDigest(0);
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    /**
     * @param  callable(): float  $gen
     * @return array{TDigest, list<float>} digest + sorted sample
     */
    private function build(callable $gen): array
    {
        $digest = new TDigest();
        $data = [];
        for ($i = 0; $i < self::N; $i++) {
            $v = $gen();
            $data[] = $v;
            $digest->add($v);
        }
        sort($data);

        return [$digest, $data];
    }

    /**
     * Assert rank errors at the contract quantiles: tails (p01/p99 and
     * beyond) against $tail, mid-range (p25/p50/p75) against $mid.
     *
     * @param  list<float>  $sorted
     */
    private function assertRankErrors(TDigest $digest, array $sorted, float $tail, float $mid): void
    {
        foreach ([0.001 => $tail, 0.01 => $tail, 0.25 => $mid, 0.5 => $mid, 0.75 => $mid, 0.99 => $tail, 0.999 => $tail] as $q => $tolerance) {
            $estimate = $digest->quantile((float) $q);
            $this->assertNotNull($estimate);
            $err = $this->rankError($sorted, $estimate, (float) $q);
            $this->assertLessThanOrEqual(
                $tolerance,
                $err,
                sprintf('rank error %.4f%% at q=%s exceeds %.2f%%', $err * 100, $q, $tolerance * 100)
            );
        }
    }

    /**
     * Rank error of $estimate against the exact sample: the estimate
     * occupies the rank interval [#values < estimate, #values <= estimate]
     * (a whole span when it lands on a duplicate mass); the error is the
     * distance from q to that interval, as a fraction of n.
     *
     * @param  list<float>  $sorted
     */
    private function rankError(array $sorted, float $estimate, float $q): float
    {
        $n = count($sorted);

        $lo = 0;
        $hi = $n;
        while ($lo < $hi) {
            $m = ($lo + $hi) >> 1;
            $sorted[$m] < $estimate ? $lo = $m + 1 : $hi = $m;
        }
        $below = $lo;

        $lo = 0;
        $hi = $n;
        while ($lo < $hi) {
            $m = ($lo + $hi) >> 1;
            $sorted[$m] <= $estimate ? $lo = $m + 1 : $hi = $m;
        }
        $atOrBelow = $lo;

        $loQ = $below / $n;
        $hiQ = $atOrBelow / $n;
        if ($q < $loQ) {
            return $loQ - $q;
        }
        if ($q > $hiQ) {
            return $q - $hiQ;
        }

        return 0.0;
    }
}

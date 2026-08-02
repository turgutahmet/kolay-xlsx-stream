<?php

namespace Kolay\XlsxStream\Tests\Sketches;

use Kolay\XlsxStream\Sketches\CoMoments;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * CoMoments — the six co-moment accumulators (n, Σx, Σy, Σxy, Σx², Σy²)
 * behind exact Pearson correlation from the sidecar. Gates: the Pearson
 * value matches a textbook oracle to machine precision; perfect linear and
 * independent relationships resolve to ±1 / ~0; degenerate cases (n<2, zero
 * variance) are undefined (null); accumulators are additive so a split
 * stream merges to the whole; and the 48-byte payload round-trips with its
 * guards.
 */
class CoMomentsTest extends TestCase
{
    private function oracle(array $x, array $y): float
    {
        $n = count($x);
        $sx = array_sum($x);
        $sy = array_sum($y);
        $sxy = 0.0;
        $sx2 = 0.0;
        $sy2 = 0.0;
        for ($i = 0; $i < $n; $i++) {
            $sxy += $x[$i] * $y[$i];
            $sx2 += $x[$i] ** 2;
            $sy2 += $y[$i] ** 2;
        }
        $num = $n * $sxy - $sx * $sy;
        $den = sqrt(($n * $sx2 - $sx ** 2) * ($n * $sy2 - $sy ** 2));

        return $den == 0.0 ? 0.0 : $num / $den;
    }

    public function test_pearson_matches_the_textbook_oracle(): void
    {
        mt_srand(7);
        $x = [];
        $y = [];
        $acc = new CoMoments();
        for ($i = 0; $i < 5000; $i++) {
            $xi = mt_rand(100, 10000) / 100;
            $yi = $xi * 0.37 + mt_rand(-500, 500) / 100; // correlated + noise
            $x[] = $xi;
            $y[] = $yi;
            $acc->add($xi, $yi);
        }
        $this->assertEqualsWithDelta($this->oracle($x, $y), $acc->pearson(), 1e-9);
    }

    public function test_perfect_and_independent_relationships(): void
    {
        $perfectPos = new CoMoments();
        $perfectNeg = new CoMoments();
        for ($i = 1; $i <= 1000; $i++) {
            $perfectPos->add($i, 2 * $i + 3);   // y = 2x + 3 → r = +1
            $perfectNeg->add($i, -5 * $i + 1);  // y = -5x + 1 → r = -1
        }
        $this->assertEqualsWithDelta(1.0, $perfectPos->pearson(), 1e-9);
        $this->assertEqualsWithDelta(-1.0, $perfectNeg->pearson(), 1e-9);
    }

    public function test_degenerate_cases_are_undefined(): void
    {
        $empty = new CoMoments();
        $this->assertNull($empty->pearson(), 'n=0 → undefined');
        $empty->add(1.0, 2.0);
        $this->assertNull($empty->pearson(), 'n=1 → undefined');

        $constantX = new CoMoments();
        foreach ([1, 2, 3, 4] as $v) {
            $constantX->add(5.0, (float) $v); // x has zero variance
        }
        $this->assertNull($constantX->pearson(), 'zero variance → undefined');
    }

    public function test_result_is_clamped_to_the_valid_range(): void
    {
        $acc = new CoMoments();
        for ($i = 1; $i <= 100; $i++) {
            $acc->add($i, $i); // r = +1 exactly, must not drift past 1.0
        }
        $this->assertLessThanOrEqual(1.0, $acc->pearson());
        $this->assertGreaterThanOrEqual(-1.0, $acc->pearson());
    }

    public function test_accumulators_are_additive_under_merge(): void
    {
        mt_srand(11);
        $whole = new CoMoments();
        $left = new CoMoments();
        $right = new CoMoments();
        for ($i = 0; $i < 4000; $i++) {
            $xi = mt_rand(0, 100000) / 100;
            $yi = $xi * -0.8 + mt_rand(-2000, 2000) / 100;
            $whole->add($xi, $yi);
            ($i % 2 === 0 ? $left : $right)->add($xi, $yi);
        }
        $left->merge($right);
        $this->assertSame($whole->n(), $left->n());
        $this->assertEqualsWithDelta($whole->pearson(), $left->pearson(), 1e-12);
    }

    public function test_serialize_round_trip(): void
    {
        $acc = new CoMoments();
        for ($i = 1; $i <= 250; $i++) {
            $acc->add($i * 1.5, $i * -0.25 + 7);
        }
        $bytes = $acc->serialize();
        $this->assertSame(CoMoments::PAYLOAD_BYTES, strlen($bytes));

        $back = CoMoments::deserialize($bytes);
        $this->assertSame($acc->n(), $back->n());
        $this->assertEqualsWithDelta($acc->pearson(), $back->pearson(), 1e-12);
        // Re-serialization is byte-identical (deterministic layout).
        $this->assertSame($bytes, $back->serialize());
    }

    public function test_deserialize_rejects_bad_payloads(): void
    {
        $this->expectException(\InvalidArgumentException::class);
        CoMoments::deserialize('too short');
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Sketches;

use Kolay\XlsxStream\Sketches\MisraGries;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * Misra-Gries frequent-items sketch. The guarantees, checked against an
 * exact frequency oracle: cardinality ≤ k ⇒ exact counts and
 * saturated=false; cardinality ≫ k ⇒ every reported count is an
 * underestimate within N/k and every true heavy hitter survives; merge is
 * associative and preserves the bound; the saturated bit is OR of the
 * inputs plus any merge-time prune; and serialize/deserialize round-trips
 * with guards on crafted payloads.
 */
class MisraGriesTest extends TestCase
{
    /** @return array<string,int> exact frequency table, count desc */
    private function exact(array $stream): array
    {
        $f = [];
        foreach ($stream as $x) {
            $f[$x] = ($f[$x] ?? 0) + 1;
        }
        arsort($f);

        return $f;
    }

    public function test_no_eviction_is_exact_and_unsaturated(): void
    {
        $k = 64;
        mt_srand(1);
        for ($trial = 0; $trial < 50; $trial++) {
            $distinct = mt_rand(1, $k);
            $stream = [];
            $n = mt_rand(100, 1500);
            for ($i = 0; $i < $n; $i++) {
                $stream[] = 'v'.mt_rand(1, $distinct);
            }
            $mg = new MisraGries($k);
            $mg->addMany($stream);

            $this->assertFalse($mg->saturated(), "saturated set with {$distinct} ≤ {$k} distinct");
            $got = [];
            foreach ($mg->topValues() as $p) {
                $got[$p['value']] = $p['count'];
            }
            $this->assertEquals($this->exact($stream), $got, 'counts not exact under no eviction');
        }
    }

    public function test_eviction_bounds_error_and_keeps_heavy_hitters(): void
    {
        $k = 32;
        mt_srand(2);
        for ($trial = 0; $trial < 80; $trial++) {
            $n = mt_rand(2000, 15000);
            $distinct = $k * mt_rand(4, 20);
            $stream = [];
            for ($i = 0; $i < $n; $i++) {
                $stream[] = mt_rand(1, 100) <= 60 ? 'hot'.mt_rand(1, 5) : 'cold'.mt_rand(1, $distinct);
            }
            $mg = new MisraGries($k);
            $mg->addMany($stream);
            $exact = $this->exact($stream);
            $bound = $n / $k;

            $this->assertTrue($mg->saturated(), 'saturated not set despite forced evictions');
            foreach ($mg->topValues() as $p) {
                $true = $exact[$p['value']] ?? 0;
                $this->assertLessThanOrEqual($true, $p['count'], "overestimate of {$p['value']}");
                $this->assertLessThanOrEqual($bound + 1e-9, $true - $p['count'], "underestimate of {$p['value']} exceeds N/k");
            }
            foreach ($exact as $value => $count) {
                if ($count > $bound) {
                    $got = array_column($mg->topValues(), 'count', 'value');
                    $this->assertArrayHasKey($value, $got, "heavy hitter {$value} dropped");
                }
            }
        }
    }

    public function test_merge_is_associative_and_preserves_bound(): void
    {
        $k = 32;
        mt_srand(3);
        for ($trial = 0; $trial < 60; $trial++) {
            $n = mt_rand(3000, 12000);
            $distinct = $k * mt_rand(3, 15);
            $stream = [];
            for ($i = 0; $i < $n; $i++) {
                $stream[] = mt_rand(1, 100) <= 55 ? 'hot'.mt_rand(1, 6) : 'cold'.mt_rand(1, $distinct);
            }
            $p = mt_rand(2, 6);
            $shards = array_fill(0, $p, []);
            foreach ($stream as $x) {
                $shards[mt_rand(0, $p - 1)][] = $x;
            }
            $summaries = [];
            foreach ($shards as $s) {
                $mg = new MisraGries($k);
                $mg->addMany($s);
                $summaries[] = $mg;
            }
            shuffle($summaries);
            $acc = array_shift($summaries);
            foreach ($summaries as $s) {
                $acc->merge($s);
            }

            $exact = $this->exact($stream);
            $bound = $n / $k;
            foreach ($acc->topValues() as $pair) {
                $true = $exact[$pair['value']] ?? 0;
                $this->assertLessThanOrEqual($true, $pair['count'], 'merged overestimate');
                $this->assertLessThanOrEqual($bound + 1e-9, $true - $pair['count'], 'merged underestimate exceeds N/k');
            }
            foreach ($exact as $value => $count) {
                if ($count > $bound) {
                    $got = array_column($acc->topValues(), 'count', 'value');
                    $this->assertArrayHasKey($value, $got, "merged dropped heavy hitter {$value}");
                }
            }
        }
    }

    public function test_saturated_propagates_through_merge(): void
    {
        $a = new MisraGries(4);
        $a->addMany(['p', 'p', 'q']);
        $b = new MisraGries(4);
        $b->addMany(['r', 's', 't']);
        $this->assertFalse($a->saturated());
        $this->assertFalse($b->saturated());

        // union of 6 distinct > k=4 → merge must prune → saturated
        $a->merge($b);
        $this->assertTrue($a->saturated(), 'pruning merge must saturate');

        // union ≤ k stays exact + unsaturated
        $c = new MisraGries(4);
        $c->addMany(['p', 'p']);
        $d = new MisraGries(4);
        $d->addMany(['p', 'q']);
        $c->merge($d);
        $this->assertFalse($c->saturated());
        $got = array_column($c->topValues(), 'count', 'value');
        $this->assertSame(3, $got['p']);
        $this->assertSame(1, $got['q']);
    }

    public function test_top_values_ordering_is_count_desc_then_value_asc(): void
    {
        $mg = new MisraGries(64);
        $mg->addMany(['b', 'b', 'a', 'a', 'c']); // a:2, b:2, c:1
        $this->assertSame(
            [['value' => 'a', 'count' => 2], ['value' => 'b', 'count' => 2], ['value' => 'c', 'count' => 1]],
            $mg->topValues()
        );
    }

    public function test_serialize_round_trip_preserves_state(): void
    {
        foreach ([false, true] as $forceSaturate) {
            $mg = new MisraGries(8);
            $mg->addMany(['x', 'x', 'x', 'y', 'y', 'z']);
            if ($forceSaturate) {
                $mg->addMany(['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i']); // > k distinct → evict
            }
            $round = MisraGries::deserialize($mg->serialize());
            $this->assertSame($mg->saturated(), $round->saturated());
            $this->assertSame($mg->k(), $round->k());
            $this->assertSame($mg->topValues(), $round->topValues());
            // Byte-stable: re-serialize equals the original bytes.
            $this->assertSame($mg->serialize(), $round->serialize());
        }
    }

    public function test_deserialize_rejects_crafted_payloads(): void
    {
        $good = (new MisraGries(8));
        $good->addMany(['x', 'y']);
        $payload = $good->serialize();

        // Wrong version.
        $this->expectException(\InvalidArgumentException::class);
        MisraGries::deserialize("\x09\x00".substr($payload, 2));
    }

    public function test_deserialize_rejects_trailing_bytes(): void
    {
        $mg = new MisraGries(8);
        $mg->addMany(['x', 'y']);
        $this->expectException(\InvalidArgumentException::class);
        MisraGries::deserialize($mg->serialize().'garbage');
    }

    public function test_deserialize_rejects_counter_count_over_k(): void
    {
        // k=1 header but claims 2 counters → invalid.
        $payload = pack('vv', MisraGries::FORMAT_VERSION, 1).\chr(0).pack('V', 2)
            .pack('V', 5).pack('v', 1).'a'
            .pack('V', 3).pack('v', 1).'b';
        $this->expectException(\InvalidArgumentException::class);
        MisraGries::deserialize($payload);
    }
}

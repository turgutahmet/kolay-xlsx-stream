<?php

namespace Kolay\XlsxStream\Tests\Sketches;

use Kolay\XlsxStream\Sketches\HyperLogLog;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * Oracle tests for the HyperLogLog sketch: estimates against exact
 * distinct counts across the cardinality range the linear-counting
 * correction and the raw estimator each own, the exact-merge property
 * (register-wise max IS the union sketch, byte-for-byte), and the
 * untrusted-input deserialization guards.
 *
 * Tolerance: ±5 % at every tested cardinality (standard error at p=11
 * is 1.04/√2048 ≈ 2.3 %; measured on the fixed inputs below: 0 % at 10,
 * 0.4 % at 1K, 2.96 % at 100K).
 */
class HyperLogLogTest extends TestCase
{
    #[\PHPUnit\Framework\Attributes\DataProvider('cardinalities')]
    public function test_estimate_within_five_percent_of_exact(int $cardinality): void
    {
        $hll = new HyperLogLog();
        // Every value added 3 times — duplicates must not move the
        // estimate (that is the entire point of the sketch).
        for ($i = 0; $i < $cardinality * 3; $i++) {
            $hll->add('value-'.($i % $cardinality));
        }

        $estimate = $hll->count();

        $this->assertEqualsWithDelta(
            $cardinality,
            $estimate,
            $cardinality * 0.05,
            "estimate {$estimate} for exact cardinality {$cardinality}"
        );
    }

    /** @return array<string, array{int}> */
    public static function cardinalities(): array
    {
        return [
            'tiny (linear counting)' => [10],
            'mid (linear counting)' => [1000],
            'large (raw estimator)' => [100000],
        ];
    }

    public function test_empty_sketch_counts_zero(): void
    {
        $this->assertSame(0, (new HyperLogLog())->count());
    }

    public function test_single_value_counts_one(): void
    {
        $hll = new HyperLogLog();
        $hll->add('only');
        $hll->add('only');

        $this->assertSame(1, $hll->count());
    }

    // ------------------------------------------------------------------
    // Merge — exact union, the stitch-critical property
    // ------------------------------------------------------------------

    public function test_merged_halves_equal_union_sketch_byte_for_byte(): void
    {
        $a = new HyperLogLog();
        $b = new HyperLogLog();
        $whole = new HyperLogLog();
        for ($i = 0; $i < 50000; $i++) {
            $v = "key-{$i}";
            $whole->add($v);
            ($i % 2 === 0 ? $a : $b)->add($v);
        }

        $a->merge($b);

        // Stronger than "close": register-wise max reproduces the
        // single-pass sketch EXACTLY, so the serialized bytes match.
        $this->assertSame($whole->serialize(), $a->serialize());
        $this->assertSame($whole->count(), $a->count());
    }

    public function test_merge_with_overlapping_streams_counts_the_union(): void
    {
        $a = new HyperLogLog();
        $b = new HyperLogLog();
        for ($i = 0; $i < 1000; $i++) {
            $a->add("u-{$i}");        // 0..999
        }
        for ($i = 500; $i < 1500; $i++) {
            $b->add("u-{$i}");        // 500..1499 — union is 1500
        }

        $a->merge($b);

        $this->assertEqualsWithDelta(1500, $a->count(), 1500 * 0.05);
    }

    public function test_merge_rejects_mismatched_precision(): void
    {
        $this->expectException(\InvalidArgumentException::class);
        $this->expectExceptionMessage('different precision');

        (new HyperLogLog(11))->merge(new HyperLogLog(12));
    }

    // ------------------------------------------------------------------
    // Serialization
    // ------------------------------------------------------------------

    public function test_serialize_round_trips(): void
    {
        $hll = new HyperLogLog();
        for ($i = 0; $i < 5000; $i++) {
            $hll->add("item-{$i}");
        }

        $payload = $hll->serialize();
        $this->assertSame(3 + 2048, strlen($payload), 'p=11 payload is version + p + 2048 registers');

        $restored = HyperLogLog::deserialize($payload);
        $this->assertSame($hll->count(), $restored->count());
        $this->assertSame(11, $restored->precision());
        $this->assertSame($payload, $restored->serialize());
    }

    #[\PHPUnit\Framework\Attributes\DataProvider('corruptPayloads')]
    public function test_deserialize_rejects_corrupt_payloads(string $label, string $payload): void
    {
        $this->expectException(\InvalidArgumentException::class);

        HyperLogLog::deserialize($payload);
    }

    /** @return array<string, array{string, string}> */
    public static function corruptPayloads(): array
    {
        $valid = (new HyperLogLog())->serialize();

        $overRank = $valid;
        $overRank[3] = chr(64 - 11 + 2); // register 0 above the 64-p+1 rank bound

        return [
            'too short' => ['too short', 'ab'],
            'bad version' => ['bad version', pack('v', 9).substr($valid, 2)],
            'precision out of range' => ['precision out of range', substr($valid, 0, 2).chr(3).substr($valid, 3)],
            'length mismatch' => ['length mismatch', $valid.'x'],
            'truncated registers' => ['truncated registers', substr($valid, 0, strlen($valid) - 5)],
            'register above rank bound' => ['register above rank bound', $overRank],
        ];
    }

    public function test_rejects_out_of_range_precision(): void
    {
        $this->expectException(\InvalidArgumentException::class);

        new HyperLogLog(3);
    }

    public function test_canonical_forms_hash_distinctly(): void
    {
        // The sketch is string-in/estimate-out — byte-distinct inputs
        // are distinct values. (Canonicalization POLICY is the writer's
        // job and is tested with the writer.)
        $hll = new HyperLogLog();
        $hll->add('1');
        $hll->add('1.0');
        $hll->add('1.00');
        $hll->add(' 1');

        $this->assertSame(4, $hll->count());
    }
}

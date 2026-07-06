<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\Bucket;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * E4 — Bucket::year/month/day: ready-made monotone bucket callables for
 * groupStats() over a date column, so "GROUP BY month" is one call
 * instead of hand-rolling the Excel-serial → calendar math. The buckets
 * mirror the reader's serial convention (unix = (serial − 25569) × 86400)
 * and stay monotone, so the group-pure block pushdown still applies.
 */
class BucketTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-bucket-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /** One row per day across ~5 months in 2026; amount = day-of-run. */
    private function writeFixture(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 30)->withColumnStats([1, 2]);
        $writer->setBufferFlushInterval(30);
        $writer->startFile(['day', 'amount']);
        $base = new \DateTimeImmutable('2026-01-01 00:00:00', new \DateTimeZone('UTC'));
        for ($i = 0; $i < 150; $i++) {
            $writer->writeRow([$base->modify("+{$i} days"), (float) ($i + 1)]);
        }
        $writer->finishFile();
    }

    public function test_bucket_helpers_produce_expected_keys(): void
    {
        // Excel serial for 2026-03-15 UTC: (unix + 2209161600)/86400.
        $unix = (new \DateTimeImmutable('2026-03-15 00:00:00', new \DateTimeZone('UTC')))->getTimestamp();
        $serial = ($unix + 2209161600) / 86400;

        $this->assertSame(2026, (Bucket::year())($serial));
        $this->assertSame(202603, (Bucket::month())($serial));
        $this->assertSame(20260315, (Bucket::day())($serial));
    }

    public function test_group_by_month_matches_manual_oracle(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $stats = $reader->groupStats(1, 2, Bucket::month());

        // Oracle: bucket every row's date the same way.
        $base = new \DateTimeImmutable('2026-01-01 00:00:00', new \DateTimeZone('UTC'));
        $oracle = [];
        for ($i = 0; $i < 150; $i++) {
            $ym = (int) $base->modify("+{$i} days")->format('Ym');
            $oracle[$ym] = ($oracle[$ym] ?? 0.0) + ($i + 1);
        }

        $got = [];
        foreach ($stats as $row) {
            $got[(int) $row['group']] = $row['sum'];
        }
        ksort($got);

        $this->assertSame(array_keys($oracle), array_keys($got), 'month buckets mismatch');
        foreach ($oracle as $ym => $sum) {
            $this->assertEqualsWithDelta($sum, $got[$ym], 1e-6, "sum for {$ym}");
        }
        $reader->close();
    }

    public function test_group_by_year_collapses_to_one_group(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $stats = $reader->groupStats(1, 2, Bucket::year());
        $this->assertCount(1, $stats);              // all 150 days are in 2026
        $this->assertSame(2026, (int) $stats[0]['group']);
        $this->assertSame(150, $stats[0]['count']);
        $this->assertEqualsWithDelta(150 * 151 / 2, $stats[0]['sum'], 1e-6);
        $reader->close();
    }
}

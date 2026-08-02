<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * exactQuantile() — the Rank-Sandwich: a t-digest estimate bracketed by the
 * STAT zone-map rank certificate, then an EXACT nearest-rank quantile read
 * by scanning only the blocks that could hold the target value. Gates: the
 * answer equals a brute-force oracle across shapes and quantiles; a clustered
 * column prunes to a handful of blocks while a uniform one honestly scans
 * more; the sorted fast-path answers with zero blocks; a scan budget degrades
 * to the certified estimate with an `exceeded` flag rather than throwing.
 */
class ExactQuantileTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-exactq-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /** Nearest-rank oracle: t = max(1, ceil(q*N)); q-quantile = sorted[t-1]. */
    private function oracle(array $values, float $q): float
    {
        sort($values);
        $n = count($values);
        $t = max(1, (int) ceil($q * $n));
        $t = min($t, $n);

        return (float) $values[$t - 1];
    }

    /**
     * @param  list<float>  $data  the amount column values, one per data row
     */
    private function write(array $data): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 200)
            ->withColumnStats([2])
            ->withColumnSketches([2])
            ->setBufferFlushInterval(200);
        $writer->startFile(['id', 'amount']);
        foreach ($data as $i => $v) {
            $writer->writeRow([$i + 1, $v]);
        }
        $writer->finishFile();
    }

    public function test_clustered_column_is_exact_and_prunes(): void
    {
        // Near-sorted: id + small jitter, so most blocks are disjoint ranges
        // and the scan touches only a few.
        mt_srand(11);
        $data = [];
        for ($i = 1; $i <= 4000; $i++) {
            $data[] = round($i + mt_rand(-150, 150) / 100, 2);
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        foreach ([0.0, 0.25, 0.5, 0.9, 0.99, 1.0] as $q) {
            $res = $reader->exactQuantile('amount', $q);
            $this->assertNotNull($res);
            $this->assertTrue($res['exact'], "q={$q} must be exact");
            $this->assertFalse($res['exceeded']);
            $this->assertSame($this->oracle($data, $q), $res['value'], "q={$q} value");
        }
        // Median on a clustered column must prune hard — far fewer than all.
        $median = $reader->exactQuantile('amount', 0.5);
        $this->assertLessThan(8, $median['blocksScanned'], 'clustered median should prune to a few blocks');
        $reader->close();
    }

    public function test_uniform_column_is_still_exact(): void
    {
        mt_srand(23);
        $data = [];
        for ($i = 1; $i <= 3000; $i++) {
            $data[] = round(mt_rand(0, 1_000_000) / 100, 2);
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        foreach ([0.1, 0.5, 0.95] as $q) {
            $res = $reader->exactQuantile('amount', $q);
            $this->assertTrue($res['exact']);
            $this->assertSame($this->oracle($data, $q), $res['value'], "uniform q={$q}");
        }
        $reader->close();
    }

    public function test_duplicates_and_missing_values(): void
    {
        // Heavy ties + blanks + non-numeric text interleaved.
        $data = [];
        $numeric = [];
        for ($i = 1; $i <= 2000; $i++) {
            $v = (float) ($i % 50); // 50 distinct values, many ties
            $data[] = $v;
            $numeric[] = $v;
        }
        // Rewrite via a raw writer so we can also inject blanks/text.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 200)->withColumnStats([2])->withColumnSketches([2])->setBufferFlushInterval(200);
        $writer->startFile(['id', 'amount']);
        $row = 0;
        foreach ($data as $v) {
            $writer->writeRow([++$row, $v]);
        }
        // A run of blanks/text at the end (non-numeric → excluded from the population).
        for ($k = 0; $k < 100; $k++) {
            $writer->writeRow([++$row, $k % 2 === 0 ? '' : 'n/a']);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        foreach ([0.0, 0.5, 0.75, 1.0] as $q) {
            $res = $reader->exactQuantile('amount', $q);
            $this->assertTrue($res['exact'], "ties q={$q}");
            $this->assertSame($this->oracle($numeric, $q), $res['value'], "ties q={$q} value");
        }
        $reader->close();
    }

    public function test_sorted_fast_path_scans_zero_blocks(): void
    {
        // Strictly ascending, fully numeric → the t-th value is a single
        // rowAt, no block scan.
        $data = [];
        for ($i = 1; $i <= 2000; $i++) {
            $data[] = (float) ($i * 2);
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $res = $reader->exactQuantile('amount', 0.95);
        $this->assertTrue($res['exact']);
        $this->assertSame(0, $res['blocksScanned'], 'sorted fast path reads no blocks');
        $this->assertSame($this->oracle($data, 0.95), $res['value']);
        $reader->close();
    }

    public function test_scan_budget_degrades_to_certified_estimate(): void
    {
        mt_srand(31);
        $data = [];
        for ($i = 1; $i <= 3000; $i++) {
            $data[] = round(mt_rand(0, 1_000_000) / 100, 2);
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        // A uniform median scans many blocks; a budget of 1 forces degrade.
        $res = $reader->exactQuantile('amount', 0.5, maxScanBlocks: 1);
        $this->assertFalse($res['exact'], 'over budget must not claim exact');
        $this->assertTrue($res['exceeded']);
        $this->assertSame(0, $res['blocksScanned']);
        // The degraded value is the digest estimate — near the true median.
        $this->assertEqualsWithDelta($this->oracle($data, 0.5), $res['value'], 30000.0);
        $reader->close();
    }

    public function test_null_without_sketch(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([2])->setBufferFlushInterval(200); // STAT but no TDIG
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->exactQuantile('amount', 0.5), 'no digest → cannot bracket → null');
        $reader->close();
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * TOPK (D4) — per-column Misra-Gries frequent-items sketches, writer +
 * codec round-trip (Increment A). Gates: exact categorical distribution
 * when cardinality ≤ k (saturated=false) against an exact oracle; top-k
 * with the saturated bit set when cardinality > k, heavy hitters retained;
 * header excluded; empty cells never spend a counter; additive (a file
 * without withTopValues carries no TOPK and is byte-identical).
 */
class TopValuesTest extends TestCase
{
    private string $testFile;
    private const ENTRY = 'xl/worksheets/sheet1.xml';

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-topk-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    private function decode(): ReaderIndex
    {
        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        return ReaderIndex::decode($cd->readEntry($source, ReaderIndex::ENTRY_PATH));
    }

    public function test_low_cardinality_is_exact_and_unsaturated(): void
    {
        $statuses = ['paid', 'pending', 'refunded', 'failed'];
        $oracle = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withTopValues([2], 64)->setBufferFlushInterval(100);
        $writer->startFile(['id', 'status']);
        for ($i = 1; $i <= 2000; $i++) {
            // Skewed but only 4 distinct << k=64 → exact.
            $s = $statuses[$i % 100 < 84 ? 0 : (($i % 4))];
            $writer->writeRow([$i, $s]);
            $oracle[$s] = ($oracle[$s] ?? 0) + 1;
        }
        $writer->finishFile();

        $sketch = $this->decode()->columnTopValues(self::ENTRY, 2);
        $this->assertNotNull($sketch);
        $this->assertFalse($sketch->saturated(), 'cardinality ≤ k must not saturate');

        $got = [];
        foreach ($sketch->topValues() as $p) {
            $got[$p['value']] = $p['count'];
        }
        arsort($oracle);
        $this->assertEquals($oracle, $got, 'low-cardinality counts must be exact');
    }

    public function test_high_cardinality_saturates_and_keeps_heavy_hitters(): void
    {
        $oracle = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withTopValues([2], 16)->setBufferFlushInterval(100); // small k → forced eviction
        $writer->startFile(['id', 'city']);
        $n = 10_000;
        for ($i = 1; $i <= $n; $i++) {
            // A few hot cities + a long cold tail (≫ k distinct).
            $city = $i % 100 < 55 ? 'hot'.($i % 4) : 'cold'.($i % 800);
            $writer->writeRow([$i, $city]);
            $oracle[$city] = ($oracle[$city] ?? 0) + 1;
        }
        $writer->finishFile();

        $sketch = $this->decode()->columnTopValues(self::ENTRY, 2);
        $this->assertTrue($sketch->saturated(), 'cardinality ≫ k must saturate');

        $bound = $n / 16;
        $got = array_column($sketch->topValues(), 'count', 'value');
        foreach ($got as $value => $stored) {
            $true = $oracle[$value];
            $this->assertLessThanOrEqual($true, $stored, "overestimate of {$value}");
            $this->assertLessThanOrEqual($bound + 1e-9, $true - $stored, "underestimate of {$value} exceeds N/k");
        }
        // Every hot city (freq ≫ N/k) must survive.
        foreach ($oracle as $value => $count) {
            if ($count > $bound) {
                $this->assertArrayHasKey($value, $got, "heavy hitter {$value} dropped");
            }
        }
    }

    public function test_header_excluded_and_empty_cells_not_counted(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withTopValues([1])->setBufferFlushInterval(50);
        // Header 'status' equals no data value; every 3rd cell is empty.
        $writer->startFile(['status']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow([$i % 3 === 0 ? '' : 'status']); // 'status' also the header text
        }
        $writer->finishFile();

        $sketch = $this->decode()->columnTopValues(self::ENTRY, 1);
        $got = array_column($sketch->topValues(), 'count', 'value');

        // 200 data rows carry 'status'; the header row and 100 empty cells
        // contribute nothing.
        $this->assertSame(200, $got['status']);
        $this->assertArrayNotHasKey('', $got, 'empty cells must not be a tracked value');
    }

    public function test_numeric_values_fold_by_canonical_string(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withTopValues([2])->setBufferFlushInterval(50);
        $writer->startFile(['id', 'rating']);
        // ratings 1..5, integer, most common is 5.
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([$i, $i % 100 < 60 ? 5 : ($i % 5) + 1]);
        }
        $writer->finishFile();

        $sketch = $this->decode()->columnTopValues(self::ENTRY, 2);
        $this->assertFalse($sketch->saturated());
        $top = $sketch->topValues()[0];
        $this->assertSame('5', $top['value'], 'numeric values fold by canonical string');
    }

    public function test_no_top_values_emits_no_topk_and_is_additive(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([1])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'status']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow([$i, 'paid']);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $raw = $cd->readEntry($source, ReaderIndex::ENTRY_PATH);
        $this->assertStringNotContainsString('TOPK', $raw);
        $this->assertNull(ReaderIndex::decode($raw)->columnTopValues(self::ENTRY, 2));
        $this->assertSame([], ReaderIndex::decode($raw)->topValueColumns(self::ENTRY));
    }
}

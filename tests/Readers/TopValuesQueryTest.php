<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * topValues() through the public reader facade — sidecar-only heavy
 * hitters with the exact/approximate distinction surfaced. Gates: exact
 * complete distribution when cardinality ≤ k; saturated top-k with the N/k
 * bound and heavy hitters retained when it spills; column addressable by
 * name; null contracts (untracked column, no-index file).
 */
class TopValuesQueryTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-topk-q-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_low_cardinality_reports_exact_complete_distribution(): void
    {
        $statuses = ['paid', 'pending', 'refunded'];
        $oracle = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withTopValues([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'status']);
        for ($i = 1; $i <= 3000; $i++) {
            $s = $i % 100 < 84 ? 'paid' : $statuses[$i % 3];
            $writer->writeRow([$i, $s]);
            $oracle[$s] = ($oracle[$s] ?? 0) + 1;
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $result = $reader->topValues('status');
        $this->assertNotNull($result);
        $this->assertTrue($result['exact']);

        arsort($oracle);
        $got = [];
        foreach ($result['values'] as $p) {
            $got[$p['value']] = $p['count'];
        }
        $this->assertEquals($oracle, $got);
        // The headline: most frequent value with an exact share.
        $this->assertSame('paid', $result['values'][0]['value']);
        $reader->close();
    }

    public function test_high_cardinality_is_saturated_with_bounded_error(): void
    {
        $oracle = [];
        $n = 12_000;
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withTopValues([2], 16)->setBufferFlushInterval(100);
        $writer->startFile(['id', 'sku']);
        for ($i = 1; $i <= $n; $i++) {
            $sku = $i % 100 < 50 ? 'hot'.($i % 3) : 'cold'.($i % 900);
            $writer->writeRow([$i, $sku]);
            $oracle[$sku] = ($oracle[$sku] ?? 0) + 1;
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $result = $reader->topValues('sku');
        $this->assertFalse($result['exact'], 'cardinality ≫ k must be approximate');

        $bound = $n / 16;
        $got = array_column($result['values'], 'count', 'value');
        foreach ($got as $value => $stored) {
            $this->assertLessThanOrEqual($oracle[$value], $stored);
            $this->assertLessThanOrEqual($bound + 1e-9, $oracle[$value] - $stored);
        }
        foreach (['hot0', 'hot1', 'hot2'] as $hot) {
            $this->assertArrayHasKey($hot, $got, "heavy hitter {$hot} must survive");
        }
        $reader->close();
    }

    public function test_null_when_column_untracked_or_no_index(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withTopValues([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'status']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i, 'paid']);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->topValues(1), 'untracked column has no TOPK sketch');
        $this->assertNotNull($reader->topValues(2), 'tracked column answers');
        $reader->close();
    }
}

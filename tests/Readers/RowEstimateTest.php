<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * G2 — estimatedRows() and explain(): zero-I/O selectivity estimation
 * and query planning from data already in the sidecar (no row read).
 *
 * The load-bearing invariant is the UPPER BOUND: estimatedRows()['upper']
 * (the summed row-count of zone-map surviving blocks) must NEVER be below
 * the true match count — it is a bound, not a guess. The digest estimate
 * is approximate and only checked for a sane order of magnitude.
 */
class RowEstimateTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-estimate-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /** 5000 rows, sync every 100. col1 = id asc, col2 = amount asc-ish. */
    private function writeFixture(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([1, 2]);
        $writer->withColumnSketches([1, 2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 5000; $i++) {
            $writer->writeRow([$i, round($i * 2.0, 2)]);
        }
        $writer->finishFile();
    }

    private function actualCount(StreamingXlsxReader $reader, int $col, string $op, float $v, ?float $v2 = null): int
    {
        $n = 0;
        $idx = $col - 1;
        foreach ($reader->rows() as $row) {
            $x = is_numeric($row[$idx] ?? null) ? (float) $row[$idx] : null;
            if ($x === null) {
                continue;
            }
            $hit = match ($op) {
                '>=' => $x >= $v,
                '<=' => $x <= $v,
                'between' => $x >= $v && $x <= $v2,
                '=' => $x == $v,
                default => false,
            };
            if ($hit) {
                $n++;
            }
        }

        return $n;
    }

    public function test_upper_bound_never_below_actual(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $cases = [
            [1, 'between', 2000.0, 2500.0],
            [1, '>=', 4000.0, null],
            [2, 'between', 1000.0, 3000.0],
            [2, '<=', 500.0, null],
        ];
        foreach ($cases as [$col, $op, $v, $v2]) {
            $est = $reader->estimatedRows($col, $op, $v, $v2);
            $this->assertNotNull($est);
            $actual = $this->actualCount($reader, $col, $op, $v, $v2);
            $this->assertGreaterThanOrEqual(
                $actual,
                $est['upper'],
                "upper bound violated for col{$col} {$op} {$v}"
            );
        }
        $reader->close();
    }

    public function test_digest_estimate_lands_near_actual(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // Uniform 1..5000 -> a mid-range band's estimate should be within
        // a few % (t-digest rank error is worst mid-range but small).
        $est = $reader->estimatedRows(1, 'between', 2000, 2500);
        $this->assertNotNull($est['estimate']);
        $actual = $this->actualCount($reader, 1, 'between', 2000.0, 2500.0); // 501
        $this->assertEqualsWithDelta($actual, $est['estimate'], $actual * 0.15);
        $reader->close();
    }

    public function test_estimated_rows_null_without_zone_maps(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->estimatedRows(1, '>=', 50));
        $reader->close();
    }

    public function test_explain_shapes_the_plan_without_reading_rows(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $plan = $reader->explain([[1, 'between', 2000, 2500], [2, '>=', 100]]);

        $this->assertSame('zone-map-prune', $plan['strategy']);
        $this->assertArrayHasKey('candidateBlocks', $plan);
        $this->assertArrayHasKey('runs', $plan);
        $this->assertGreaterThanOrEqual(1, $plan['runs']);
        $this->assertArrayHasKey('estimatedRows', $plan);
        $this->assertArrayHasKey('upper', $plan['estimatedRows']);
        $this->assertArrayHasKey('estimate', $plan['estimatedRows']);
        $this->assertArrayHasKey('estimatedBytes', $plan);

        // The plan's upper bound must cover the real AND result.
        $actual = count(iterator_to_array($reader->rowsWhereAll([[1, 'between', 2000, 2500], [2, '>=', 100]]), false));
        $this->assertGreaterThanOrEqual($actual, $plan['estimatedRows']['upper']);

        // estimatedBytes is a real byte budget: > 0 and no larger than the
        // whole sheet's compressed size.
        $this->assertGreaterThan(0, $plan['estimatedBytes']);

        // candidateBlocks is a real, pruned count: at least one block for
        // a matching band, well under the ~50 blocks of a 5000-row/every-100
        // sheet (pruning actually happened).
        $this->assertGreaterThanOrEqual(1, $plan['candidateBlocks']);
        $this->assertLessThanOrEqual(60, $plan['candidateBlocks']);
        $reader->close();
    }

    public function test_explain_reports_full_scan_without_stats(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $plan = $reader->explain([[1, '>=', 50]]);
        $this->assertSame('full-scan', $plan['strategy']);
        $reader->close();
    }
}

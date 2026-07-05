<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * D3 — sampleRows(): a uniform random sample of k rows fetched via the
 * index (≈k block reads, not a full N-row scan). k distinct row numbers
 * are drawn uniformly from the data range with a seeded Mt19937
 * Randomizer, then those exact rows are read. Unbiasedness is proven in
 * poc/d3_sample.php (chi-square); these tests pin the surface: a fixed
 * seed reproduces the sample, the sampled rows are REAL rows read
 * correctly, the count is exact, and edge sizes behave.
 */
class RowSampleTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-sample-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    private function writeFixture(bool $indexed = true): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        if ($indexed) {
            $writer->withRandomAccessIndex(every: 100);
        }
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'label']);
        for ($i = 1; $i <= 1000; $i++) {
            $writer->writeRow([$i, 'label-'.$i]);
        }
        $writer->finishFile();
    }

    public function test_same_seed_reproduces_the_sample(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $a = $reader->sampleRows(30, seed: 42);
        $b = $reader->sampleRows(30, seed: 42);
        $this->assertSame($a, $b, 'a fixed seed must reproduce the exact sample');
        $reader->close();
    }

    public function test_different_seeds_differ(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $a = array_keys($reader->sampleRows(30, seed: 1));
        $b = array_keys($reader->sampleRows(30, seed: 2));
        $this->assertNotSame($a, $b, 'different seeds should draw different rows');
        $reader->close();
    }

    public function test_sampled_rows_are_real_rows_read_correctly(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $sample = $reader->sampleRows(40, seed: 7);
        $this->assertCount(40, $sample);

        foreach ($sample as $rn => $row) {
            // Data row rn holds id = rn - 1 (row 2 => id 1) and its label.
            $this->assertSame((string) ($rn - 1), $row[0], "row {$rn} id mismatch");
            $this->assertSame('label-'.($rn - 1), $row[1], "row {$rn} label mismatch");
        }

        // Keys are ascending, distinct, and within the data range.
        $keys = array_keys($sample);
        $sorted = $keys;
        sort($sorted);
        $this->assertSame($sorted, $keys, 'sample must be in ascending row order');
        $this->assertSame($keys, array_values(array_unique($keys)), 'no duplicate rows');
        $this->assertGreaterThanOrEqual(2, min($keys));
        $this->assertLessThanOrEqual(1001, max($keys));
        $reader->close();
    }

    public function test_k_at_or_above_row_count_returns_all(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $all = $reader->sampleRows(999999, seed: 1);
        $this->assertCount(1000, $all);
        $this->assertSame(range(2, 1001), array_keys($all));
        $reader->close();
    }

    public function test_k_zero_or_negative_returns_empty(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertSame([], $reader->sampleRows(0));
        $this->assertSame([], $reader->sampleRows(-5));
        $reader->close();
    }

    public function test_sample_spreads_across_the_file(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // A uniform sample of 100 from 1000 should touch both halves, not
        // cluster at one end (a coarse unbiasedness canary; the rigorous
        // chi-square lives in poc/d3_sample.php).
        $keys = array_keys($reader->sampleRows(100, seed: 123));
        $firstHalf = array_filter($keys, fn ($rn) => $rn <= 501);
        $secondHalf = array_filter($keys, fn ($rn) => $rn > 501);
        $this->assertGreaterThan(20, count($firstHalf));
        $this->assertGreaterThan(20, count($secondHalf));
        $reader->close();
    }

    public function test_works_without_an_index(): void
    {
        $this->writeFixture(indexed: false);
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $sample = $reader->sampleRows(15, seed: 5);
        $this->assertCount(15, $sample);
        foreach ($sample as $rn => $row) {
            $this->assertSame((string) ($rn - 1), $row[0]);
        }
        $reader->close();
    }
}

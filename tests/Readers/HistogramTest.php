<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * histogram() — an equi-width distribution of a column over [min, max],
 * each bin's count read from the t-digest CDF (zero I/O). Gates: the bins
 * tile the range with no gap, their counts sum to the numeric total exactly
 * (cumulative rounding), each count tracks a brute-force oracle within the
 * digest's tolerance, a constant column collapses to one bin, and the null /
 * validation contracts hold.
 */
class HistogramTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-hist-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /** @param  list<float>  $data */
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

    public function test_bins_tile_the_range_and_counts_sum_to_total(): void
    {
        $data = [];
        for ($i = 1; $i <= 1000; $i++) {
            $data[] = (float) $i;
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $hist = $reader->histogram('amount', 10);
        $this->assertNotNull($hist);
        $this->assertCount(10, $hist);

        // Edges tile [min, max] with no gap.
        $this->assertSame(1.0, $hist[0]['lo']);
        $this->assertSame(1000.0, $hist[9]['hi']);
        for ($i = 1; $i < 10; $i++) {
            $this->assertSame($hist[$i - 1]['hi'], $hist[$i]['lo'], "bin {$i} abuts the previous");
        }

        // Counts sum to the numeric total exactly.
        $this->assertSame(1000, array_sum(array_column($hist, 'count')));
        $reader->close();
    }

    public function test_counts_track_a_brute_force_oracle(): void
    {
        mt_srand(53);
        $data = [];
        for ($i = 1; $i <= 5000; $i++) {
            $data[] = round(mt_rand(0, 1_000_000) / 100, 2);
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $bins = 8;
        $hist = $reader->histogram('amount', $bins);

        // Oracle over the SAME edges: bin 0 is [lo, hi] inclusive, the rest
        // (lo, hi] — mirroring the CDF's cumulative binning.
        $oracle = array_fill(0, $bins, 0);
        foreach ($data as $v) {
            foreach ($hist as $b => $bin) {
                $lo = $bin['lo'];
                $hi = $bin['hi'];
                if (($b === 0 ? $v >= $lo : $v > $lo) && $v <= $hi) {
                    $oracle[$b]++;
                    break;
                }
            }
        }

        foreach ($hist as $b => $bin) {
            // Digest rank error is small; allow a modest per-bin tolerance.
            $this->assertEqualsWithDelta($oracle[$b], $bin['count'], 120, "bin {$b} count");
        }
        $this->assertSame(5000, array_sum(array_column($hist, 'count')));
        $reader->close();
    }

    public function test_constant_column_collapses_to_one_bin(): void
    {
        $data = array_fill(0, 300, 42.0);
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $hist = $reader->histogram('amount', 10);
        $this->assertCount(1, $hist, 'a zero-width range is a single bin');
        $this->assertSame(42.0, $hist[0]['lo']);
        $this->assertSame(42.0, $hist[0]['hi']);
        $this->assertSame(300, $hist[0]['count']);
        $reader->close();
    }

    public function test_invalid_bin_count_throws(): void
    {
        $this->write([1.0, 2.0, 3.0]);
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->expectException(\InvalidArgumentException::class);
        try {
            $reader->histogram('amount', 0);
        } finally {
            $reader->close();
        }
    }

    public function test_null_without_digest(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([2])->setBufferFlushInterval(200); // no sketch
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->histogram('amount', 10), 'no digest → no distribution');
        $reader->close();
    }
}

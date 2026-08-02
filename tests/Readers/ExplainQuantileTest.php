<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * explainQuantile() — the zero-I/O quantile plan: the estimate, its
 * deterministic rank certificate, and the cost an exact answer would pay.
 * Gates: the certificate soundly brackets the estimate's true rank; the
 * quoted block/byte cost matches what exactQuantile actually reads; the
 * sorted fast-path quotes zero blocks; and the whole call reads no rows
 * (verified against a stream-spy source).
 */
class ExplainQuantileTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-explainq-'.uniqid('', true).'.xlsx';
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

    public function test_certificate_brackets_the_true_rank_of_the_estimate(): void
    {
        mt_srand(41);
        $data = [];
        for ($i = 1; $i <= 3000; $i++) {
            $data[] = round(mt_rand(0, 1_000_000) / 100, 2);
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        foreach ([0.1, 0.5, 0.9] as $q) {
            $plan = $reader->explainQuantile('amount', $q);
            $this->assertNotNull($plan);

            // True rank of the estimate: how many values are <= it.
            $trueRank = 0;
            foreach ($data as $v) {
                if ($v <= $plan['estimate']) {
                    $trueRank++;
                }
            }
            $this->assertLessThanOrEqual($trueRank, $plan['rank_lo'], "q={$q} rank_lo bound");
            $this->assertGreaterThanOrEqual($trueRank, $plan['rank_hi'], "q={$q} rank_hi bound");
            $this->assertLessThanOrEqual($plan['rank_hi'], $plan['rank_lo'], 'lo <= hi');
        }
        $reader->close();
    }

    public function test_cost_matches_what_exact_actually_scans(): void
    {
        mt_srand(43);
        $data = [];
        for ($i = 1; $i <= 3000; $i++) {
            $data[] = round($i + mt_rand(-200, 200) / 100, 2); // clustered
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        foreach ([0.25, 0.5, 0.75] as $q) {
            $plan = $reader->explainQuantile('amount', $q);
            $exact = $reader->exactQuantile('amount', $q);
            $this->assertSame(
                $exact['blocksScanned'],
                $plan['exact_would_scan_blocks'],
                "q={$q}: quoted block count must equal the executed scan"
            );
            $this->assertGreaterThan(0, $plan['exact_est_bytes'], 'a real scan quotes some bytes');
        }
        $reader->close();
    }

    public function test_sorted_fast_path_quotes_zero_blocks(): void
    {
        $data = [];
        for ($i = 1; $i <= 2000; $i++) {
            $data[] = (float) ($i * 3);
        }
        $this->write($data);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $plan = $reader->explainQuantile('amount', 0.95);
        $this->assertSame(0, $plan['exact_would_scan_blocks'], 'sorted → single indexed row, no block scan');
        $this->assertGreaterThan(0, $plan['exact_est_bytes'], 'the one covering block still has a byte cost');
        $reader->close();
    }

    public function test_null_without_digest(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([2])->setBufferFlushInterval(200); // no sketch
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 400; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->explainQuantile('amount', 0.5));
        $reader->close();
    }

    public function test_explain_quantile_reads_no_rows(): void
    {
        mt_srand(47);
        $data = [];
        for ($i = 1; $i <= 2000; $i++) {
            $data[] = round(mt_rand(0, 100000) / 100, 2);
        }
        $this->write($data);

        // Spy source: record every streamFrom so we can prove the plan is
        // computed purely from the sidecar (already in memory after open).
        $inner = new LocalFileSource($this->testFile);
        $spy = new class ($inner) implements Source {
            public array $streamOffsets = [];

            public function __construct(private Source $inner)
            {
            }

            public function size(): int
            {
                return $this->inner->size();
            }

            public function range(int $offset, int $length): string
            {
                return $this->inner->range($offset, $length);
            }

            public function streamFrom(int $offset, ?int $length = null)
            {
                $this->streamOffsets[] = $offset;

                return $this->inner->streamFrom($offset, $length);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };

        $reader = StreamingXlsxReader::from($spy);
        $reader->explainQuantile('amount', 0.5); // warms the index once
        $spy->streamOffsets = [];
        $reader->explainQuantile('amount', 0.9);
        $reader->explainQuantile('amount', 0.99);
        $this->assertSame([], $spy->streamOffsets, 'explainQuantile must read no rows');
        $reader->close();
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * End-to-end tests for the sketch query surface: quantile(), median()
 * and countDistinct() answered from the KXSI TDIG/CHLL sidecar
 * sections.
 *
 * The correctness oracle is the raw data the fixture was written from —
 * exact percentiles and exact distinct counts recomputed independently,
 * with tolerance assertions matching the sketches' documented error.
 * A counting-spy Source additionally pins the ZERO-REQUEST property:
 * once the reader is open (index cached), sketch queries must touch the
 * underlying source not at all — identical answers alone cannot prove
 * the sidecar (rather than a quiet scan) produced them.
 */
class SketchQueryTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-sketchq-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * 2000 data rows: col 1 ascending id, col 2 skewed amount (i²/1000),
     * col 3 text city from a 25-value pool, col 4 untracked.
     *
     * @return array{list<float>, list<float>} sorted id + amount values
     */
    private function writeFixture(): array
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 500);
        $writer->withColumnSketches([1, 2, 3]);
        $writer->setBufferFlushInterval(500);
        $writer->startFile(['id', 'amount', 'city', 'notes']);

        $ids = [];
        $amounts = [];
        for ($i = 1; $i <= 2000; $i++) {
            $amount = $i * $i / 1000;
            $ids[] = (float) $i;
            $amounts[] = (float) $amount; // PHP '/' yields int on exact division
            $writer->writeRow([$i, $amount, 'city-'.($i % 25), 'note '.$i]);
        }
        $writer->finishFile();

        return [$ids, $amounts];
    }

    public function test_quantiles_match_exact_percentiles_within_tolerance(): void
    {
        [$ids, $amounts] = $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // Exact extremes by contract.
        $this->assertSame(1.0, $reader->quantile(1, 0.0));
        $this->assertSame(2000.0, $reader->quantile(1, 1.0));
        $this->assertSame(min($amounts), $reader->quantile(2, 0.0));
        $this->assertSame(max($amounts), $reader->quantile(2, 1.0));

        // Interior quantiles within 1% RANK error against the exact
        // sorted data (documented digest error at n=2000 is far lower).
        foreach ([0.01, 0.25, 0.5, 0.75, 0.99] as $q) {
            foreach ([[1, $ids], [2, $amounts]] as [$col, $sorted]) {
                $estimate = $reader->quantile($col, $q);
                $exactAt = fn (float $p) => $sorted[(int) min(count($sorted) - 1, floor($p * count($sorted)))];
                $lo = $exactAt(max(0.0, $q - 0.01));
                $hi = $exactAt(min(1.0, $q + 0.01));
                $this->assertGreaterThanOrEqual($lo, $estimate, "col {$col} q={$q}");
                $this->assertLessThanOrEqual($hi, $estimate, "col {$col} q={$q}");
            }
        }

        // median() is quantile(0.5).
        $this->assertSame($reader->quantile(1, 0.5), $reader->median(1));
        $reader->close();
    }

    public function test_count_distinct_matches_exact_within_tolerance(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // 2000 distinct ids (±5%), 2000 distinct amounts, exactly 25
        // cities (small cardinality => linear counting, near-exact).
        $this->assertEqualsWithDelta(2000, $reader->countDistinct(1), 100);
        $this->assertEqualsWithDelta(2000, $reader->countDistinct(2), 100);
        $this->assertEqualsWithDelta(25, $reader->countDistinct(3), 1);
        $reader->close();
    }

    public function test_text_column_has_distinct_count_but_no_quantiles(): void
    {
        // Half the point of CHLL: text columns get distinct counts even
        // though they are invisible to every numeric surface.
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertNotNull($reader->countDistinct(3));
        $this->assertNull($reader->quantile(3, 0.5), 'no numeric values => no quantile');
        $this->assertNull($reader->median(3));
        $reader->close();
    }

    public function test_untracked_column_and_unsketched_file_return_null(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->quantile(4, 0.5));
        $this->assertNull($reader->countDistinct(4));
        $reader->close();

        // A file with an index + stats but NO sketches: quantile and
        // countDistinct stay null while columnStats() answers — the two
        // opt-ins are orthogonal reader-side too.
        @unlink($this->testFile);
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100)->withColumnStats([1]);
        $writer->startFile(['id']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->quantile(1, 0.5));
        $this->assertNull($reader->countDistinct(1));
        $this->assertNotNull($reader->columnStats(1));
        $reader->close();
    }

    public function test_sketches_without_stats_answer_while_stats_surface_is_null(): void
    {
        $this->writeFixture(); // sketches only — no withColumnStats
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertNotNull($reader->median(1));
        $this->assertNull($reader->columnStats(1), 'stats must not materialise from sketches');
        $reader->close();
    }

    public function test_quantile_validates_q_range(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        try {
            $reader->quantile(1, -0.001);
            $this->fail('q < 0 accepted');
        } catch (\InvalidArgumentException $e) {
            $this->assertStringContainsString('[0, 1]', $e->getMessage());
        }
        try {
            $reader->quantile(1, 1.001);
            $this->fail('q > 1 accepted');
        } catch (\InvalidArgumentException $e) {
            $this->assertStringContainsString('[0, 1]', $e->getMessage());
        }
        $reader->close();
    }

    public function test_multi_sheet_answers_are_per_sheet(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnSketches([1]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['v']);
        for ($i = 1; $i <= 400; $i++) {
            $writer->writeRow([$i]);              // sheet 1: 1..400
        }
        $writer->newSheet('Second', ['v']);
        for ($i = 1; $i <= 400; $i++) {
            $writer->writeRow([1000 + ($i % 4)]); // sheet 2: 4 distinct
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertEqualsWithDelta(200.5, $reader->median(1), 5.0);
        $this->assertEqualsWithDelta(400, $reader->countDistinct(1), 20);

        $reader->onSheetIndex(1);
        $this->assertEqualsWithDelta(1001.5, $reader->median(1), 1.0);
        $this->assertSame(4, $reader->countDistinct(1));
        $this->assertSame(1000.0, $reader->quantile(1, 0.0));
        $this->assertSame(1003.0, $reader->quantile(1, 1.0));

        // Switching back re-answers sheet 1 (per-entry memoization).
        $reader->onSheetIndex(0);
        $this->assertSame(1.0, $reader->quantile(1, 0.0));
        $reader->close();
    }

    public function test_sketch_queries_issue_zero_additional_requests(): void
    {
        // THE zero-request pin. After the index is loaded (one-time,
        // at open/first indexed call), quantile/median/countDistinct
        // must add NO range reads and NO stream opens — the sidecar
        // alone answers. This is what "one range request against a
        // multi-GB S3 file" rests on.
        $this->writeFixture();

        $spy = new class(new LocalFileSource($this->testFile)) implements Source
        {
            public int $rangeCalls = 0;

            public int $streamCalls = 0;

            public function __construct(private LocalFileSource $inner)
            {
            }

            public function size(): int
            {
                return $this->inner->size();
            }

            public function range(int $offset, int $length): string
            {
                $this->rangeCalls++;

                return $this->inner->range($offset, $length);
            }

            public function streamFrom(int $offset, ?int $length = null)
            {
                $this->streamCalls++;

                return $this->inner->streamFrom($offset, $length);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };

        $reader = StreamingXlsxReader::from($spy);
        $reader->rowCount(); // forces sidecar load + staleness validation

        $rangeBefore = $spy->rangeCalls;
        $streamBefore = $spy->streamCalls;

        $reader->median(1);
        $reader->quantile(2, 0.99);
        $reader->quantile(2, 0.01);
        $reader->countDistinct(1);
        $reader->countDistinct(3);
        $reader->quantile(4, 0.5);    // null path must not scan either
        $reader->countDistinct(4);

        $this->assertSame($rangeBefore, $spy->rangeCalls, 'sketch query issued a range request');
        $this->assertSame($streamBefore, $spy->streamCalls, 'sketch query opened a stream');
        $reader->close();
    }

    public function test_stale_sidecar_degrades_sketch_answers_to_null(): void
    {
        // Simulate the "another tool rewrote the sheet" staleness signal
        // exactly where the reader checks it: flip one bit of the sheet's
        // CRC-32 in the LIVE ZIP central directory, so it no longer
        // matches the sidecar's recorded sheet_crc32. The reader must
        // silently drop the WHOLE index — sketches included: a quantile
        // from a sheet that may have been replaced is worse than no
        // answer. Sequential reading keeps working (the file is fine).
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNotNull($reader->median(1), 'precondition: fresh file answers');
        $reader->close();

        $bytes = (string) file_get_contents($this->testFile);
        $needle = pack('V', 0x02014b50); // central file header signature
        $pos = strpos($bytes, $needle);
        $found = false;
        while ($pos !== false) {
            $nameLen = unpack('v', substr($bytes, $pos + 28, 2))[1];
            if (substr($bytes, $pos + 46, $nameLen) === 'xl/worksheets/sheet1.xml') {
                $crcOffset = $pos + 16; // CRC-32 field of the central header
                $bytes[$crcOffset] = chr(ord($bytes[$crcOffset]) ^ 0x01);
                $found = true;
                break;
            }
            $pos = strpos($bytes, $needle, $pos + 4);
        }
        $this->assertTrue($found, 'central directory entry for sheet1 not found');
        file_put_contents($this->testFile, $bytes);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->median(1), 'stale sidecar must not answer quantiles');
        $this->assertNull($reader->countDistinct(3), 'stale sidecar must not answer distinct counts');
        // The workbook itself still reads sequentially.
        $this->assertSame(['id', 'amount', 'city', 'notes'], $reader->header());
        $reader->close();
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * quantile(col, q, from, to) — range quantile over a sheet-row span,
 * answered by merging the TDGB superblock digests that fall entirely
 * inside the span and scanning only the partial rows at each edge.
 * Gates: aligned span reads zero sheet bytes, misaligned span stays
 * accurate against the exact-range oracle, no-TDGB column degrades to an
 * honest range scan (onFullScan fires), bad bounds throw.
 */
class RangeQuantileQueryTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-tdgb-q-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /** @return array<int, float> data-row ordinal (1-based) => value */
    private function writeFixture(int $rows = 40_000, int $sync = 1_000, bool $withTdgb = true): array
    {
        $written = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        if ($withTdgb) {
            $writer->withRangeQuantiles([2]);
        }
        $writer->withRandomAccessIndex(every: $sync)->setBufferFlushInterval($sync);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= $rows; $i++) {
            // A non-monotone but well-spread distribution so range slices
            // differ from one another and from the whole column.
            $v = round((($i * 7919) % 10_000) + $i * 0.001, 3);
            $written[$i] = $v;
            $writer->writeRow([$i, $v]);
        }
        $writer->finishFile();

        return $written;
    }

    /** @return list<array{end_row: int}> */
    private function superblocks(): array
    {
        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $index = ReaderIndex::decode($cd->readEntry($source, ReaderIndex::ENTRY_PATH));

        return $index->rangeQuantileSuperblocks('xl/worksheets/sheet1.xml', 2) ?? [];
    }

    /**
     * Exact rank of the estimate among the values whose SHEET row (data
     * row i lives at sheet row i+1; the header is row 1) falls in the
     * inclusive span [$from, $to].
     *
     * @param  array<int, float>  $written
     */
    private function assertRankAccurate(array $written, int $from, int $to, float $q, ?float $est, float $delta = 0.02): void
    {
        $vals = [];
        foreach ($written as $i => $v) {
            $sheetRow = $i + 1;
            if ($sheetRow >= $from && $sheetRow <= $to) {
                $vals[] = $v;
            }
        }
        sort($vals);
        $n = count($vals);
        $this->assertNotNull($est, "quantile returned null for span [{$from}, {$to}]");
        $this->assertGreaterThan(0, $n, 'oracle span is empty');

        $below = 0;
        foreach ($vals as $v) {
            if ($v < $est) {
                $below++;
            }
        }
        $this->assertEqualsWithDelta($q, $below / $n, $delta, "rank error at q={$q}, span [{$from}, {$to}]");
    }

    public function test_full_range_matches_whole_column_and_reads_no_sheet_bytes(): void
    {
        $written = $this->writeFixture();
        $spy = $this->spySource();
        $reader = StreamingXlsxReader::from($spy);

        // Warm the sidecar so the assertion isolates SHEET reads. Column
        // is given by index (2) so resolveColumnName() does not read the
        // header row — the assertion is about DATA bytes.
        $reader->rowCount();
        $spy->streamOffsets = [];

        foreach ([0.25, 0.5, 0.9, 0.99] as $q) {
            $est = $reader->quantile(2, $q, 1, 40_001);
            $this->assertRankAccurate($written, 1, 40_001, $q, $est);
        }

        // Every superblock is fully covered → the answer comes purely
        // from merged sidecar digests, with no sheet scan at all.
        $this->assertSame([], $spy->streamOffsets, 'a fully-covered range must read zero sheet bytes');
        $reader->close();
    }

    public function test_superblock_aligned_subrange_reads_no_sheet_bytes(): void
    {
        $written = $this->writeFixture();
        $sbs = $this->superblocks();
        $this->assertGreaterThanOrEqual(3, count($sbs), 'need ≥3 superblocks to align on an interior one');

        // Span covering exactly the interior superblock #1: its rows are
        // (sbs[0].end_row, sbs[1].end_row].
        $from = $sbs[0]['end_row'] + 1;
        $to = $sbs[1]['end_row'];

        $spy = $this->spySource();
        $reader = StreamingXlsxReader::from($spy);
        $reader->rowCount();
        $spy->streamOffsets = [];

        $est = $reader->quantile(2, 0.5, $from, $to);
        $this->assertRankAccurate($written, $from, $to, 0.5, $est);
        $this->assertSame([], $spy->streamOffsets, 'a superblock-aligned range must read zero sheet bytes');
        $reader->close();
    }

    public function test_misaligned_subrange_scans_only_the_edges(): void
    {
        $written = $this->writeFixture();

        $scans = [];
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $reader->onFullScan(function (array $ctx) use (&$scans): void {
            $scans[] = $ctx;
        });

        // 5000..25000 lands mid-superblock at both ends → edge scans, but
        // the interior superblocks are merged from the sidecar.
        foreach ([0.1, 0.5, 0.95] as $q) {
            $est = $reader->quantile('amount', $q, 5_000, 25_000);
            $this->assertRankAccurate($written, 5_000, 25_000, $q, $est);
        }

        // An edge scan is bounded and expected — it is NOT a full scan.
        $this->assertSame([], $scans, 'edge scans must not be reported as full scans');
        $reader->close();
    }

    public function test_no_tdgb_column_degrades_to_an_honest_range_scan(): void
    {
        $written = $this->writeFixture(withTdgb: false);

        $scans = [];
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $reader->onFullScan(function (array $ctx) use (&$scans): void {
            $scans[] = $ctx;
        });

        $est = $reader->quantile('amount', 0.5, 5_000, 25_000);
        $this->assertRankAccurate($written, 5_000, 25_000, 0.5, $est);

        $this->assertNotEmpty($scans, 'a column without TDGB must announce the honest scan');
        $this->assertSame('column-not-range-indexed', $scans[0]['reason']);
        $reader->close();
    }

    public function test_empty_span_returns_null(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        // Past the last data row → no values.
        $this->assertNull($reader->quantile('amount', 0.5, 50_000, 60_000));
        $reader->close();
    }

    public function test_inverted_or_out_of_domain_bounds_throw(): void
    {
        $this->writeFixture(rows: 500);
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\InvalidArgumentException::class);
        $reader->quantile('amount', 0.5, 100, 50);
    }

    private function spySource(): Source
    {
        return new class(new LocalFileSource($this->testFile)) implements Source
        {
            /** @var list<int> */
            public array $streamOffsets = [];

            public function __construct(private LocalFileSource $inner) {}

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
    }
}

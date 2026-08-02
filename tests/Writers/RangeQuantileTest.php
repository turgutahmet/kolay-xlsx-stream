<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sketches\TDigest;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * TDGB (D2) — per-superblock t-digests, writer + codec round-trip
 * (Increment A). Superblocks are ROW-space: a fixed width snapped to the
 * next sync point, bounding them to ≤64 per sheet regardless of block
 * count, built single-pass. Gates: bounded count, contiguous row spans,
 * accurate merge (all superblocks == whole-column digest), sync-snap edges.
 */
class RangeQuantileTest extends TestCase
{
    private string $testFile;
    private const ENTRY = 'xl/worksheets/sheet1.xml';
    private const S = 16384; // BaseXlsxWriter::SUPERBLOCK_ROWS

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-tdgb-'.uniqid('', true).'.xlsx';
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

    /** @return list<float> the values written to $col */
    private function write(int $rows, int $syncPeriod, callable $value, array $cols = [2]): array
    {
        $written = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRangeQuantiles($cols)->withRandomAccessIndex(every: $syncPeriod)->setBufferFlushInterval($syncPeriod);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= $rows; $i++) {
            $v = $value($i);
            $written[] = (float) $v;
            $writer->writeRow([$i, $v]);
        }
        $writer->finishFile();

        return $written;
    }

    public function test_round_trip_spans_are_contiguous_and_bounded(): void
    {
        // 40k rows > 2·S → multiple superblocks. sync every 1000.
        $this->write(40_000, 1000, fn ($i) => round($i * 0.5, 2));

        $sbs = $this->decode()->rangeQuantileSuperblocks(self::ENTRY, 2);
        $this->assertNotNull($sbs);
        $this->assertLessThanOrEqual(64, count($sbs));
        $this->assertGreaterThan(1, count($sbs));

        // end_rows strictly increasing; the last covers the final data row
        // (header = sheet row 1, data rows 2..40001 → last end_row = 40001).
        $prev = 0;
        foreach ($sbs as $sb) {
            $this->assertGreaterThan($prev, $sb['end_row']);
            $prev = $sb['end_row'];
        }
        $this->assertSame(40_001, end($sbs)['end_row']);

        // Every superblock except possibly the last spans ≥ S rows (snap).
        $start = 1; // spans are (start, end]; first data row is 2, start=1
        foreach (array_slice($sbs, 0, -1) as $sb) {
            $this->assertGreaterThanOrEqual(self::S, $sb['end_row'] - $start);
            $start = $sb['end_row'];
        }
    }

    public function test_merged_superblocks_match_whole_column_quantile(): void
    {
        $values = $this->write(40_000, 1000, fn ($i) => round($i * 0.5 + (($i * 37) % 500), 2));
        sort($values);
        $n = count($values);

        $sbs = $this->decode()->rangeQuantileSuperblocks(self::ENTRY, 2);
        $merged = new TDigest();
        foreach ($sbs as $sb) {
            $merged->merge($sb['digest']);
        }

        // Merging all superblocks reproduces the whole-column distribution
        // within t-digest rank error.
        foreach ([0.5, 0.9, 0.99] as $q) {
            $est = $merged->quantile($q);
            $below = 0;
            foreach ($values as $v) {
                if ($v < $est) {
                    $below++;
                }
            }
            $this->assertEqualsWithDelta($q, $below / $n, 0.01, "rank error at q={$q}");
        }
    }

    public function test_sync_period_larger_than_superblock_degrades_to_one_block(): void
    {
        // syncPeriod 20000 > S 16384 → each superblock is a single block.
        $this->write(40_000, 20_000, fn ($i) => (float) $i);

        $sbs = $this->decode()->rangeQuantileSuperblocks(self::ENTRY, 2);
        // 40000 rows / 20000 per block = 2 blocks = 2 superblocks.
        $this->assertCount(2, $sbs);
        $this->assertSame(20_001, $sbs[0]['end_row']); // first block ends at data row 20000 (sheet row 20001)
        $this->assertSame(40_001, $sbs[1]['end_row']);
    }

    public function test_single_superblock_for_small_sheet(): void
    {
        $this->write(500, 1000, fn ($i) => (float) $i);
        $sbs = $this->decode()->rangeQuantileSuperblocks(self::ENTRY, 2);
        $this->assertCount(1, $sbs);
        $this->assertSame(501, $sbs[0]['end_row']);
    }

    public function test_no_range_quantiles_emits_no_tdgb(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([1])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow([$i, $i * 1.5]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $raw = $cd->readEntry($source, ReaderIndex::ENTRY_PATH);
        $this->assertStringNotContainsString('TDGB', $raw);
        $this->assertNull(ReaderIndex::decode($raw)->rangeQuantileSuperblocks(self::ENTRY, 2));
    }
}

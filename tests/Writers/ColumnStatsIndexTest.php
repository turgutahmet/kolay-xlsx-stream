<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\RandomAccessIndex;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * KXSI column statistics (zone maps): encoder/decoder format contract
 * plus end-to-end writer integration.
 *
 * The load-bearing invariants:
 *   - stats-less files keep emitting version 2 BYTE-IDENTICAL to v3.0.x,
 *     so readers already in the wild keep their random access;
 *   - block_count == sync_count + 1 for every (sheet, column);
 *   - unknown TLV tags are skipped, known ones parsed — the framing that
 *     lets future sections ship without a version bump;
 *   - stats widen, never narrow: every value that could render as a
 *     numeric cell is folded into min/max/sum, so a block whose stats
 *     exclude the predicate range provably contains no match.
 */
class ColumnStatsIndexTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-stats-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    // ---------------------------------------------------------------
    // Format contract (encoder <-> decoder, no file involved)
    // ---------------------------------------------------------------

    public function test_encoder_without_stats_still_emits_version_2(): void
    {
        $payload = RandomAccessIndex::encode(10000, [
            ['entry' => 'xl/worksheets/sheet1.xml', 'total_rows' => 50, 'sheet_crc32' => 123, 'sync_points' => []],
        ]);

        $this->assertSame(2, ord($payload[4]));
        // And the v3.0.x reader path still decodes it.
        $decoded = ReaderIndex::decode($payload);
        $this->assertSame(50, $decoded->totalRows('xl/worksheets/sheet1.xml'));
        $this->assertNull($decoded->columnStats('xl/worksheets/sheet1.xml', 1));
    }

    public function test_stats_round_trip_through_reader_decoder(): void
    {
        $entry = 'xl/worksheets/sheet1.xml';
        $syncPoints = [
            ['row' => 101, 'comp_offset' => 1000, 'uncomp_offset' => 5000],
            ['row' => 201, 'comp_offset' => 2000, 'uncomp_offset' => 10000],
        ];
        $blocks = [
            ['min' => 1.0, 'max' => 100.0, 'sum' => 5050.0, 'count' => 100, 'other' => 1],
            ['min' => 101.0, 'max' => 200.0, 'sum' => 15050.0, 'count' => 100, 'other' => 0],
            ['min' => 201.0, 'max' => 250.0, 'sum' => 11275.0, 'count' => 50, 'other' => 2],
        ];

        $payload = RandomAccessIndex::encode(
            100,
            [['entry' => $entry, 'total_rows' => 253, 'sheet_crc32' => 42, 'sync_points' => $syncPoints]],
            [$entry => [['col' => 3, 'sorted_asc' => true, 'sorted_desc' => false, 'blocks' => $blocks]]]
        );

        // Version stays 2 even with stats — the v3.0.x decoder parses
        // exactly sheet_count core sections and ignores trailing TLV
        // bytes, so old readers keep random access on stats files.
        $this->assertSame(2, ord($payload[4]));

        $decoded = ReaderIndex::decode($payload);
        $stats = $decoded->columnStats($entry, 3);

        $this->assertNotNull($stats);
        $this->assertTrue($stats['sorted_asc']);
        $this->assertFalse($stats['sorted_desc']);
        $this->assertSame($blocks, $stats['blocks']);
        $this->assertSame([3], $decoded->statsColumns($entry));

        // Core fields still intact alongside the TLV section.
        $this->assertSame(253, $decoded->totalRows($entry));
        $this->assertCount(2, $decoded->syncPoints($entry));
    }

    public function test_decoder_skips_unknown_tlv_sections(): void
    {
        $entry = 'xl/worksheets/sheet1.xml';
        $payload = RandomAccessIndex::encode(
            100,
            [['entry' => $entry, 'total_rows' => 10, 'sheet_crc32' => 7, 'sync_points' => []]],
            [$entry => [['col' => 1, 'sorted_asc' => false, 'sorted_desc' => false, 'blocks' => [
                ['min' => 5.0, 'max' => 9.0, 'sum' => 30.0, 'count' => 4, 'other' => 6],
            ]]]]
        );

        // Splice an unknown section BEFORE the STAT one and refresh the CRC:
        // a decoder that fails to skip unknown tags will misparse STAT.
        $body = substr($payload, 16);
        $statPos = strpos($body, 'STAT');
        $unknown = 'ZZZZ'.pack('V', 13).str_repeat("\xAB", 13);
        $body = substr($body, 0, $statPos).$unknown.substr($body, $statPos);

        $header = substr($payload, 0, 12).pack('V', crc32($body));
        $decoded = ReaderIndex::decode($header.$body);

        $stats = $decoded->columnStats($entry, 1);
        $this->assertNotNull($stats);
        $this->assertSame(4, $stats['blocks'][0]['count']);
    }

    public function test_decoder_rejects_block_count_mismatch(): void
    {
        $entry = 'xl/worksheets/sheet1.xml';
        // 2 sync points => 3 blocks expected; supply only 1.
        $payload = RandomAccessIndex::encode(
            100,
            [['entry' => $entry, 'total_rows' => 300, 'sheet_crc32' => 7, 'sync_points' => [
                ['row' => 101, 'comp_offset' => 1000, 'uncomp_offset' => 5000],
                ['row' => 201, 'comp_offset' => 2000, 'uncomp_offset' => 10000],
            ]]],
            [$entry => [['col' => 1, 'sorted_asc' => true, 'sorted_desc' => false, 'blocks' => [
                ['min' => 1.0, 'max' => 2.0, 'sum' => 3.0, 'count' => 2, 'other' => 0],
            ]]]]
        );

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/block count/');
        ReaderIndex::decode($payload);
    }

    public function test_block_ranges_align_with_sync_points(): void
    {
        $entry = 'xl/worksheets/sheet1.xml';
        $payload = RandomAccessIndex::encode(100, [
            ['entry' => $entry, 'total_rows' => 253, 'sheet_crc32' => 42, 'sync_points' => [
                ['row' => 101, 'comp_offset' => 1000, 'uncomp_offset' => 5000],
                ['row' => 201, 'comp_offset' => 2000, 'uncomp_offset' => 10000],
            ]],
        ]);

        $ranges = ReaderIndex::decode($payload)->blockRanges($entry);

        $this->assertCount(3, $ranges);
        $this->assertSame(['first_row' => 1, 'last_row' => 100, 'comp_offset' => null, 'uncomp_offset' => null, 'start_row_at_offset' => null], $ranges[0]);
        $this->assertSame(['first_row' => 101, 'last_row' => 200, 'comp_offset' => 1000, 'uncomp_offset' => 5000, 'start_row_at_offset' => 101], $ranges[1]);
        $this->assertSame(['first_row' => 201, 'last_row' => 253, 'comp_offset' => 2000, 'uncomp_offset' => 10000, 'start_row_at_offset' => 201], $ranges[2]);
    }

    // ---------------------------------------------------------------
    // Writer integration (real files)
    // ---------------------------------------------------------------

    public function test_with_column_stats_produces_correct_zone_maps(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([1, 2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'mixed']);

        // Col 1: strictly ascending ints 1..250.
        // Col 2: numerics with holes — nulls and plain strings land in
        // `other`; a numeric STRING and a bool must be folded in as
        // numbers (over-inclusion keeps pruning sound).
        for ($i = 1; $i <= 250; $i++) {
            $mixed = match (true) {
                $i % 10 === 0 => null,
                $i % 7 === 0 => 'text-'.$i,
                $i === 41 => '9000',      // numeric string -> 9000.0
                $i === 43 => true,        // bool -> 1.0
                default => $i * 2.5,
            };
            $writer->writeRow([$i, $mixed]);
        }
        $writer->finishFile();

        $payload = $this->readIndexPayload();
        $this->assertSame(2, ord($payload[4])); // version stays 2 (TLV is trailing)

        $decoded = ReaderIndex::decode($payload);
        $entry = 'xl/worksheets/sheet1.xml';

        $syncCount = count($decoded->syncPoints($entry));
        $this->assertGreaterThan(0, $syncCount);

        $id = $decoded->columnStats($entry, 1);
        $this->assertNotNull($id);
        $this->assertTrue($id['sorted_asc']);
        $this->assertFalse($id['sorted_desc']);
        $this->assertCount($syncCount + 1, $id['blocks']);

        // First block covers rows 1..101 including the header, which the
        // writer folds into block 0 (as `other` for a text header) so
        // zone-map pruning stays consistent with the full-scan path —
        // rowsWhere() can match a numeric-looking header on both paths.
        $first = $id['blocks'][0];
        $this->assertSame(1.0, $first['min']);
        $this->assertSame(100.0, $first['max']);
        $this->assertSame(100, $first['count']);
        $this->assertSame(1, $first['other']); // the 'id' header cell

        // Whole-sheet aggregate folds exactly to n(n+1)/2.
        $totalSum = array_sum(array_column($id['blocks'], 'sum'));
        $totalCount = array_sum(array_column($id['blocks'], 'count'));
        $this->assertSame((float) (250 * 251 / 2), $totalSum);
        $this->assertSame(250, $totalCount);

        $mixed = $decoded->columnStats($entry, 2);
        $this->assertNotNull($mixed);
        $this->assertFalse($mixed['sorted_asc']); // 9000 at i=41 then 1.0 at i=43
        $blockOfRow41 = $mixed['blocks'][0];      // i=41 lands in the first block
        $this->assertSame(9000.0, $blockOfRow41['max']);
        $this->assertSame(1.0, $blockOfRow41['min']); // the bool at i=43

        // The file must still be a perfectly normal XLSX.
        $zip = new ZipArchive();
        $this->assertTrue($zip->open($this->testFile));
        $this->assertNotFalse($zip->locateName('xl/worksheets/sheet1.xml'));
        $zip->close();
    }

    public function test_multi_sheet_stats_are_tracked_per_sheet(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([1]); // implies withRandomAccessIndex()
        $writer->startFile(['v']);

        for ($i = 1; $i <= 10; $i++) {
            $writer->writeRow([$i]);          // ascending
        }
        $writer->newSheet('Second');
        for ($i = 10; $i >= 1; $i--) {
            $writer->writeRow([$i]);          // descending
        }
        $writer->finishFile();

        $decoded = ReaderIndex::decode($this->readIndexPayload());

        $s1 = $decoded->columnStats('xl/worksheets/sheet1.xml', 1);
        $s2 = $decoded->columnStats('xl/worksheets/sheet2.xml', 1);

        $this->assertTrue($s1['sorted_asc']);
        $this->assertFalse($s1['sorted_desc']);
        $this->assertFalse($s2['sorted_asc']);
        $this->assertTrue($s2['sorted_desc']);

        // Small sheets: no sync points, exactly one (tail) block each.
        $this->assertCount(1, $s1['blocks']);
        $this->assertSame(1.0, $s1['blocks'][0]['min']);
        $this->assertSame(10.0, $s1['blocks'][0]['max']);
        $this->assertSame(10, $s1['blocks'][0]['count']);
    }

    public function test_datetime_values_are_folded_as_excel_serials(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([1]);
        $writer->startFile(['when']);

        $early = new \DateTimeImmutable('2020-01-01 00:00:00', new \DateTimeZone('UTC'));
        $late = new \DateTimeImmutable('2026-06-15 12:00:00', new \DateTimeZone('UTC'));
        $writer->writeRow([$late]);
        $writer->writeRow([$early]);
        $writer->finishFile();

        $decoded = ReaderIndex::decode($this->readIndexPayload());
        $stats = $decoded->columnStats('xl/worksheets/sheet1.xml', 1);

        $expectedEarly = (float) (($early->getTimestamp() - SinkableXlsxWriter::EXCEL_EPOCH_TIMESTAMP) / 86400);
        $expectedLate = (float) (($late->getTimestamp() - SinkableXlsxWriter::EXCEL_EPOCH_TIMESTAMP) / 86400);

        $this->assertSame($expectedEarly, $stats['blocks'][0]['min']);
        $this->assertSame($expectedLate, $stats['blocks'][0]['max']);
        $this->assertSame(2, $stats['blocks'][0]['count']);
    }

    public function test_with_column_stats_validates_input(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(XlsxStreamException::class);
        $writer->withColumnStats([0]);
    }

    public function test_with_column_stats_rejected_after_start(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a']);

        try {
            $this->expectException(XlsxStreamException::class);
            $writer->withColumnStats([1]);
        } finally {
            // Leave no half-open writer behind; the sink owns a handle.
            $writer->writeRow([1]);
            $writer->finishFile();
        }
    }

    private function readIndexPayload(): string
    {
        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        return $cd->readEntry($source, RandomAccessIndex::ENTRY_PATH);
    }
}

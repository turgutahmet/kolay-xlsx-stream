<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\RandomAccessIndex as WriterIndex;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * KXSI "SCRC" TLV section — per-sheet running CRC32 of the sheet's
 * uncompressed bytes at each sync point.
 *
 * Three layers are pinned here:
 *   - plumbing: encoder/decoder round trip, alignment invariants,
 *     truncation guards, section-order independence through the
 *     unknown-tag skip path (the version byte stays 2);
 *   - semantics: for a file produced by the real writer, each SCRC value
 *     is INDEPENDENTLY recomputed as crc32 of the inflated sheet's first
 *     uncomp_offset bytes — proving the values mean what the spec says,
 *     not merely that the writer's bytes survive the decoder;
 *   - compatibility: a stats-bearing + SCRC-bearing sidecar still decodes
 *     every core field, and the flags byte is enforced as must-understand.
 */
class ScrcTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-scrc-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    // ------------------------------------------------------------------
    // Encoder
    // ------------------------------------------------------------------

    public function test_encoder_emits_scrc_section_bytes(): void
    {
        $payload = WriterIndex::encode(
            100,
            [$this->sheetSection('xl/worksheets/sheet1.xml', 250, [
                ['row' => 102, 'comp_offset' => 10, 'uncomp_offset' => 100],
                ['row' => 202, 'comp_offset' => 20, 'uncomp_offset' => 200],
            ])],
            [],
            ['xl/worksheets/sheet1.xml' => [0xDEADBEEF, 0x00C0FFEE]]
        );

        // Core body for one sheet with 2 sync points:
        // 2 (path len) + 24 (path) + 4 + 4 + 4 + 2*24 = 86 bytes.
        $body = substr($payload, 16);
        $tlv = substr($body, 86);

        $this->assertSame('SCRC', substr($tlv, 0, 4));
        $this->assertSame(12, unpack('V', substr($tlv, 4, 4))[1]); // count + 2 crcs
        $this->assertSame(2, unpack('V', substr($tlv, 8, 4))[1]);
        $this->assertSame(0xDEADBEEF, unpack('V', substr($tlv, 12, 4))[1]);
        $this->assertSame(0x00C0FFEE, unpack('V', substr($tlv, 16, 4))[1]);
        $this->assertSame(strlen($body), 86 + strlen($tlv)); // nothing after SCRC
    }

    public function test_encoder_omits_scrc_when_no_crcs_passed(): void
    {
        // Bytes of a CRC-less encode stay identical to the v3.1 output —
        // callers that never pass the argument are unaffected.
        $sheets = [$this->sheetSection('xl/worksheets/sheet1.xml', 50, [])];

        $this->assertSame(
            WriterIndex::encode(100, $sheets),
            WriterIndex::encode(100, $sheets, [], [])
        );
        $this->assertStringNotContainsString('SCRC', WriterIndex::encode(100, $sheets));
    }

    public function test_encoder_rejects_crc_list_misaligned_with_sync_points(): void
    {
        $this->expectException(\InvalidArgumentException::class);

        WriterIndex::encode(
            100,
            [$this->sheetSection('xl/worksheets/sheet1.xml', 250, [
                ['row' => 102, 'comp_offset' => 10, 'uncomp_offset' => 100],
                ['row' => 202, 'comp_offset' => 20, 'uncomp_offset' => 200],
            ])],
            [],
            ['xl/worksheets/sheet1.xml' => [0xDEADBEEF]] // 1 CRC for 2 sync points
        );
    }

    // ------------------------------------------------------------------
    // Decoder
    // ------------------------------------------------------------------

    public function test_scrc_round_trips_through_decoder(): void
    {
        $payload = WriterIndex::encode(
            100,
            [
                $this->sheetSection('xl/worksheets/sheet1.xml', 250, [
                    ['row' => 102, 'comp_offset' => 10, 'uncomp_offset' => 100],
                    ['row' => 202, 'comp_offset' => 20, 'uncomp_offset' => 200],
                ]),
                $this->sheetSection('xl/worksheets/sheet2.xml', 50, []),
            ],
            [],
            [
                'xl/worksheets/sheet1.xml' => [0xDEADBEEF, 0x00C0FFEE],
                'xl/worksheets/sheet2.xml' => [],
            ]
        );

        $index = ReaderIndex::decode($payload);

        $this->assertSame([0xDEADBEEF, 0x00C0FFEE], $index->syncPointCrcs('xl/worksheets/sheet1.xml'));
        $this->assertSame([], $index->syncPointCrcs('xl/worksheets/sheet2.xml'));
        // Core fields untouched by the extra section.
        $this->assertSame(250, $index->totalRows('xl/worksheets/sheet1.xml'));
        $this->assertCount(2, $index->syncPoints('xl/worksheets/sheet1.xml'));
    }

    public function test_scrc_absent_yields_empty_list(): void
    {
        $payload = WriterIndex::encode(100, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 250, [
                ['row' => 102, 'comp_offset' => 10, 'uncomp_offset' => 100],
            ]),
        ]);

        $index = ReaderIndex::decode($payload);

        $this->assertSame([], $index->syncPointCrcs('xl/worksheets/sheet1.xml'));
    }

    public function test_decoder_rejects_scrc_count_mismatching_sync_count(): void
    {
        // Sheet has 2 sync points but the crafted SCRC record claims 1.
        // The alignment invariant (count == sync_count) must reject it.
        $core = $this->corePayload(2);
        $scrc = pack('V', 1).pack('V', 0xDEADBEEF);
        $payload = $this->withTlvSections($core, ['SCRC'.pack('V', strlen($scrc)).$scrc]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('does not match sync count');

        ReaderIndex::decode($payload);
    }

    public function test_decoder_rejects_truncated_scrc_value_list(): void
    {
        // count matches sync_count (2) but only one CRC value is present.
        $core = $this->corePayload(2);
        $scrc = pack('V', 2).pack('V', 0xDEADBEEF);
        $payload = $this->withTlvSections($core, ['SCRC'.pack('V', strlen($scrc)).$scrc]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('truncated SCRC value list');

        ReaderIndex::decode($payload);
    }

    public function test_decoder_rejects_truncated_scrc_sheet_header(): void
    {
        // Empty SCRC payload for a 1-sheet body — not even the uint32
        // count fits.
        $core = $this->corePayload(1);
        $payload = $this->withTlvSections($core, ['SCRC'.pack('V', 0)]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('truncated SCRC section sheet header');

        ReaderIndex::decode($payload);
    }

    public function test_tlv_sections_decode_in_any_order_with_unknown_tags(): void
    {
        // Round-trip of the extension contract: unknown tags MUST be
        // skipped and section order MUST NOT matter, all under version
        // byte 2. Every permutation of {unknown, SCRC, STAT} decodes the
        // same core fields, stats, and CRCs.
        $core = $this->corePayload(1); // 1 sync point => 2 stat blocks
        $scrcBody = pack('V', 1).pack('V', 0xDEADBEEF);
        $scrc = 'SCRC'.pack('V', strlen($scrcBody)).$scrcBody;

        $statBody = pack('v', 1).pack('vCV', 1, 0x01, 2)
            .pack('eeeVV', 1.0, 5.0, 9.0, 3, 0)
            .pack('eeeVV', 6.0, 8.0, 14.0, 2, 1);
        $stat = 'STAT'.pack('V', strlen($statBody)).$statBody;

        $unknown = 'ZZZZ'.pack('V', 7).'garbage';

        $permutations = [
            [$unknown, $scrc, $stat],
            [$stat, $scrc, $unknown],
            [$scrc, $unknown, $stat],
            [$stat, $unknown, $scrc],
        ];

        foreach ($permutations as $i => $sections) {
            $index = ReaderIndex::decode($this->withTlvSections($core, $sections));

            $this->assertSame(250, $index->totalRows('xl/worksheets/sheet1.xml'), "permutation {$i}");
            $this->assertCount(1, $index->syncPoints('xl/worksheets/sheet1.xml'), "permutation {$i}");
            $this->assertSame([0xDEADBEEF], $index->syncPointCrcs('xl/worksheets/sheet1.xml'), "permutation {$i}");

            $stats = $index->columnStats('xl/worksheets/sheet1.xml', 1);
            $this->assertNotNull($stats, "permutation {$i}");
            $this->assertTrue($stats['sorted_asc'], "permutation {$i}");
            $this->assertSame(5.0, $stats['blocks'][0]['max'], "permutation {$i}");
        }
    }

    public function test_nonzero_flags_byte_is_rejected(): void
    {
        // The flags byte is a must-understand bitmask — a set bit this
        // decoder doesn't implement makes the whole sidecar unusable.
        $payload = WriterIndex::encode(100, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 50, []),
        ]);
        $payload[5] = "\x01";

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('flags byte');

        ReaderIndex::decode($payload);
    }

    // ------------------------------------------------------------------
    // Semantics — the CRC values mean what the spec says
    // ------------------------------------------------------------------

    public function test_written_file_scrc_matches_crc_of_uncompressed_prefix(): void
    {
        // THE semantics test: write a real indexed file, then for every
        // sync point independently recompute crc32 of the inflated
        // sheet's first uncomp_offset bytes and compare with the SCRC
        // value. This proves "value k pins the uncompressed prefix
        // [0, uncomp_offset_k)" — not just that bytes round-trip.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'name']);
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([$i, "user-{$i}"]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $index = ReaderIndex::decode($cd->readEntry($source, ReaderIndex::ENTRY_PATH));

        $entry = 'xl/worksheets/sheet1.xml';
        $sheetXml = $cd->readEntry($source, $entry);
        $points = $index->syncPoints($entry);
        $crcs = $index->syncPointCrcs($entry);

        $this->assertNotEmpty($points);
        $this->assertSameSize($points, $crcs);

        foreach ($points as $k => $sp) {
            $this->assertSame(
                crc32(substr($sheetXml, 0, $sp['uncomp_offset'])),
                $crcs[$k],
                "SCRC value {$k} does not match crc32 of the first {$sp['uncomp_offset']} uncompressed bytes"
            );
        }

        // Sanity anchor: the whole-sheet CRC pin still matches the same
        // derivation applied to the full entry.
        $this->assertSame(crc32($sheetXml), $index->sheetCrc32($entry));
    }

    public function test_multi_sheet_written_file_scrc_verifies_per_sheet(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 50);
        $writer->setBufferFlushInterval(50);
        $writer->startFile(['a']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->newSheet('Other', ['b']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i * 2]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $index = ReaderIndex::decode($cd->readEntry($source, ReaderIndex::ENTRY_PATH));

        foreach (['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'] as $entry) {
            $sheetXml = $cd->readEntry($source, $entry);
            $points = $index->syncPoints($entry);
            $crcs = $index->syncPointCrcs($entry);

            $this->assertNotEmpty($points, $entry);
            $this->assertSameSize($points, $crcs, $entry);

            foreach ($points as $k => $sp) {
                $this->assertSame(
                    crc32(substr($sheetXml, 0, $sp['uncomp_offset'])),
                    $crcs[$k],
                    "{$entry} SCRC value {$k}"
                );
            }
        }
    }

    public function test_indexed_file_without_sync_points_still_carries_scrc_section(): void
    {
        // Too few rows to ever trigger a sync flush — the SCRC section is
        // still emitted (count = 0) so its per-sheet records stay aligned
        // with the core body.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 10000);
        $writer->startFile(['id']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $payload = $cd->readEntry($source, ReaderIndex::ENTRY_PATH);

        $this->assertStringContainsString('SCRC', $payload);

        $index = ReaderIndex::decode($payload);
        $this->assertSame([], $index->syncPoints('xl/worksheets/sheet1.xml'));
        $this->assertSame([], $index->syncPointCrcs('xl/worksheets/sheet1.xml'));
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    /**
     * Core-only payload for one sheet with $syncCount evenly spaced sync
     * points — the canvas the TLV crafting helpers draw on.
     */
    private function corePayload(int $syncCount): string
    {
        $points = [];
        for ($k = 1; $k <= $syncCount; $k++) {
            $points[] = [
                'row' => $k * 100 + 2,
                'comp_offset' => $k * 1000,
                'uncomp_offset' => $k * 10000,
            ];
        }

        return WriterIndex::encode(100, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 250, $points),
        ]);
    }

    /**
     * Append raw TLV section bytes to a payload's body and refresh the
     * header CRC so the structural decoder accepts the crafted result.
     *
     * @param  list<string>  $sections  fully framed TLV byte strings
     */
    private function withTlvSections(string $payload, array $sections): string
    {
        $body = substr($payload, 16).implode('', $sections);

        return substr($payload, 0, 12).pack('V', crc32($body)).$body;
    }

    /**
     * @param  list<array{row: int, comp_offset: int, uncomp_offset: int}>  $points
     * @return array{entry: string, total_rows: int, sheet_crc32: int, sync_points: list<array{row: int, comp_offset: int, uncomp_offset: int}>}
     */
    private function sheetSection(string $entry, int $totalRows, array $points): array
    {
        return [
            'entry' => $entry,
            'total_rows' => $totalRows,
            'sheet_crc32' => 0,
            'sync_points' => $points,
        ];
    }
}

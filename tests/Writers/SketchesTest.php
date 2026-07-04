<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sketches\HyperLogLog;
use Kolay\XlsxStream\Sketches\TDigest;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\RandomAccessIndex as WriterIndex;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * KXSI "TDIG" / "CHLL" TLV sections — per-column sketch payloads for
 * sidecar-only quantiles and distinct counts.
 *
 * Mirrors ScrcTest's three layers:
 *   - plumbing: encoder/decoder round trip through the shared section
 *     frame, byte-identical output for sketch-less callers, positional
 *     per-sheet alignment, truncation/1-based guards, unknown-tag order
 *     independence (version byte stays 2);
 *   - semantics: for a file produced by the real writer, sketch answers
 *     are INDEPENDENTLY verified against the raw data — including the
 *     header-row exclusion this format deliberately pins (the opposite
 *     of STAT's header fold) and the CHLL canonicalization rule;
 *   - orthogonality: sketches without stats and stats without sketches
 *     each decode with the other surface absent, not broken.
 */
class SketchesTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-sketch-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    // ------------------------------------------------------------------
    // Encoder
    // ------------------------------------------------------------------

    public function test_encoder_emits_tdig_and_chll_section_bytes(): void
    {
        $digestPayload = $this->digestFor([1.0, 2.0, 3.0]);
        $hllPayload = $this->hllFor(['a', 'b']);

        $payload = WriterIndex::encode(
            100,
            [$this->sheetSection('xl/worksheets/sheet1.xml', 250, [])],
            [],
            [],
            ['xl/worksheets/sheet1.xml' => [2 => $digestPayload]],
            ['xl/worksheets/sheet1.xml' => [2 => $hllPayload]]
        );

        // Core body for one sheet with 0 sync points:
        // 2 (path len) + 24 (path) + 4 + 4 + 4 = 38 bytes.
        $body = substr($payload, 16);
        $tlv = substr($body, 38);

        $this->assertSame('TDIG', substr($tlv, 0, 4));
        $tdigLen = unpack('V', substr($tlv, 4, 4))[1];
        $this->assertSame(2 + 6 + strlen($digestPayload), $tdigLen);
        $this->assertSame(1, unpack('v', substr($tlv, 8, 2))[1]);   // tracked_column_count
        $this->assertSame(2, unpack('v', substr($tlv, 10, 2))[1]);  // column
        $this->assertSame(strlen($digestPayload), unpack('V', substr($tlv, 12, 4))[1]);
        $this->assertSame($digestPayload, substr($tlv, 16, strlen($digestPayload)));

        $chll = substr($tlv, 8 + $tdigLen);
        $this->assertSame('CHLL', substr($chll, 0, 4));
        $chllLen = unpack('V', substr($chll, 4, 4))[1];
        $this->assertSame(2 + 6 + strlen($hllPayload), $chllLen);
        $this->assertSame($hllPayload, substr($chll, 16, strlen($hllPayload)));

        // Nothing after CHLL.
        $this->assertSame(strlen($body), 38 + 8 + $tdigLen + 8 + $chllLen);
    }

    public function test_encoder_output_byte_identical_without_sketches(): void
    {
        // Callers that never pass the new arguments keep the exact
        // pre-sketch bytes — same guarantee SCRC gave when it landed.
        $sheets = [$this->sheetSection('xl/worksheets/sheet1.xml', 50, [])];

        $this->assertSame(
            WriterIndex::encode(100, $sheets, [], []),
            WriterIndex::encode(100, $sheets, [], [], [], [])
        );
        $this->assertStringNotContainsString('TDIG', WriterIndex::encode(100, $sheets));
        $this->assertStringNotContainsString('CHLL', WriterIndex::encode(100, $sheets));
    }

    public function test_encoder_emits_count_zero_record_for_sheet_without_sketches(): void
    {
        // Positional alignment: a 2-sheet body where only sheet 2 has
        // sketches still writes a count-0 record for sheet 1.
        $payload = WriterIndex::encode(
            100,
            [
                $this->sheetSection('xl/worksheets/sheet1.xml', 10, []),
                $this->sheetSection('xl/worksheets/sheet2.xml', 10, []),
            ],
            [],
            [],
            ['xl/worksheets/sheet2.xml' => [1 => $this->digestFor([5.0])]],
            []
        );

        $index = ReaderIndex::decode($payload);

        $this->assertNull($index->columnDigest('xl/worksheets/sheet1.xml', 1));
        $this->assertNotNull($index->columnDigest('xl/worksheets/sheet2.xml', 1));
        $this->assertSame([], $index->digestColumns('xl/worksheets/sheet1.xml'));
        $this->assertSame([1], $index->digestColumns('xl/worksheets/sheet2.xml'));
    }

    // ------------------------------------------------------------------
    // Decoder
    // ------------------------------------------------------------------

    public function test_sketches_round_trip_through_decoder(): void
    {
        $digestPayload = $this->digestFor([10.0, 20.0, 30.0, 40.0]);
        $hllPayload = $this->hllFor(['x', 'y', 'z']);

        $payload = WriterIndex::encode(
            100,
            [
                $this->sheetSection('xl/worksheets/sheet1.xml', 250, [
                    ['row' => 102, 'comp_offset' => 10, 'uncomp_offset' => 100],
                ]),
                $this->sheetSection('xl/worksheets/sheet2.xml', 50, []),
            ],
            [],
            [],
            [
                'xl/worksheets/sheet1.xml' => [2 => $digestPayload, 5 => $this->digestFor([7.5])],
                'xl/worksheets/sheet2.xml' => [],
            ],
            [
                'xl/worksheets/sheet1.xml' => [2 => $hllPayload],
                'xl/worksheets/sheet2.xml' => [],
            ]
        );

        $index = ReaderIndex::decode($payload);

        $digest = $index->columnDigest('xl/worksheets/sheet1.xml', 2);
        $this->assertNotNull($digest);
        $this->assertSame(4, $digest->count());
        $this->assertSame(10.0, $digest->quantile(0.0));
        $this->assertSame(40.0, $digest->quantile(1.0));
        $this->assertSame([2, 5], $index->digestColumns('xl/worksheets/sheet1.xml'));

        $hll = $index->columnHll('xl/worksheets/sheet1.xml', 2);
        $this->assertNotNull($hll);
        $this->assertSame(3, $hll->count());
        $this->assertSame([2], $index->hllColumns('xl/worksheets/sheet1.xml'));

        // Untracked column / sheet => null, and core fields untouched.
        $this->assertNull($index->columnDigest('xl/worksheets/sheet1.xml', 3));
        $this->assertNull($index->columnHll('xl/worksheets/sheet2.xml', 2));
        $this->assertSame(250, $index->totalRows('xl/worksheets/sheet1.xml'));
        $this->assertCount(1, $index->syncPoints('xl/worksheets/sheet1.xml'));
    }

    public function test_sketches_absent_yield_null(): void
    {
        $index = ReaderIndex::decode(WriterIndex::encode(100, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 250, []),
        ]));

        $this->assertNull($index->columnDigest('xl/worksheets/sheet1.xml', 1));
        $this->assertNull($index->columnHll('xl/worksheets/sheet1.xml', 1));
        $this->assertSame([], $index->digestColumns('xl/worksheets/sheet1.xml'));
        $this->assertSame([], $index->hllColumns('xl/worksheets/sheet1.xml'));
    }

    public function test_decoder_rejects_truncated_sketch_sheet_header(): void
    {
        // Empty TDIG payload for a 1-sheet body — not even the uint16
        // tracked_column_count fits.
        $payload = $this->withTlvSections($this->corePayload(), ['TDIG'.pack('V', 0)]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('truncated TDIG section sheet header');

        ReaderIndex::decode($payload);
    }

    public function test_decoder_rejects_truncated_sketch_column_header(): void
    {
        // Record claims 1 column but only 3 of the 6 header bytes exist.
        $section = pack('v', 1).pack('v', 2).chr(0);
        $payload = $this->withTlvSections($this->corePayload(), ['CHLL'.pack('V', strlen($section)).$section]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('truncated CHLL column header');

        ReaderIndex::decode($payload);
    }

    public function test_decoder_rejects_sketch_payload_overrunning_section(): void
    {
        // payload_len claims 100 bytes; only 4 remain in the section.
        $section = pack('v', 1).pack('vV', 2, 100).'abcd';
        $payload = $this->withTlvSections($this->corePayload(), ['TDIG'.pack('V', strlen($section)).$section]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('truncated TDIG sketch payload');

        ReaderIndex::decode($payload);
    }

    public function test_decoder_rejects_zero_column_index(): void
    {
        $digest = $this->digestFor([1.0]);
        $section = pack('v', 1).pack('vV', 0, strlen($digest)).$digest;
        $payload = $this->withTlvSections($this->corePayload(), ['TDIG'.pack('V', strlen($section)).$section]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('must be 1-based');

        ReaderIndex::decode($payload);
    }

    public function test_corrupt_sketch_payload_is_rejected_at_access(): void
    {
        // Framing is valid; the sketch bytes themselves are garbage.
        // decode() succeeds (payloads parse lazily); the accessor must
        // convert the sketch's own rejection into a read exception.
        $garbage = 'not a digest';
        $section = pack('v', 1).pack('vV', 2, strlen($garbage)).$garbage;
        $payload = $this->withTlvSections($this->corePayload(), ['TDIG'.pack('V', strlen($section)).$section]);

        $index = ReaderIndex::decode($payload);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessage('TDIG sketch');

        $index->columnDigest('xl/worksheets/sheet1.xml', 2);
    }

    public function test_tlv_sections_decode_in_any_order_with_unknown_tags(): void
    {
        // The §4 compatibility contract, extended to the new tags:
        // every permutation of {unknown, TDIG, CHLL, SCRC} under
        // version byte 2 decodes identical surfaces.
        $core = $this->corePayload(1);

        $scrcBody = pack('V', 1).pack('V', 0xDEADBEEF);
        $scrc = 'SCRC'.pack('V', strlen($scrcBody)).$scrcBody;

        $digestPayload = $this->digestFor([1.0, 9.0]);
        $tdigBody = pack('v', 1).pack('vV', 1, strlen($digestPayload)).$digestPayload;
        $tdig = 'TDIG'.pack('V', strlen($tdigBody)).$tdigBody;

        $hllPayload = $this->hllFor(['p', 'q']);
        $chllBody = pack('v', 1).pack('vV', 1, strlen($hllPayload)).$hllPayload;
        $chll = 'CHLL'.pack('V', strlen($chllBody)).$chllBody;

        $unknown = 'ZZZZ'.pack('V', 7).'garbage';

        $permutations = [
            [$unknown, $tdig, $chll, $scrc],
            [$chll, $scrc, $tdig, $unknown],
            [$tdig, $unknown, $scrc, $chll],
            [$scrc, $chll, $unknown, $tdig],
        ];

        foreach ($permutations as $i => $sections) {
            $index = ReaderIndex::decode($this->withTlvSections($core, $sections));

            $this->assertSame(250, $index->totalRows('xl/worksheets/sheet1.xml'), "permutation {$i}");
            $this->assertSame([0xDEADBEEF], $index->syncPointCrcs('xl/worksheets/sheet1.xml'), "permutation {$i}");
            $this->assertSame(2, $index->columnDigest('xl/worksheets/sheet1.xml', 1)?->count(), "permutation {$i}");
            $this->assertSame(2, $index->columnHll('xl/worksheets/sheet1.xml', 1)?->count(), "permutation {$i}");
        }
    }

    // ------------------------------------------------------------------
    // Writer semantics — sketches mean what the spec says
    // ------------------------------------------------------------------

    public function test_written_file_sketches_verify_against_raw_data(): void
    {
        // THE semantics test: write a real file, then rebuild both
        // sketches INDEPENDENTLY from the same raw values under the
        // spec's rules and compare serialized bytes. Proves the writer
        // feeds exactly the specified value populations (numeric rule
        // for TDIG, canonical strings for CHLL, header excluded), not
        // merely that bytes survive the round trip.
        //
        // Byte equality for the t-digest additionally requires the same
        // FEED CADENCE (centroid boundaries depend on when the internal
        // buffer folds). 500 rows stay under both the writer's staging
        // threshold and the digest's buffer limit, so writer and oracle
        // each fold exactly once, at serialize(). The HLL is cadence-
        // independent — its bytes match under any feeding order.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnSketches([1, 2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'label']);

        $expectDigest = new TDigest();
        $expectHll1 = new HyperLogLog();
        $expectHll2 = new HyperLogLog();
        for ($i = 1; $i <= 500; $i++) {
            $amount = $i % 10 === 0 ? 'n/a' : $i * 0.5;
            $writer->writeRow([$amount, 'label-'.($i % 40)]);

            if (! is_string($amount)) {
                $expectDigest->add($amount);
                $expectHll1->add((string) $amount);
            } else {
                $expectHll1->add($amount);
            }
            $expectHll2->add('label-'.($i % 40));
        }
        $writer->finishFile();

        $index = $this->decodeSidecar();
        $entry = 'xl/worksheets/sheet1.xml';

        $this->assertSame($expectDigest->serialize(), $index->columnDigest($entry, 1)?->serialize());
        $this->assertSame($expectHll1->serialize(), $index->columnHll($entry, 1)?->serialize());
        $this->assertSame($expectHll2->serialize(), $index->columnHll($entry, 2)?->serialize());

        // Sanity on the derived answers: 450 numeric values spanning
        // 0.5..249.5 feed the digest (every 10th row is 'n/a', so
        // i=500 is text); col 2 has exactly 40 labels.
        $this->assertSame(450, $index->columnDigest($entry, 1)->count());
        $this->assertSame(0.5, $index->columnDigest($entry, 1)->quantile(0.0));
        $this->assertSame(249.5, $index->columnDigest($entry, 1)->quantile(1.0));
        // The byte-equality above is the oracle; the derived count is
        // an estimate (here 39 — one register collision among the 40).
        $this->assertEqualsWithDelta(40, $index->columnHll($entry, 2)->count(), 2);
    }

    public function test_header_row_is_excluded_from_sketches(): void
    {
        // Numeric-looking header over a constant column: STAT must fold
        // it (pruning soundness), sketches must NOT (estimation bias).
        // With the header excluded: 1 distinct value, every quantile 7.
        // A header leak would show 2 distinct and max 999999.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnSketches([1]);
        $writer->withColumnStats([1]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['999999']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow([7]);
        }
        $writer->finishFile();

        $index = $this->decodeSidecar();
        $entry = 'xl/worksheets/sheet1.xml';

        $digest = $index->columnDigest($entry, 1);
        $this->assertSame(300, $digest->count(), 'header row leaked into the t-digest');
        $this->assertSame(7.0, $digest->quantile(1.0));
        $this->assertSame(1, $index->columnHll($entry, 1)->count(), 'header row leaked into the HLL');

        // ...while the SAME file's STAT block 0 does carry the header
        // fold — the asymmetry is deliberate and both live together.
        $stats = $index->columnStats($entry, 1);
        $this->assertSame(999999.0, $stats['blocks'][0]['max']);
    }

    public function test_hll_canonicalization_collapses_forms_that_render_identically(): void
    {
        // int 7 and string '7' both render <v>7</v> -> one canonical
        // form. '1.50' (string, preserved) vs float 1.5 -> distinct.
        // bool true renders '1' -> collapses with int 1. null and ''
        // are not values. DateTime hashes as its Excel serial string.
        $noon = new \DateTimeImmutable('2024-01-15 12:00:00', new \DateTimeZone('UTC'));
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnSketches([1]);
        $writer->startFile(['v']);
        foreach ([7, '7', '1.50', 1.5, true, 1, null, '', $noon, $noon] as $value) {
            $writer->writeRow([$value]);
        }
        $writer->finishFile();

        // Expected canonical forms: '7', '1.50', '1.5', '1', serial.
        $this->assertSame(5, $this->decodeSidecar()->columnHll('xl/worksheets/sheet1.xml', 1)->count());
    }

    public function test_multi_sheet_written_file_carries_per_sheet_sketches(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 50);
        $writer->withColumnSketches([1]);
        $writer->setBufferFlushInterval(50);
        $writer->startFile(['a']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i]);          // sheet 1: 1..100
        }
        $writer->newSheet('Other', ['b']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i % 10]);     // sheet 2: 10 distinct values
        }
        $writer->finishFile();

        $index = $this->decodeSidecar();

        $s1 = $index->columnDigest('xl/worksheets/sheet1.xml', 1);
        $this->assertSame(100, $s1->count());
        $this->assertSame(100.0, $s1->quantile(1.0));
        $this->assertSame(100, $index->columnHll('xl/worksheets/sheet1.xml', 1)->count());

        $s2 = $index->columnDigest('xl/worksheets/sheet2.xml', 1);
        $this->assertSame(200, $s2->count());
        $this->assertSame(9.0, $s2->quantile(1.0));
        $this->assertSame(10, $index->columnHll('xl/worksheets/sheet2.xml', 1)->count());
    }

    public function test_sketches_without_stats_and_stats_without_sketches(): void
    {
        // Orthogonality at the sidecar level: each opt-in emits only
        // its own sections.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnSketches([1]);
        $writer->startFile(['n']);
        for ($i = 1; $i <= 50; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $index = $this->decodeSidecar();
        $this->assertNotNull($index->columnDigest('xl/worksheets/sheet1.xml', 1));
        $this->assertNull($index->columnStats('xl/worksheets/sheet1.xml', 1));

        @unlink($this->testFile);

        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([1]);
        $writer->startFile(['n']);
        for ($i = 1; $i <= 50; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $index = $this->decodeSidecar();
        $this->assertNotNull($index->columnStats('xl/worksheets/sheet1.xml', 1));
        $this->assertNull($index->columnDigest('xl/worksheets/sheet1.xml', 1));
        $this->assertNull($index->columnHll('xl/worksheets/sheet1.xml', 1));
    }

    public function test_with_column_sketches_implies_random_access_index(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnSketches([1]); // no explicit withRandomAccessIndex
        $writer->startFile(['n']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $this->assertTrue($cd->has(ReaderIndex::ENTRY_PATH));
        $source->close();
    }

    public function test_with_column_sketches_argument_guards(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $guard = function (callable $call, string $label): void {
            try {
                $call();
                $this->fail($label.' accepted');
            } catch (XlsxStreamException $e) {
                $this->assertNotSame('', $e->getMessage(), $label);
            }
        };

        $guard(fn () => $writer->withColumnSketches([]), 'empty column list');
        $guard(fn () => $writer->withColumnSketches([0]), 'column 0');

        $writer->startFile(['n']);
        $guard(fn () => $writer->withColumnSketches([1]), 'opt-in after startFile');
        $writer->writeRow([1]);
        $writer->finishFile();
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    /** @param  list<float>  $values */
    private function digestFor(array $values): string
    {
        $digest = new TDigest();
        foreach ($values as $v) {
            $digest->add($v);
        }

        return $digest->serialize();
    }

    /** @param  list<string>  $values */
    private function hllFor(array $values): string
    {
        $hll = new HyperLogLog();
        foreach ($values as $v) {
            $hll->add($v);
        }

        return $hll->serialize();
    }

    private function decodeSidecar(): ReaderIndex
    {
        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $index = ReaderIndex::decode($cd->readEntry($source, ReaderIndex::ENTRY_PATH));
        $source->close();

        return $index;
    }

    /**
     * Core-only payload for one sheet with $syncCount evenly spaced sync
     * points — the canvas the TLV crafting helpers draw on.
     */
    private function corePayload(int $syncCount = 0): string
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

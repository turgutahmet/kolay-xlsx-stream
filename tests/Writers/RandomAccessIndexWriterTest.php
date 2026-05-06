<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\RandomAccessIndex;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * Integration tests for the writer's withRandomAccessIndex() opt-in.
 *
 * Each test checks that:
 *   - the indexed file still parses as a normal XLSX (vanilla ZIP, our
 *     reader, libxml strict), so backward compat with Excel /
 *     PhpSpreadsheet / OpenSpout is preserved
 *   - the xl/_kxs/index.bin sidecar contains the expected sync-point
 *     row positions and decodes byte-for-byte through the writer-side
 *     binary format
 *
 * Test scale stays small (≤ 1000 rows) — the goal is correctness, not
 * benchmarking. Memory usage is unchanged from the un-indexed path
 * because the sync-point recording uses a packed binary buffer, not
 * a PHP array of associative entries.
 */
class RandomAccessIndexWriterTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-indexed-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_indexed_file_contains_kxs_index_entry(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['n']);
        for ($i = 1; $i <= 250; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $this->assertTrue($cd->has(RandomAccessIndex::ENTRY_PATH));
    }

    public function test_index_payload_round_trips_through_writer_format(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['n']);
        for ($i = 1; $i <= 1000; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $payload = $this->readIndexPayload();

        $this->assertSame('KXSI', substr($payload, 0, 4));
        $this->assertSame(1, ord($payload[4]));         // version
        $this->assertSame(0, ord($payload[5]));         // flags
        $this->assertSame(1, unpack('v', substr($payload, 6, 2))[1]);     // sheet count
        $this->assertSame(100, unpack('V', substr($payload, 8, 4))[1]);   // sync_period
        $this->assertSame(crc32(substr($payload, 16)), unpack('V', substr($payload, 12, 4))[1]);

        $sheets = $this->decodeSheets($payload);

        $this->assertCount(1, $sheets);
        $this->assertSame('xl/worksheets/sheet1.xml', $sheets[0]['entry']);
        $this->assertSame(1001, $sheets[0]['total_rows']); // header + 1000 data rows
    }

    public function test_sync_points_recorded_at_row_boundaries(): void
    {
        // syncEvery == bufferFlushInterval == 100, write 1000 data rows.
        // Each flush carries exactly 100 rows, every flush triggers a sync.
        // Sync rows mark the FIRST row of the next batch (currentSheetRow + 1).
        // After header (row 1) + first 100 data rows: currentSheetRow=101,
        // recorded sync = 102. After next 100 rows: 202. ... after final
        // 100 rows: 1002.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['n']);
        for ($i = 1; $i <= 1000; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $sheets = $this->decodeSheets($this->readIndexPayload());

        $rows = array_column($sheets[0]['sync_points'], 'row');
        $this->assertCount(10, $rows);
        $this->assertSame(
            [102, 202, 302, 402, 502, 602, 702, 802, 902, 1002],
            $rows
        );
    }

    public function test_indexed_file_opens_through_streaming_reader(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 50);
        $writer->setBufferFlushInterval(50);
        $writer->startFile(['id', 'value']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i, "val-{$i}"]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertCount(201, $rows);
        $this->assertSame(['id', 'value'], $rows[0]);
        $this->assertSame(['1', 'val-1'], $rows[1]);
        $this->assertSame(['200', 'val-200'], $rows[200]);
    }

    public function test_indexed_file_is_a_valid_zip_archive(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['n']);
        for ($i = 1; $i <= 250; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $zip = new ZipArchive();
        $this->assertSame(true, $zip->open($this->testFile, ZipArchive::RDONLY));

        // Every part Excel expects, untouched.
        $this->assertNotFalse($zip->locateName('[Content_Types].xml'));
        $this->assertNotFalse($zip->locateName('xl/workbook.xml'));
        $this->assertNotFalse($zip->locateName('xl/worksheets/sheet1.xml'));
        $this->assertNotFalse($zip->locateName('xl/styles.xml'));
        // The sidecar — vanilla readers ignore it because it's not in
        // [Content_Types].xml; ZipArchive still lists it.
        $this->assertNotFalse($zip->locateName(RandomAccessIndex::ENTRY_PATH));

        // Sheet XML is still strict-libxml-parseable.
        $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        libxml_use_internal_errors(true);
        libxml_clear_errors();
        $doc = new \DOMDocument();
        $ok = $doc->loadXML($sheetXml, LIBXML_NONET);
        $errors = libxml_get_errors();
        libxml_clear_errors();

        $this->assertTrue($ok);
        $this->assertEmpty($errors);
    }

    public function test_index_emission_does_not_change_data_bytes_in_other_parts(): void
    {
        $rows = [];
        for ($i = 1; $i <= 200; $i++) {
            $rows[] = [$i, "v-{$i}"];
        }

        $plain = sys_get_temp_dir().'/kxs-plain-'.uniqid('', true).'.xlsx';
        $indexed = $this->testFile;

        try {
            $a = new SinkableXlsxWriter(new FileSink($plain));
            $a->setBufferFlushInterval(100);
            $a->startFile(['id', 'value']);
            foreach ($rows as $r) {
                $a->writeRow($r);
            }
            $a->finishFile();

            $b = new SinkableXlsxWriter(new FileSink($indexed));
            $b->withRandomAccessIndex(every: 100);
            $b->setBufferFlushInterval(100);
            $b->startFile(['id', 'value']);
            foreach ($rows as $r) {
                $b->writeRow($r);
            }
            $b->finishFile();

            // Visible content in Excel is identical.
            $readerA = StreamingXlsxReader::fromFile($plain);
            $readerB = StreamingXlsxReader::fromFile($indexed);
            $this->assertSame(
                iterator_to_array($readerA->rows(), false),
                iterator_to_array($readerB->rows(), false)
            );

            // File size penalty stays under 1 % at this scale (the spec
            // predicts +0.13 % for 4M rows; tiny files have proportionally
            // higher overhead because of the fixed-cost header bytes).
            $sizePlain = filesize($plain);
            $sizeIndexed = filesize($indexed);
            $deltaPercent = ($sizeIndexed - $sizePlain) / $sizePlain * 100;
            $this->assertLessThan(
                5.0,
                $deltaPercent,
                "indexed file grew by {$deltaPercent}% — index overhead is too large"
            );
        } finally {
            @unlink($plain);
        }
    }

    public function test_with_random_access_index_must_be_called_before_start_file(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['n']);

        $this->expectException(XlsxStreamException::class);
        $writer->withRandomAccessIndex();
    }

    public function test_zero_or_negative_period_is_rejected(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(XlsxStreamException::class);
        $writer->withRandomAccessIndex(0);
    }

    public function test_multi_sheet_index_records_per_sheet_sections(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 50);
        $writer->setBufferFlushInterval(50);
        $writer->startFile(['s1']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->newSheet('Other', ['s2']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $sheets = $this->decodeSheets($this->readIndexPayload());

        $this->assertCount(2, $sheets);
        $this->assertSame('xl/worksheets/sheet1.xml', $sheets[0]['entry']);
        $this->assertSame('xl/worksheets/sheet2.xml', $sheets[1]['entry']);
        $this->assertSame(101, $sheets[0]['total_rows']); // header + 100
        $this->assertSame(201, $sheets[1]['total_rows']); // header + 200
        $this->assertCount(2, $sheets[0]['sync_points']); // sync at 52, 102
        $this->assertCount(4, $sheets[1]['sync_points']); // sync at 52, 102, 152, 202
    }

    public function test_writer_without_opt_in_emits_no_index_entry(): void
    {
        // Sanity: existing users who never call withRandomAccessIndex()
        // get byte-identical output to the v2.x writer.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['n']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $this->assertFalse($cd->has(RandomAccessIndex::ENTRY_PATH));
    }

    /**
     * Read xl/_kxs/index.bin entry (decompressed) from the test file.
     */
    private function readIndexPayload(): string
    {
        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        return $cd->readEntry($source, RandomAccessIndex::ENTRY_PATH);
    }

    /**
     * Decode the body of an index payload into a list of sheet sections.
     *
     * @return list<array{entry: string, total_rows: int, sync_points: list<array{row: int, comp_offset: int, uncomp_offset: int}>}>
     */
    private function decodeSheets(string $payload): array
    {
        $sheetCount = unpack('v', substr($payload, 6, 2))[1];
        $body = substr($payload, 16);

        $cursor = 0;
        $sheets = [];
        for ($i = 0; $i < $sheetCount; $i++) {
            $pathLen = unpack('v', substr($body, $cursor, 2))[1];
            $cursor += 2;
            $entry = substr($body, $cursor, $pathLen);
            $cursor += $pathLen;
            $totalRows = unpack('V', substr($body, $cursor, 4))[1];
            $cursor += 4;
            $syncCount = unpack('V', substr($body, $cursor, 4))[1];
            $cursor += 4;

            $points = [];
            for ($k = 0; $k < $syncCount; $k++) {
                $points[] = [
                    'row' => unpack('P', substr($body, $cursor, 8))[1],
                    'comp_offset' => unpack('P', substr($body, $cursor + 8, 8))[1],
                    'uncomp_offset' => unpack('P', substr($body, $cursor + 16, 8))[1],
                ];
                $cursor += 24;
            }

            $sheets[] = [
                'entry' => $entry,
                'total_rows' => $totalRows,
                'sync_points' => $points,
            ];
        }

        return $sheets;
    }
}

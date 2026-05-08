<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Tests\TestCase;
use ZipArchive;

/**
 * Tests for reading XLSX archives that use the shared-strings table.
 *
 * SinkableXlsxWriter only emits inline strings, so to exercise the
 * t="s" code path each test builds a minimal OPC archive matching the
 * shape PhpSpreadsheet, openpyxl, Apache POI, etc. produce. The shared
 * strings table is small in every test (a few KB at most) — within the
 * package's bounded-RAM contract, no memory_limit jugglery.
 */
class ExternalXlsxTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-external-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_reader_resolves_shared_string_references(): void
    {
        $this->buildExternalXlsx(
            sharedStrings: ['Istanbul', 'Ankara', 'İzmir'],
            sheetRows: [
                [['t' => 's', 'v' => '0']], // header → "Istanbul"
                [['t' => 's', 'v' => '1'], ['t' => 'n', 'v' => '42']],
                [['t' => 's', 'v' => '2'], ['t' => 'n', 'v' => '7']],
            ]
        );

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertCount(3, $rows);
        $this->assertSame(['Istanbul'], $rows[0]);
        $this->assertSame(['Ankara', '42'], $rows[1]);
        $this->assertSame(['İzmir', '7'], $rows[2]);
    }

    public function test_unicode_and_entities_in_shared_strings_round_trip(): void
    {
        $this->buildExternalXlsx(
            sharedStrings: ['foo & bar', '<tag>', 'İstanbul 🌊'],
            sheetRows: [
                [['t' => 's', 'v' => '0']],
                [['t' => 's', 'v' => '1']],
                [['t' => 's', 'v' => '2']],
            ]
        );

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame(['foo & bar'], $rows[0]);
        $this->assertSame(['<tag>'], $rows[1]);
        $this->assertSame(['İstanbul 🌊'], $rows[2]);
    }

    public function test_mixed_inline_and_shared_strings_in_same_sheet(): void
    {
        // Some external writers mix inline strings (long text) with sst
        // references (short repeated values) in the same sheet. Reader
        // must handle both within a single row.
        $this->buildExternalXlsx(
            sharedStrings: ['Active', 'Pending'],
            sheetRows: [
                [
                    ['t' => 'inlineStr', 'is' => 'first user'],
                    ['t' => 's', 'v' => '0'],
                    ['t' => 'n', 'v' => '100'],
                ],
                [
                    ['t' => 'inlineStr', 'is' => 'second user'],
                    ['t' => 's', 'v' => '1'],
                    ['t' => 'n', 'v' => '200'],
                ],
            ]
        );

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame(['first user', 'Active', '100'], $rows[0]);
        $this->assertSame(['second user', 'Pending', '200'], $rows[1]);
    }

    public function test_50_row_phpspreadsheet_style_workbook_round_trips(): void
    {
        // Mirrors the layout PhpSpreadsheet / openpyxl produce by default:
        // multi-column data, a small dedup'd shared-strings table holding
        // header labels + the few repeated category strings, numeric cells
        // for IDs and amounts. 50 data rows is a realistic minimum size
        // for a "category report" export.
        $sst = ['ID', 'Category', 'Status', 'Amount', 'Active', 'Pending', 'Closed'];

        $rows = [];
        // Header row — every cell points into the sst (PhpSpreadsheet's typical pattern)
        $rows[] = [
            ['t' => 's', 'v' => '0'], // ID
            ['t' => 's', 'v' => '1'], // Category
            ['t' => 's', 'v' => '2'], // Status
            ['t' => 's', 'v' => '3'], // Amount
        ];
        // 50 data rows, status string deduped via sst, amount as numeric
        $statuses = ['4', '5', '6']; // sst indexes for Active, Pending, Closed
        for ($i = 1; $i <= 50; $i++) {
            $rows[] = [
                ['t' => 'n', 'v' => (string) $i],
                ['t' => 'inlineStr', 'is' => "Category-{$i}"],
                ['t' => 's', 'v' => $statuses[$i % 3]],
                ['t' => 'n', 'v' => (string) ($i * 10.5)],
            ];
        }

        $this->buildExternalXlsx(sharedStrings: $sst, sheetRows: $rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $allRows = iterator_to_array($reader->rows(), false);

        $this->assertCount(51, $allRows);
        $this->assertSame(['ID', 'Category', 'Status', 'Amount'], $allRows[0]);

        // First data row
        $this->assertSame(['1', 'Category-1', 'Pending', '10.5'], $allRows[1]);
        // Last data row — i=50, 50 % 3 = 2 → 'Closed'
        $this->assertSame(['50', 'Category-50', 'Closed', '525'], $allRows[50]);

        // Sample deduped statuses cycle Pending(1) → Closed(2) → Active(0)
        $this->assertSame('Active', $allRows[3][2]);  // i=3, 3 % 3 = 0
        $this->assertSame('Pending', $allRows[4][2]); // i=4, 4 % 3 = 1

        $this->assertSame(50, $reader->rowCount() - 1);
    }

    public function test_sst_with_extreme_deflate_ratio_is_rejected_via_uncompressed_guard(): void
    {
        // Builds a synthetic XLSX whose xl/sharedStrings.xml fits well
        // under the 20 MB compressed threshold but would inflate to
        // ~110 MB — pathological deflate ratio that adversarial inputs
        // (and some accidental exports) can produce. The new
        // uncompressed-size guard must reject it before any inflation
        // happens, preserving the bounded-RAM contract.
        //
        // Fixture creation needs ~250 MB transiently (raw sst string +
        // ZipArchive deflate buffer). The bumped limit is left in place
        // for the rest of the test process — restoring it would race
        // with the still-allocated reader/zip data and itself fail.
        // The follow-up ExternalXlsxTest cases don't need extra memory.
        ini_set('memory_limit', '512M');

        $sstXml = $this->buildHighRatioSstXml(targetUncompressedBytes: 110 * 1024 * 1024);

        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->testFile, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true);
        $zip->addFromString('[Content_Types].xml', $this->contentTypes());
        $zip->addFromString('_rels/.rels', $this->packageRels());
        $zip->addFromString('xl/workbook.xml', $this->workbookXml());
        $zip->addFromString('xl/_rels/workbook.xml.rels', $this->workbookRels());
        $zip->addFromString('xl/sharedStrings.xml', $sstXml);
        $zip->addFromString('xl/worksheets/sheet1.xml', $this->buildSheetXml([[['t' => 's', 'v' => '0']]]));
        $zip->close();

        unset($sstXml, $zip);
        gc_collect_cycles();

        // SST load is lazy — triggered the first time row data needs a
        // shared-string lookup. fromFile() alone won't reach the guard.
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxReadException::class);
        $this->expectExceptionMessageMatches('/inflates to .+ MB/');

        iterator_to_array($reader->rows());
    }

    public function test_reader_handles_stored_worksheet_method_zero(): void
    {
        // Some editors choose STORED (no compression) for very small
        // worksheets to skip deflate setup cost. The reader must
        // tokenize the raw XML directly instead of feeding it through
        // inflate_add — which would otherwise produce garbage or an
        // inflate error. This test pins parity with ZipDirectory's
        // metadata-read path, which already supports both methods.
        $sheetXml = '<?xml version="1.0" encoding="UTF-8"?>'.
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<sheetData>'.
            '<row r="1"><c r="A1" t="inlineStr"><is><t>id</t></is></c><c r="B1" t="inlineStr"><is><t>name</t></is></c></row>'.
            '<row r="2"><c r="A2" t="n"><v>1</v></c><c r="B2" t="inlineStr"><is><t>Alice</t></is></c></row>'.
            '<row r="3"><c r="A3" t="n"><v>2</v></c><c r="B3" t="inlineStr"><is><t>Bob</t></is></c></row>'.
            '</sheetData></worksheet>';

        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->testFile, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true);
        $zip->addFromString('[Content_Types].xml', $this->contentTypesNoSst());
        $zip->addFromString('_rels/.rels', $this->packageRels());
        $zip->addFromString('xl/workbook.xml', $this->workbookXml());
        $zip->addFromString('xl/_rels/workbook.xml.rels', $this->workbookRelsNoSst());
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheetXml);
        $zip->setCompressionName('xl/worksheets/sheet1.xml', \ZipArchive::CM_STORE);
        $zip->close();

        // Verify CD actually carries STORED method — fixture invariant
        $verify = new \ZipArchive();
        $verify->open($this->testFile);
        $stat = $verify->statName('xl/worksheets/sheet1.xml');
        $verify->close();
        $this->assertSame(\ZipArchive::CM_STORE, $stat['comp_method'], 'fixture must use STORED method');

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertCount(3, $rows);
        $this->assertSame(['id', 'name'], $rows[0]);
        $this->assertSame(['1', 'Alice'], $rows[1]);
        $this->assertSame(['2', 'Bob'], $rows[2]);
    }

    public function test_reader_rejects_unsupported_compression_method(): void
    {
        // Synthesize a CD with an arbitrary unsupported method (e.g. 12 = BZIP2).
        // ZipArchive only emits 0 and 8; we have to write the bytes by hand.
        $stub = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<sheetData><row r="1"/></sheetData></worksheet>';

        $zip = new \ZipArchive();
        $zip->open($this->testFile, \ZipArchive::CREATE | \ZipArchive::OVERWRITE);
        $zip->addFromString('[Content_Types].xml', $this->contentTypesNoSst());
        $zip->addFromString('_rels/.rels', $this->packageRels());
        $zip->addFromString('xl/workbook.xml', $this->workbookXml());
        $zip->addFromString('xl/_rels/workbook.xml.rels', $this->workbookRelsNoSst());
        $zip->addFromString('xl/worksheets/sheet1.xml', $stub);
        $zip->close();

        // Patch the central directory: change the compression method byte
        // for the sheet entry from 8 (DEFLATE) to 12 (BZIP2). We don't
        // need the ZIP to be inflatable — only the CD-time guard fires.
        $bytes = file_get_contents($this->testFile);
        $bytes = $this->patchCdMethod($bytes, 'xl/worksheets/sheet1.xml', 12);
        file_put_contents($this->testFile, $bytes);

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxReadException::class);
        $this->expectExceptionMessageMatches('/unsupported ZIP compression method 12/');

        iterator_to_array($reader->rows());
    }

    public function test_chunked_works_on_external_xlsx(): void
    {
        $sst = [];
        for ($i = 0; $i < 50; $i++) {
            $sst[] = "row-{$i}";
        }
        $rows = [];
        for ($i = 0; $i < 50; $i++) {
            $rows[] = [['t' => 's', 'v' => (string) $i]];
        }

        $this->buildExternalXlsx(sharedStrings: $sst, sheetRows: $rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $batches = iterator_to_array($reader->chunked(20), false);

        $this->assertCount(3, $batches);
        $this->assertCount(20, $batches[0]);
        $this->assertCount(20, $batches[1]);
        $this->assertCount(10, $batches[2]);
        $this->assertSame(['row-0'], $batches[0][0]);
        $this->assertSame(['row-49'], $batches[2][9]);
    }

    /**
     * @param  list<string>  $sharedStrings
     * @param  list<list<array{t: string, v?: string, is?: string}>>  $sheetRows
     */
    private function buildExternalXlsx(array $sharedStrings, array $sheetRows): void
    {
        $zip = new ZipArchive();
        $opened = $zip->open($this->testFile, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        $this->assertTrue($opened === true);

        $zip->addFromString('[Content_Types].xml', $this->contentTypes());
        $zip->addFromString('_rels/.rels', $this->packageRels());
        $zip->addFromString('xl/workbook.xml', $this->workbookXml());
        $zip->addFromString('xl/_rels/workbook.xml.rels', $this->workbookRels());
        $zip->addFromString('xl/sharedStrings.xml', $this->buildSstXml($sharedStrings));
        $zip->addFromString('xl/worksheets/sheet1.xml', $this->buildSheetXml($sheetRows));

        $zip->close();
    }

    /**
     * Stream-style construct a shared-strings XML whose uncompressed
     * size approximates $targetUncompressedBytes, using a single
     * repeated <si> entry. The output deflate-compresses to a tiny
     * fraction of its raw size, exercising the threshold mismatch
     * between the compressed and uncompressed sst guards.
     */
    private function buildHighRatioSstXml(int $targetUncompressedBytes): string
    {
        $header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        $entry = '<si><t>repeated content for deflate ratio testing</t></si>';
        $footer = '</sst>';

        $copies = (int) ceil(($targetUncompressedBytes - strlen($header) - strlen($footer)) / strlen($entry));
        if ($copies < 1) {
            $copies = 1;
        }

        return $header.str_repeat($entry, $copies).$footer;
    }

    /**
     * @param  list<string>  $strings
     */
    private function buildSstXml(array $strings): string
    {
        $count = count($strings);
        $body = '';
        foreach ($strings as $s) {
            $encoded = htmlspecialchars($s, ENT_QUOTES | ENT_XML1, 'UTF-8');
            $body .= '<si><t xml:space="preserve">'.$encoded.'</t></si>';
        }

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'.
            ' count="'.$count.'" uniqueCount="'.$count.'">'.
            $body.'</sst>';
    }

    /**
     * @param  list<list<array{t: string, v?: string, is?: string}>>  $rows
     */
    private function buildSheetXml(array $rows): string
    {
        $body = '';
        foreach ($rows as $rowIdx => $cells) {
            $rowNum = $rowIdx + 1;
            $cellXml = '';
            foreach ($cells as $colIdx => $cell) {
                $ref = $this->columnLetters($colIdx).$rowNum;
                if ($cell['t'] === 'inlineStr') {
                    $text = htmlspecialchars($cell['is'] ?? '', ENT_QUOTES | ENT_XML1, 'UTF-8');
                    $cellXml .= '<c r="'.$ref.'" t="inlineStr"><is><t>'.$text.'</t></is></c>';
                } else {
                    $cellXml .= '<c r="'.$ref.'" t="'.$cell['t'].'"><v>'.($cell['v'] ?? '').'</v></c>';
                }
            }
            $body .= '<row r="'.$rowNum.'">'.$cellXml.'</row>';
        }

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<sheetData>'.$body.'</sheetData>'.
            '</worksheet>';
    }

    private function columnLetters(int $col0): string
    {
        $col = $col0 + 1;
        $letters = '';
        while ($col > 0) {
            $col--;
            $letters = chr(65 + ($col % 26)).$letters;
            $col = intdiv($col, 26);
        }

        return $letters;
    }

    /**
     * Patch the compression-method byte (offset 10-11 of every CD
     * entry) for the entry with the given name. Used to fabricate
     * archives carrying a method ZipArchive itself never emits — lets
     * us exercise the reader's "unsupported method" rejection path
     * without writing a full ZIP encoder.
     */
    private function patchCdMethod(string $bytes, string $name, int $newMethod): string
    {
        // Locate the CD entry signature for the named file.
        $signature = "PK\x01\x02";
        $cursor = 0;
        while (($pos = strpos($bytes, $signature, $cursor)) !== false) {
            $fnameLen = unpack('v', substr($bytes, $pos + 28, 2))[1];
            $extraLen = unpack('v', substr($bytes, $pos + 30, 2))[1];
            $commentLen = unpack('v', substr($bytes, $pos + 32, 2))[1];
            $entryName = substr($bytes, $pos + 46, $fnameLen);
            if ($entryName === $name) {
                // Method field is at offset +10..+11 of the CD entry.
                $patched = substr($bytes, 0, $pos + 10).pack('v', $newMethod).substr($bytes, $pos + 12);

                return $patched;
            }
            $cursor = $pos + 46 + $fnameLen + $extraLen + $commentLen;
        }

        throw new \RuntimeException("CD entry not found in patchCdMethod: {$name}");
    }

    private function contentTypesNoSst(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'.
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'.
            '<Default Extension="xml" ContentType="application/xml"/>'.
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'.
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.
            '</Types>';
    }

    private function workbookRelsNoSst(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'.
            '</Relationships>';
    }

    private function contentTypes(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'.
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'.
            '<Default Extension="xml" ContentType="application/xml"/>'.
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'.
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'.
            '</Types>';
    }

    private function packageRels(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'.
            '</Relationships>';
    }

    private function workbookXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'.
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.
            '<sheets>'.
            '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'.
            '</sheets>'.
            '</workbook>';
    }

    private function workbookRels(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'.
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'.
            '</Relationships>';
    }
}

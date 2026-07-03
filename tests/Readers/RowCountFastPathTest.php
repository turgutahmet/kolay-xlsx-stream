<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingSheetReader;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * Pins the '</row>' boundary-count fast path behind rowCount() /
 * StreamingSheetReader::countRows().
 *
 * The contract under test: countRows() returns exactly what
 * iterator_count(rows()) used to — for every input shape the tokenizer
 * can meet, including self-closing rows (which yield no row of their
 * own in rows()), '</row>' straddling a read-chunk boundary, and empty
 * sheets. Each fixture asserts parity against the tokenizer AND the
 * expected absolute value, so a regression in either path surfaces.
 */
class RowCountFastPathTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-rowcount-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_row_count_matches_tokenizer_on_writer_output(): void
    {
        // DEFLATE path with enough rows that '</row>' inevitably
        // straddles inflate-chunk boundaries many times over.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'name']);
        for ($i = 1; $i <= 5000; $i++) {
            $writer->writeRow([$i, "row-{$i}"]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertSame(5001, $reader->rowCount());
        $this->assertSame(iterator_count($reader->rows()), $reader->rowCount());
    }

    public function test_row_count_with_closing_tag_straddling_chunk_boundary(): void
    {
        // STORED entry + a 3-byte read chunk: raw XML reaches countRows
        // in 3-byte slices, so every '</row>' is guaranteed to be split
        // across chunk boundaries. Exercises the 5-byte carry logic
        // deterministically (with DEFLATE the inflated chunk boundaries
        // are not controllable from a test).
        $this->buildStoredSheetZip(
            '<row r="1"><c r="A1" t="n"><v>1</v></c></row>'.
            '<row r="2"><c r="A2" t="n"><v>2</v></c></row>'.
            '<row r="3"><c r="A3" t="n"><v>3</v></c></row>'
        );

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $reader = new StreamingSheetReader($source, $cd, 'xl/worksheets/sheet1.xml', 3);

        $this->assertSame(3, $reader->countRows());
        $this->assertSame(iterator_count($reader->rows()), $reader->countRows());
    }

    public function test_row_count_with_self_closing_rows_matches_tokenizer(): void
    {
        // External writers emit <row r="N"/> for empty rows. rows()
        // yields no row for them (the slice runs from their opening to
        // the NEXT row's '</row>'), and they contribute no '</row>' —
        // so the boundary count and the tokenizer agree on 2 here, not
        // 4. Parity with the tokenizer is the contract rowCount() pins.
        $this->buildStoredSheetZip(
            '<row r="1"><c r="A1" t="n"><v>1</v></c></row>'.
            '<row r="2"/>'.
            '<row r="3"><c r="A3" t="n"><v>3</v></c></row>'.
            '<row r="4"/>' // trailing self-closing row: no yield either
        );

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $reader = new StreamingSheetReader($source, $cd, 'xl/worksheets/sheet1.xml', 7);

        $this->assertSame(2, $reader->countRows());
        $this->assertSame(iterator_count($reader->rows()), $reader->countRows());
    }

    public function test_row_count_of_empty_sheet_is_zero(): void
    {
        $this->buildSingleSheetWorkbook('<sheetData/>');

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertSame(0, $reader->rowCount());
        $this->assertSame(0, iterator_count($reader->rows()));
    }

    public function test_row_count_via_public_reader_on_deflated_entry(): void
    {
        $this->buildSingleSheetWorkbook(
            '<sheetData>'.
            '<row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c></row>'.
            '<row r="2"/>'.
            '<row r="3"><c r="A3" t="n"><v>9</v></c></row>'.
            '</sheetData>',
            stored: false
        );

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertSame(iterator_count($reader->rows()), $reader->rowCount());
        $this->assertSame(2, $reader->rowCount());
    }

    /**
     * Minimal ZIP with only a STORED worksheet entry — enough for
     * StreamingSheetReader, which never consults workbook.xml. STORED
     * keeps raw XML bytes identical to what fread delivers, so the
     * chunk size fully controls where countRows sees its boundaries.
     */
    private function buildStoredSheetZip(string $rowsXml): void
    {
        $sheetXml = '<?xml version="1.0" encoding="UTF-8"?>'.
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<sheetData>'.$rowsXml.'</sheetData></worksheet>';

        $zip = new ZipArchive();
        $this->assertTrue($zip->open($this->testFile, ZipArchive::CREATE | ZipArchive::OVERWRITE) === true);
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheetXml);
        $zip->setCompressionName('xl/worksheets/sheet1.xml', ZipArchive::CM_STORE);
        $zip->close();
    }

    /**
     * Full minimal workbook so StreamingXlsxReader's constructor can
     * resolve sheets — mirrors the synthetic fixtures used by
     * CastColumnTest / ExternalXlsxTest.
     */
    private function buildSingleSheetWorkbook(string $sheetDataXml, bool $stored = true): void
    {
        $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            $sheetDataXml.'</worksheet>';

        $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'.
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'.
            '<Default Extension="xml" ContentType="application/xml"/>'.
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'.
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.
            '</Types>';

        $packageRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'.
            '</Relationships>';

        $workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'.
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'.
            '</workbook>';

        $workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'.
            '</Relationships>';

        $zip = new ZipArchive();
        $this->assertTrue($zip->open($this->testFile, ZipArchive::CREATE | ZipArchive::OVERWRITE) === true);
        $zip->addFromString('[Content_Types].xml', $contentTypes);
        $zip->addFromString('_rels/.rels', $packageRels);
        $zip->addFromString('xl/workbook.xml', $workbookXml);
        $zip->addFromString('xl/_rels/workbook.xml.rels', $workbookRels);
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheetXml);
        if ($stored) {
            $zip->setCompressionName('xl/worksheets/sheet1.xml', ZipArchive::CM_STORE);
        }
        $zip->close();
    }
}

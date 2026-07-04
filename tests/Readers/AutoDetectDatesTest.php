<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\BaseXlsxWriter;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * autoDetectDates() — opt-in numFmt-based date detection.
 *
 * External writers store dates as t="n" serials whose only marker is
 * the cell style's number format; these tests pin the full chain
 * (styles.xml bitmap → tokenizer conversion → facade semantics) against
 * files from our own writer AND hand-built external-shaped archives.
 *
 * Serial 45292 = 2024-01-01 00:00 UTC in the 1900 epoch — used
 * throughout as the canonical probe value.
 */
class AutoDetectDatesTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-autodate-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_date_formatted_column_yields_datetimeimmutable(): void
    {
        $this->writeThreeColumnFile(); // col 2 (1-based) styled 'date'

        $reader = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame(['id', 'when', 'amount'], $rows[0], 'header strings untouched');
        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][1]);
        $this->assertSame('2024-01-01 00:00:00', $rows[1][1]->format('Y-m-d H:i:s'));
        $this->assertSame('UTC', $rows[1][1]->getTimezone()->getName());
        // Unstyled numeric neighbours stay raw strings.
        $this->assertSame('1', $rows[1][0]);
        $this->assertSame('99.5', $rows[1][2]);
    }

    public function test_without_opt_in_serials_stay_raw(): void
    {
        // Regression pin: the raw-serial output shape is a contract;
        // detection must never turn itself on.
        $this->writeThreeColumnFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame('45292', $rows[1][1]);
    }

    public function test_builtin_numfmt_date_is_detected(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['when']);
        $writer->setColumnFormat(1, BaseXlsxWriter::BUILTIN_NUMFMT_DATE); // id 14, locale-aware
        $writer->writeRow([45292]);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][0]);
        $this->assertSame('2024-01-01', $rows[1][0]->format('Y-m-d'));
    }

    public function test_with_time_false_truncates_to_midnight(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['when']);
        $writer->setColumnFormat(1, 'datetime');
        $writer->writeRow([45292.5]); // 2024-01-01 12:00
        $writer->finishFile();

        $withTime = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();
        $rows = iterator_to_array($withTime->rows(), false);
        $this->assertSame('2024-01-01 12:00:00', $rows[1][0]->format('Y-m-d H:i:s'));

        $dateOnly = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates(withTime: false);
        $rows = iterator_to_array($dateOnly->rows(), false);
        $this->assertSame('2024-01-01 00:00:00', $rows[1][0]->format('Y-m-d H:i:s'));
    }

    public function test_explicit_castColumn_takes_precedence(): void
    {
        $this->writeThreeColumnFile();

        // Cast the date-styled column (0-based index 1) to int — the
        // explicit cast must see the RAW serial, not a DateTimeImmutable
        // that is_numeric() would reject into null.
        $reader = StreamingXlsxReader::fromFile($this->testFile)
            ->autoDetectDates()
            ->castColumn(1, 'int');
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame(45292, $rows[1][1]);
    }

    public function test_writer_datetime_round_trips_without_any_cast(): void
    {
        // Our writer stamps DateTime cells with its default datetime
        // style (numFmtId 164, 'yyyy-mm-dd hh:mm:ss') — a date format,
        // so the opt-in reader recovers DateTimeImmutable symmetrically
        // with zero per-column configuration on either side.
        $original = new \DateTimeImmutable('2026-05-06 12:00:00', new \DateTimeZone('UTC'));

        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['when']);
        $writer->writeRow([$original]);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][0]);
        $this->assertSame($original->format('Y-m-d H:i:s'), $rows[1][0]->format('Y-m-d H:i:s'));
    }

    public function test_cast_timezone_applies_to_detected_dates(): void
    {
        $this->writeThreeColumnFile(serial: 45292.5);

        $reader = StreamingXlsxReader::fromFile($this->testFile)
            ->autoDetectDates()
            ->castTimezone('Europe/Istanbul');
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame('Europe/Istanbul', $rows[1][1]->getTimezone()->getName());
        $this->assertSame('2024-01-01 15:00:00', $rows[1][1]->format('Y-m-d H:i:s')); // UTC+3
    }

    public function test_out_of_range_serial_passes_through_unconverted(): void
    {
        // Detection is an inference — a date-styled cell holding a value
        // castDate() rejects keeps its raw value rather than nulling out.
        $this->writeThreeColumnFile(serial: -5);

        $reader = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame('-5', $rows[1][1]);
    }

    public function test_external_file_without_t_attribute_detects_dates(): void
    {
        // Excel itself omits t= on numeric cells entirely — only s=
        // marks the date. Hand-built external-shaped archive (a
        // PhpSpreadsheet-style stylesheet with cellStyleXfs) because
        // PhpSpreadsheet isn't a test dependency; the fixture mirrors
        // its output byte shapes.
        $this->buildExternalDateXlsx(date1904: false);

        $reader = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][0]);
        $this->assertSame('2024-01-01', $rows[1][0]->format('Y-m-d'));
        // Same style on a SHARED STRING cell must not convert (t="s").
        $this->assertSame('not a date', $rows[1][1]);
        // Unstyled numeric stays raw.
        $this->assertSame('45292', $rows[1][2]);
    }

    public function test_1904_epoch_file_detects_shifted_dates(): void
    {
        // Mac-origin file: workbookPr/@date1904 + a date style. Serial 0
        // anchors at 1904-01-01 — detection must ride the auto-detected
        // epoch exactly like castColumn('date') does.
        $this->buildExternalDateXlsx(date1904: true, serial: '0');

        $reader = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][0]);
        $this->assertSame('1904-01-01', $rows[1][0]->format('Y-m-d'));
    }

    public function test_rowsWhere_matches_raw_serials_on_queried_column(): void
    {
        // Query semantics contract: rowsWhere() compares raw numeric
        // values (they must line up with writer block stats), so the
        // queried column is exempt from detection — in the predicate
        // AND in the yielded row. Other columns convert normally.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['when', 'due']);
        $writer->setColumnFormat(1, 'date');
        $writer->setColumnFormat(2, 'date');
        $writer->writeRow([45292, 45300]);
        $writer->writeRow([45293, 45301]);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();
        $matches = iterator_to_array($reader->rowsWhere(1, '=', 45293), true);

        $this->assertCount(1, $matches);
        $row = $matches[3]; // 1-based sheet row (header is row 1)
        $this->assertSame('45293', $row[0], 'queried column stays raw');
        $this->assertInstanceOf(\DateTimeImmutable::class, $row[1], 'other columns still convert');
    }

    public function test_header_helper_and_skip_interplay(): void
    {
        // The recommended pattern — header() then rows(skip: 1) — works
        // with detection on: header text unaffected, data rows typed.
        $this->writeThreeColumnFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)->autoDetectDates();

        $this->assertSame(['id', 'when', 'amount'], $reader->header());

        $dataRows = iterator_to_array($reader->rows(skip: 1), false);
        $this->assertInstanceOf(\DateTimeImmutable::class, $dataRows[0][1]);
    }

    /**
     * Three columns, middle one styled with the 'date' preset
     * (yyyy-mm-dd custom numFmt): [id, serial, amount].
     */
    private function writeThreeColumnFile(int|float $serial = 45292): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'when', 'amount']);
        $writer->setColumnFormat(2, 'date');
        $writer->writeRow([1, $serial, 99.5]);
        $writer->writeRow([2, $serial + 1, 147.25]);
        $writer->finishFile();
    }

    /**
     * Minimal external-shaped archive: numeric cells WITHOUT t=, s="1"
     * pointing at builtin date format 14, a PhpSpreadsheet-style
     * stylesheet (cellStyleXfs present), a small sst, and an optional
     * 1904-epoch workbook declaration.
     */
    private function buildExternalDateXlsx(bool $date1904, string $serial = '45292'): void
    {
        $stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'.
            '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'.
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'.
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'.
            '<cellXfs count="2">'.
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'.
            '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'.
            '</cellXfs>'.
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'.
            '</styleSheet>';

        $workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'.
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.
            ($date1904 ? '<workbookPr date1904="1"/>' : '').
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'.
            '</workbook>';

        $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<sheetData>'.
            '<row r="1">'.
            '<c r="A1" t="s"><v>0</v></c>'.
            '<c r="B1" t="s"><v>1</v></c>'.
            '<c r="C1" t="s"><v>2</v></c>'.
            '</row>'.
            '<row r="2">'.
            '<c r="A2" s="1"><v>'.$serial.'</v></c>'.       // date-styled, no t= — Excel's shape
            '<c r="B2" s="1" t="s"><v>3</v></c>'.           // date style on a STRING cell
            '<c r="C2"><v>45292</v></c>'.                   // unstyled numeric
            '</row>'.
            '</sheetData></worksheet>';

        $sstXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">'.
            '<si><t>when</t></si><si><t>label</t></si><si><t>raw</t></si><si><t>not a date</t></si>'.
            '</sst>';

        $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'.
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'.
            '<Default Extension="xml" ContentType="application/xml"/>'.
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'.
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'.
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'.
            '</Types>';

        $packageRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'.
            '</Relationships>';

        $workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'.
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'.
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'.
            '</Relationships>';

        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->testFile, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true);
        $zip->addFromString('[Content_Types].xml', $contentTypes);
        $zip->addFromString('_rels/.rels', $packageRels);
        $zip->addFromString('xl/workbook.xml', $workbookXml);
        $zip->addFromString('xl/_rels/workbook.xml.rels', $workbookRels);
        $zip->addFromString('xl/styles.xml', $stylesXml);
        $zip->addFromString('xl/sharedStrings.xml', $sstXml);
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheetXml);
        $zip->close();
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * castColumn() / castColumns() / castTimezone() / use1904Epoch() —
 * opt-in cell value casting at the reader layer.
 *
 * Edge cases pinned by these tests:
 *   - 1900 leap-year quirk (serial 60 = real-world 1900-02-28, not 02-29)
 *   - UTC default regardless of date_default_timezone_get()
 *   - Explicit timezone override via castTimezone()
 *   - Invalid timezone fails fast at config time
 */
class CastColumnTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-cast-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_castDate_handles_1900_leap_year_quirk(): void
    {
        $this->writeNumericRows([60]);
        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'date');

        $rows = iterator_to_array($reader->rows(), false);
        // Excel UI shows 1900-02-29 for serial 60, but the real date is 1900-02-28
        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][0]);
        $this->assertSame('1900-02-28', $rows[1][0]->format('Y-m-d'));
    }

    public function test_castDate_uses_UTC_by_default(): void
    {
        $original = date_default_timezone_get();
        try {
            date_default_timezone_set('Europe/Istanbul'); // server config farklı

            // Excel serial 46148.5 = 2026-05-06 12:00 UTC (canonical 1900 epoch)
            $this->writeNumericRows([46148.5]);
            $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'datetime');

            $rows = iterator_to_array($reader->rows(), false);
            $this->assertSame('UTC', $rows[1][0]->getTimezone()->getName());
            $this->assertSame('2026-05-06 12:00:00', $rows[1][0]->format('Y-m-d H:i:s'));
        } finally {
            date_default_timezone_set($original);
        }
    }

    public function test_castDate_explicit_timezone_override(): void
    {
        $this->writeNumericRows([46148.5]);
        $reader = StreamingXlsxReader::fromFile($this->testFile)
            ->castTimezone('Europe/Istanbul')
            ->castColumn(0, 'datetime');

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame('Europe/Istanbul', $rows[1][0]->getTimezone()->getName());
        // UTC 12:00 = Istanbul 15:00 (UTC+3)
        $this->assertSame('2026-05-06 15:00:00', $rows[1][0]->format('Y-m-d H:i:s'));
    }

    public function test_writer_datetime_round_trips_via_reader_cast(): void
    {
        // Symmetric proof: write a real DateTime through the writer,
        // read back through the reader cast. Round-trip equality is
        // the strongest contract we can pin.
        $original = new \DateTimeImmutable('2026-05-06 12:00:00', new \DateTimeZone('UTC'));

        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['when']);
        $writer->writeRow([$original]);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'datetime');
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][0]);
        $this->assertSame(
            $original->format('Y-m-d H:i:s'),
            $rows[1][0]->format('Y-m-d H:i:s')
        );
    }

    public function test_castTimezone_change_invalidates_cached_timezone(): void
    {
        // castDate caches the DateTimeZone instance to avoid a tz
        // database lookup per date cell. Reconfiguring the timezone
        // after rows have already been cast must swap that cache —
        // a stale instance would silently keep the old offset.
        $this->writeNumericRows([46148.5]); // 2026-05-06 12:00 UTC
        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'datetime');

        // First pass primes the cache with the UTC default.
        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame('UTC', $rows[1][0]->getTimezone()->getName());
        $this->assertSame('2026-05-06 12:00:00', $rows[1][0]->format('Y-m-d H:i:s'));

        // Second pass after the override must use the new zone.
        $reader->castTimezone('Europe/Istanbul');
        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame('Europe/Istanbul', $rows[1][0]->getTimezone()->getName());
        $this->assertSame('2026-05-06 15:00:00', $rows[1][0]->format('Y-m-d H:i:s'));

        // And overriding twice keeps tracking the latest setting.
        $reader->castTimezone('America/New_York');
        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame('America/New_York', $rows[1][0]->getTimezone()->getName());
        $this->assertSame('2026-05-06 08:00:00', $rows[1][0]->format('Y-m-d H:i:s'));
    }

    public function test_castTimezone_rejects_invalid_tz(): void
    {
        $this->writeNumericRows([1]);
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\InvalidArgumentException::class);
        $reader->castTimezone('Mars/Olympus_Mons');
    }

    public function test_int_cast(): void
    {
        $this->writeNumericRows(['42', '13.7', 'not-a-number']);
        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'int');

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame(42, $rows[1][0]);
        $this->assertSame(13, $rows[2][0]);
        // Header row 'h0' is non-numeric and gets cast to null
        $this->assertNull($rows[0][0]);
    }

    public function test_float_cast(): void
    {
        $this->writeNumericRows([1.5, 2, 3.14]);
        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'float');

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame(1.5, $rows[1][0]);
        $this->assertSame(2.0, $rows[2][0]);
        $this->assertSame(3.14, $rows[3][0]);
    }

    public function test_bool_cast(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['flag']);
        $writer->writeRow([true]);
        $writer->writeRow([false]);
        $writer->writeRow(['1']); // string truthy
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'bool');
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertTrue($rows[1][0]);
        $this->assertFalse($rows[2][0]);
        $this->assertTrue($rows[3][0]);
    }

    public function test_callable_cast(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['amount']);
        $writer->writeRow(['100']);
        $writer->writeRow(['250']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)
            ->castColumn(0, fn ($v) => '$'.$v);

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame('$100', $rows[1][0]);
        $this->assertSame('$250', $rows[2][0]);
    }

    public function test_castColumns_bulk(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'price', 'serial']);
        $writer->writeRow(['1', '9.95', '46148']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumns([
            0 => 'int',
            1 => 'float',
            2 => 'date',
        ]);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame(1, $rows[1][0]);
        $this->assertSame(9.95, $rows[1][1]);
        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][2]);
        $this->assertSame('2026-05-06', $rows[1][2]->format('Y-m-d'));
    }

    public function test_unknown_cast_type_throws(): void
    {
        $this->writeNumericRows([1]);
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\InvalidArgumentException::class);
        $reader->castColumn(0, 'octopus');
    }

    public function test_workbook_pr_date1904_auto_detected_from_xml(): void
    {
        // Mac-origin XLSX declares <workbookPr date1904="1"/>. Reader
        // must honour that flag on construction so castDate uses the
        // 1904 epoch — without auto-detect, dates silently shift four
        // years and a day. Synthesised file mirrors what Excel-for-Mac
        // emits on a workbook saved with the 1904 date system.
        $this->buildWorkbookWithDate1904Pr(true);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $reader->castColumn(0, 'date');

        $rows = iterator_to_array($reader->rows(), false);
        // Serial 0 is the epoch anchor — 1904-01-01 in 1904 mode,
        // 1899-12-30 in 1900 mode. Auto-detection forces 1904.
        $this->assertSame('1904-01-01', $rows[1][0]->format('Y-m-d'));
    }

    public function test_workbook_pr_date1904_false_keeps_1900_epoch(): void
    {
        $this->buildWorkbookWithDate1904Pr(false);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $reader->castColumn(0, 'date');

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame('1899-12-30', $rows[1][0]->format('Y-m-d'));
    }

    public function test_manual_use1904Epoch_overrides_auto_detect(): void
    {
        // Workbook declares 1900 mode but caller manually opts into
        // 1904 — the explicit method call wins over the auto-detected
        // default, preserving caller intent.
        $this->buildWorkbookWithDate1904Pr(false);

        $reader = StreamingXlsxReader::fromFile($this->testFile)
            ->use1904Epoch()
            ->castColumn(0, 'date');

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame('1904-01-01', $rows[1][0]->format('Y-m-d'));
    }

    public function test_use1904Epoch_shifts_serial_origin(): void
    {
        // Serial 0 in 1904 epoch = 1904-01-01
        $this->writeNumericRows([0]);
        $reader = StreamingXlsxReader::fromFile($this->testFile)
            ->use1904Epoch()
            ->castColumn(0, 'date');

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertSame('1904-01-01', $rows[1][0]->format('Y-m-d'));
    }

    public function test_sparse_row_with_cast_returns_unmodified_when_column_absent(): void
    {
        // Header has only 1 column; cast on col 5 should be a no-op
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id']);
        $writer->writeRow([42]);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(5, 'int');
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame('42', $rows[1][0]);
        $this->assertArrayNotHasKey(5, $rows[1]);
    }

    public function test_castDate_returns_null_for_negative_serial(): void
    {
        // Negative serial is meaningless for Excel dates. Common cause:
        // an upstream tool exporting -1 / -100 from a numeric formula
        // and labelling the column "date" anyway.
        $this->writeNumericRows([-1, -100, -2958466]);
        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'date');

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertNull($rows[1][0]);
        $this->assertNull($rows[2][0]);
        $this->assertNull($rows[3][0]);
    }

    public function test_header_returns_raw_strings_while_rows_apply_casts(): void
    {
        // The recommended pattern when using casts: call header() to
        // grab column names cast-free, then iterate rows(skip: 1) for
        // typed data. Pins the contract that header() never runs casts
        // — otherwise a cast like 'int' would null out string headers
        // like "id".
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'amount']);
        $writer->writeRow(['1', '99.5']);
        $writer->writeRow(['2', '147.25']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $reader->castColumns([0 => 'int', 1 => 'float']);

        $header = $reader->header();
        $this->assertSame(['id', 'amount'], $header, 'header must come back as raw strings');

        $dataRows = iterator_to_array($reader->rows(skip: 1), false);
        $this->assertCount(2, $dataRows);
        $this->assertSame([1, 99.5], $dataRows[0]);
        $this->assertSame([2, 147.25], $dataRows[1]);
    }

    public function test_onSheet_resets_column_casts(): void
    {
        // Multi-sheet workbook with different schemas per sheet. Casts
        // configured for sheet 1 must NOT leak into sheet 2 — different
        // columns mean different semantics. Workbook-wide settings
        // (timezone, 1904 epoch) survive the switch.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'amount']);
        $writer->writeRow([1, '99.5']);
        $writer->newSheet('Notes', ['title', 'body']);
        $writer->writeRow(['Alice', 'free-form text']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // Configure casts for sheet 1
        $reader->onSheetIndex(0)->castColumns([0 => 'int', 1 => 'float']);
        $sheet1 = iterator_to_array($reader->rows(skip: 1), false);
        $this->assertSame(1, $sheet1[0][0]);
        $this->assertSame(99.5, $sheet1[0][1]);

        // Switch to Notes — cast on column 0 (now "title", a string)
        // would corrupt the result if it leaked. Switching must reset.
        $reader->onSheet('Notes');
        $sheet2 = iterator_to_array($reader->rows(skip: 1), false);
        $this->assertSame(['Alice', 'free-form text'], $sheet2[0]);
    }

    public function test_castDate_handles_excel_max_date_boundary(): void
    {
        // 2958465 = 9999-12-31, the highest serial Excel renders as a
        // date. One above it is out of range — null guards callers
        // against time travel into year 12345.
        $this->writeNumericRows([2958465, 2958466, 1e10]);
        $reader = StreamingXlsxReader::fromFile($this->testFile)->castColumn(0, 'date');

        $rows = iterator_to_array($reader->rows(), false);
        $this->assertInstanceOf(\DateTimeImmutable::class, $rows[1][0]);
        $this->assertSame('9999-12-31', $rows[1][0]->format('Y-m-d'));
        $this->assertNull($rows[2][0]);
        $this->assertNull($rows[3][0]);
    }

    /**
     * Build a synthetic XLSX whose workbook.xml carries an explicit
     * <workbookPr date1904="1|0"/> attribute. Used to drive the reader's
     * auto-detection path without depending on a binary fixture file
     * shipped in the repo.
     */
    private function buildWorkbookWithDate1904Pr(bool $date1904): void
    {
        $flag = $date1904 ? '1' : '0';
        $workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'.
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.
            '<workbookPr date1904="'.$flag.'"/>'.
            '<sheets>'.
            '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'.
            '</sheets>'.
            '</workbook>';

        $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<sheetData>'.
            '<row r="1"><c r="A1" t="inlineStr"><is><t>serial</t></is></c></row>'.
            '<row r="2"><c r="A2" t="n"><v>0</v></c></row>'.
            '</sheetData></worksheet>';

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

        $workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'.
            '</Relationships>';

        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->testFile, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true);
        $zip->addFromString('[Content_Types].xml', $contentTypes);
        $zip->addFromString('_rels/.rels', $packageRels);
        $zip->addFromString('xl/workbook.xml', $workbookXml);
        $zip->addFromString('xl/_rels/workbook.xml.rels', $workbookRels);
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheetXml);
        $zip->close();
    }

    /**
     * Helper: write a single-column XLSX where each row has one numeric value.
     *
     * @param  array<int, mixed>  $values
     */
    private function writeNumericRows(array $values): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['h0']);
        foreach ($values as $v) {
            $writer->writeRow([$v]);
        }
        $writer->finishFile();
    }
}

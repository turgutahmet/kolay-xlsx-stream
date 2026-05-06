<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Tests\TestCase;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * v2.2.1 polish — colour validation, font name, empty workbook, column-index
 * out-of-range, and a handful of dedup/edge cases.
 */
class V221PolishTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/v221_'.uniqid().'.xlsx';
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }

    // ── Colour hex validation ───────────────────────────────────────

    public function test_invalid_fill_color_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('6-character hex');
        $writer->setHeaderStyle(['fill' => 'not-a-color']);
    }

    public function test_invalid_text_color_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $this->expectException(XlsxStreamException::class);
        $writer->setHeaderStyle(['color' => '#XYZ123']);
    }

    public function test_short_hex_color_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $this->expectException(XlsxStreamException::class);
        $writer->setHeaderStyle(['fill' => '#abc']);
    }

    public function test_valid_hex_color_with_or_without_hash_accepted()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setHeaderStyle(['fill' => '#4F81BD', 'color' => 'FFFFFF']);
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        $this->assertStringContainsString('rgb="FF4F81BD"', $styles);
        $this->assertStringContainsString('rgb="FFFFFFFF"', $styles);
    }

    // ── Font name option ────────────────────────────────────────────

    public function test_custom_font_name_applied()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setHeaderStyle(['bold' => true, 'name' => 'Arial']);
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        $this->assertStringContainsString('<name val="Arial"/>', $styles);
    }

    public function test_font_name_xml_escaped()
    {
        // A pathological font name with quotes shouldn't break the XML
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setHeaderStyle(['bold' => true, 'name' => 'My "Font" & Co.']);
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        $this->assertStringContainsString('My &quot;Font&quot; &amp; Co.', $styles);
        // ZIP must still validate
        $zip = new ZipArchive();
        $this->assertSame(true, $zip->open($this->testFile, ZipArchive::CHECKCONS));
        $zip->close();
    }

    public function test_default_font_name_is_calibri()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        $this->assertStringContainsString('<name val="Calibri"/>', $styles);
    }

    // ── Empty workbook ──────────────────────────────────────────────

    public function test_empty_workbook_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('empty workbook');
        $writer->finishFile();
    }

    public function test_empty_workbook_after_aborted_sheet_throws()
    {
        // newSheet eagerly opens a sheet, so this exercises the legitimate
        // "no rows, no sheets" path: startFile + immediate finishFile.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setHeaderStyle(['bold' => true]);
        $writer->startFile(['Header']);

        $this->expectException(XlsxStreamException::class);
        $writer->finishFile();
    }

    public function test_workbook_with_only_new_sheet_is_valid()
    {
        // newSheet() opens a sheet eagerly so calling it before the first
        // writeRow takes the place of the auto-created "Report" — the
        // workbook ends up with one named sheet, not zero.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->newSheet('Empty', ['B']);
        $stats = $writer->finishFile();

        $this->assertSame(1, $stats['sheets']);
        $this->assertSame('Empty', $stats['sheet_details'][0]['name']);
    }

    // ── setColumnFormat out-of-range ────────────────────────────────

    public function test_set_column_format_out_of_range_throws_at_sheet_start()
    {
        // Validation deferred to startNewSheet — fires on first writeRow.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(99, 'date');
        $writer->startFile(['A', 'B', 'C']);

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('out of range');
        $writer->writeRow([1, 2, 3]);
    }

    public function test_set_column_format_out_of_range_via_new_sheet_throws()
    {
        // Same check fires when newSheet() opens its sheet.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A', 'B', 'C', 'D', 'E']);
        $writer->writeRow([1, 2, 3, 4, 5]);

        // Pre-configure a format that's only valid for the upcoming 12-col sheet
        $writer->setColumnFormat(12, 'date');

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('out of range');
        // newSheet's headers are only 3 cols → col 12 invalid here
        $writer->newSheet('Tiny', ['x', 'y', 'z']);
    }

    public function test_set_column_format_in_range_after_start_works()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A', 'B', 'C']);
        $writer->setColumnFormat(2, 'date');
        $writer->writeRow([1, new \DateTime('2026-01-15'), 'x']);
        $stats = $writer->finishFile();
        $this->assertSame(1, $stats['rows']);
    }

    public function test_set_column_format_pre_configured_for_new_sheet()
    {
        // The whole point of deferred validation: callers can stage the
        // formats before newSheet() rotates to the larger column set.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['Tiny']);
        $writer->writeRow(['only']);

        $writer->clearColumnFormats();
        $writer->setColumnFormat(2, 'integer');
        $writer->setColumnFormat(3, 'date');
        $writer->setColumnFormat(4, 'currency_try');

        $writer->newSheet('Wider', ['Label', 'Count', 'When', 'Amount']);
        $writer->writeRow(['x', 1, new \DateTime('2026-01-15'), 99.50]);
        $stats = $writer->finishFile();
        $this->assertSame(2, $stats['sheets']);
    }

    // ── Style dedup (===) ───────────────────────────────────────────

    public function test_same_header_style_registered_once()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setHeaderStyle(['bold' => true, 'fill' => '#4F81BD']);
        $writer->setHeaderStyle(['bold' => true, 'fill' => '#4F81BD']); // identical
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        // Only one fill of rgb="FF4F81BD"
        $count = substr_count($styles, 'rgb="FF4F81BD"');
        $this->assertSame(1, $count);
    }

    private function extract(string $entry): string
    {
        $zip = new ZipArchive();
        $zip->open($this->testFile);
        $content = $zip->getFromName($entry);
        $zip->close();

        return $content;
    }
}

<?php

namespace Kolay\XlsxStream\Tests;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

class V22StylingTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/v22_'.uniqid().'.xlsx';
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }

    // ── Header styling ──────────────────────────────────────────────

    public function test_header_style_applied_to_header_row()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setHeaderStyle(['bold' => true, 'fill' => '#4F81BD', 'color' => '#FFFFFF']);
        $writer->startFile(['Name', 'Email']);
        $writer->writeRow(['Alice', 'a@x.com']);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $styles = $this->extract('xl/styles.xml');

        // Header row carries a style id (s="N") that data rows don't
        $this->assertMatchesRegularExpression('#<row r="1">.*<c r="A1" s="\d+"#', $sheet);
        $this->assertStringContainsString('<b/>', $styles);
        $this->assertStringContainsString('rgb="FF4F81BD"', $styles);
        $this->assertStringContainsString('rgb="FFFFFFFF"', $styles);
    }

    public function test_header_style_can_be_changed_between_sheets()
    {
        // Each newSheet() can have its own header style — the registry
        // accumulates styles, the writer just records the most recent id.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setHeaderStyle(['bold' => true, 'fill' => '#4F81BD']);
        $writer->startFile(['A']);
        $writer->writeRow([1]);

        $writer->setHeaderStyle(['bold' => true, 'fill' => '#9BBB59']);
        $writer->newSheet('Second', ['B']);
        $writer->writeRow([2]);
        $stats = $writer->finishFile();

        $this->assertEquals(2, $stats['sheets']);

        $styles = $this->extract('xl/styles.xml');
        $this->assertStringContainsString('rgb="FF4F81BD"', $styles);
        $this->assertStringContainsString('rgb="FF9BBB59"', $styles);
    }

    // ── Column formats ──────────────────────────────────────────────

    public function test_set_column_format_preset_date()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(2, 'date');
        $writer->startFile(['ID', 'Born']);
        $writer->writeRow([1, new \DateTime('2026-01-15', new \DateTimeZone('UTC'))]);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $styles = $this->extract('xl/styles.xml');

        // Cell B2 (Born column) carries the date style id
        $this->assertMatchesRegularExpression('#<c r="B2" s="\d+" t="n"><v>46037</v>#', $sheet);
        $this->assertStringContainsString('formatCode="yyyy-mm-dd"', $styles);
    }

    public function test_set_column_format_preset_currency_try()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, 'currency_try');
        $writer->startFile(['Price']);
        $writer->writeRow([99.50]);
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        $this->assertStringContainsString('₺', $styles);
    }

    public function test_set_column_format_custom_format_code()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, '0.000');
        $writer->startFile(['Precise']);
        $writer->writeRow([3.14159]);
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        $this->assertStringContainsString('formatCode="0.000"', $styles);
    }

    public function test_set_column_format_dedupes_same_code()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, 'date');
        $writer->setColumnFormat(2, 'date'); // same preset → same style id
        $writer->startFile(['A', 'B']);
        $writer->writeRow([1, 2]);
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        // 'yyyy-mm-dd' should appear only once in <numFmts>
        $count = substr_count($styles, 'formatCode="yyyy-mm-dd"');
        $this->assertEquals(1, $count, 'Same format code must be registered once');
    }

    public function test_set_column_format_rejects_invalid_column()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $this->expectException(XlsxStreamException::class);
        $writer->setColumnFormat(0, 'date');
    }

    public function test_string_cells_are_unaffected_by_column_format()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, 'currency_try');
        $writer->startFile(['Mixed']);
        $writer->writeRow(['hello']);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        // String cells stay inlineStr without a style id
        $this->assertStringContainsString('<c r="A2" t="inlineStr">', $sheet);
        $this->assertStringNotContainsString('<c r="A2" s=', $sheet);
    }

    // ── Freeze pane ─────────────────────────────────────────────────

    public function test_freeze_first_row_emits_pane()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->freezeFirstRow();
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<sheetViews>', $sheet);
        $this->assertStringContainsString('ySplit="1"', $sheet);
        $this->assertStringContainsString('topLeftCell="A2"', $sheet);
        $this->assertStringContainsString('state="frozen"', $sheet);
    }

    public function test_freeze_rows_and_columns()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->freezeRowsAndColumns(rows: 2, columns: 3);
        $writer->startFile(['A', 'B', 'C', 'D']);
        $writer->writeRow([1, 2, 3, 4]);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('xSplit="3"', $sheet);
        $this->assertStringContainsString('ySplit="2"', $sheet);
        $this->assertStringContainsString('topLeftCell="D3"', $sheet);
    }

    public function test_no_freeze_pane_when_unset()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertStringNotContainsString('<sheetViews>', $sheet);
    }

    // ── Auto filter ─────────────────────────────────────────────────

    public function test_auto_filter_emits_range_after_sheet_data()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->enableAutoFilter();
        $writer->startFile(['ID', 'Name', 'Email']);
        for ($i = 1; $i <= 5; $i++) {
            $writer->writeRow([$i, "User $i", "u$i@x.com"]);
        }
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        // Range covers all populated rows (header + 5 data = 6)
        $this->assertStringContainsString('<autoFilter ref="A1:C6"/>', $sheet);
        // <autoFilter> must appear AFTER </sheetData>
        $this->assertGreaterThan(
            strpos($sheet, '</sheetData>'),
            strpos($sheet, '<autoFilter')
        );
    }

    // ── Column widths ───────────────────────────────────────────────

    public function test_explicit_column_widths()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnWidths([1 => 8, 2 => 30, 3 => 15.5]);
        $writer->startFile(['ID', 'Name', 'Phone']);
        $writer->writeRow([1, 'Alice', '+90555']);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<col min="1" max="1" width="8"', $sheet);
        $this->assertStringContainsString('<col min="2" max="2" width="30"', $sheet);
        $this->assertStringContainsString('<col min="3" max="3" width="15.5"', $sheet);
    }

    public function test_auto_column_width_uses_header_length()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth();
        $writer->startFile(['ID', 'Customer Email Address']);
        $writer->writeRow([1, 'a@x.com']);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        // 'ID' (2 chars) → max(8, 4) = 8
        $this->assertStringContainsString('<col min="1" max="1" width="8"', $sheet);
        // 'Customer Email Address' (22 chars) → max(8, 24) = 24
        $this->assertStringContainsString('<col min="2" max="2" width="24"', $sheet);
    }

    public function test_manual_widths_override_auto()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth();
        $writer->setColumnWidths([1 => 50]); // override col 1
        $writer->startFile(['ID', 'Name']);
        $writer->writeRow([1, 'x']);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<col min="1" max="1" width="50"', $sheet);
        // Col 2 still gets auto width (4 chars 'Name' → max(8, 6) = 8)
        $this->assertStringContainsString('<col min="2" max="2" width="8"', $sheet);
    }

    public function test_no_cols_block_when_widths_unset()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertStringNotContainsString('<cols>', $sheet);
    }

    // ── newSheet manual multi-sheet ────────────────────────────────

    public function test_new_sheet_creates_named_sheet_with_new_headers()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['ID', 'Name']);
        $writer->writeRow([1, 'Alice']);
        $writer->newSheet('Orders', ['OrderID', 'Total']);
        $writer->writeRow([100, 49.90]);
        $stats = $writer->finishFile();

        $this->assertEquals(2, $stats['sheets']);
        $this->assertEquals('Report', $stats['sheet_details'][0]['name']);
        $this->assertEquals('Orders', $stats['sheet_details'][1]['name']);

        // Verify both sheets exist
        $zip = new ZipArchive();
        $zip->open($this->testFile);
        $sheet1 = $zip->getFromName('xl/worksheets/sheet1.xml');
        $sheet2 = $zip->getFromName('xl/worksheets/sheet2.xml');
        $zip->close();

        $this->assertStringContainsString('<t>ID</t>', $sheet1);
        $this->assertStringContainsString('<t>OrderID</t>', $sheet2);
    }

    public function test_new_sheet_reuses_headers_when_not_specified()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['ID', 'Name']);
        $writer->writeRow([1, 'A']);
        $writer->newSheet('Q2');
        $writer->writeRow([2, 'B']);
        $stats = $writer->finishFile();

        $this->assertEquals(2, $stats['sheets']);
        $zip = new ZipArchive();
        $zip->open($this->testFile);
        $sheet2 = $zip->getFromName('xl/worksheets/sheet2.xml');
        $zip->close();

        // Same headers on Q2
        $this->assertStringContainsString('<t>ID</t>', $sheet2);
        $this->assertStringContainsString('<t>Name</t>', $sheet2);
    }

    public function test_new_sheet_before_start_file_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $this->expectException(XlsxStreamException::class);
        $writer->newSheet('Foo');
    }

    public function test_new_sheet_after_finish_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $this->expectException(XlsxStreamException::class);
        $writer->newSheet('Foo');
    }

    public function test_new_sheet_empty_name_rejected()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $this->expectException(XlsxStreamException::class);
        $writer->newSheet('');
    }

    // ── End-to-end integrity ────────────────────────────────────────

    public function test_full_styling_produces_valid_xlsx()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer
            ->setHeaderStyle(['bold' => true, 'fill' => '#4F81BD', 'color' => '#FFFFFF'])
            ->setColumnFormat(2, 'date')
            ->setColumnFormat(3, 'currency_try')
            ->setColumnFormat(4, 'percent')
            ->setColumnWidths([1 => 8, 2 => 14, 3 => 14, 4 => 10])
            ->freezeFirstRow()
            ->enableAutoFilter();

        $writer->startFile(['ID', 'Date', 'Price', 'Discount']);
        for ($i = 1; $i <= 50; $i++) {
            $writer->writeRow([
                $i,
                new \DateTime("2026-0".(($i % 9) + 1)."-01", new \DateTimeZone('UTC')),
                99.99 + $i,
                ($i % 100) / 100,
            ]);
        }

        $writer->newSheet('Summary', ['Metric', 'Value']);
        $writer->writeRow(['Total', 100]);
        $writer->writeRow(['Average', 50.5]);
        $stats = $writer->finishFile();

        $this->assertEquals(2, $stats['sheets']);
        $this->assertEquals(52, $stats['rows']);

        $zip = new ZipArchive();
        $this->assertTrue($zip->open($this->testFile, ZipArchive::CHECKCONS) === true);
        $zip->close();
    }

    // ── helper ──────────────────────────────────────────────────────

    private function extract(string $entry): string
    {
        $zip = new ZipArchive();
        $zip->open($this->testFile);
        $content = $zip->getFromName($entry);
        $zip->close();

        return $content;
    }
}

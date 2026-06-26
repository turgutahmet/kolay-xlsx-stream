<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * v3.0.2 — per-row styling via writeRow($row, $styleId).
 */
class V302RowStyleTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/v302_'.uniqid().'.xlsx';
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }

    private function extract(string $entry): string
    {
        $zip = new ZipArchive();
        $zip->open($this->testFile);
        $content = $zip->getFromName($entry);
        $zip->close();

        return $content;
    }

    public function test_row_style_stamps_every_cell_in_the_row()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $red = $writer->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
        $writer->startFile(['ID', 'Name', 'Status']);
        $writer->writeRow([1, 'Alice', 'OK']);            // unstyled
        $writer->writeRow([2, 'Bob', 'FAILED'], $red);    // styled
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $styles = $this->extract('xl/styles.xml');

        // The styled fill + font color land in styles.xml
        $this->assertStringContainsString('rgb="FFFFC7CE"', $styles);
        $this->assertStringContainsString('rgb="FF9C0006"', $styles);

        // Header is row 1; data rows follow. Row 3 is the styled "FAILED" row.
        preg_match('#<row r="3">(.*?)</row>#', $sheet, $m);
        $this->assertNotEmpty($m, 'row 3 not found');
        $this->assertSame(3, substr_count($m[1], '<c '), 'expected 3 cells in styled row');
        $this->assertSame(3, substr_count($m[1], ' s="'), 'every cell in styled row must carry s=');

        // Row 2 (unstyled data) must NOT carry any style attribute — fast path intact
        preg_match('#<row r="2">(.*?)</row>#', $sheet, $m1);
        $this->assertStringNotContainsString(' s="', $m1[1]);
    }

    public function test_null_passes_through_to_unstyled_fast_path()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->registerRowStyle(['fill' => '#FFC7CE']); // registered but not applied
        $writer->startFile(['A', 'B']);
        $writer->writeRow(['x', 'y']);          // implicit null style
        $writer->writeRow(['z', 'w'], null);    // explicit null style
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        // Header is row 1; data rows are rows 2 and 3. Neither may carry a style id.
        $this->assertStringNotContainsString('<c r="A2" s="', $sheet);
        $this->assertStringNotContainsString('<c r="A3" s="', $sheet);
    }

    public function test_empty_cells_in_styled_row_are_still_filled()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $blue = $writer->registerRowStyle(['fill' => '#1F4E78', 'color' => '#FFFFFF']);
        $writer->startFile(['A', 'B', 'C']);
        $writer->writeRow(['x', null, 'z'], $blue);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        // Header is row 1; the styled row is row 2. The empty B2 cell still
        // carries the style so the highlight is contiguous.
        $this->assertMatchesRegularExpression('#<c r="B2" s="\d+"/>#', $sheet);
    }

    public function test_dedup_one_style_for_many_styled_rows()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $red = $writer->registerRowStyle(['fill' => '#FFC7CE']);
        $writer->startFile(['A']);
        for ($i = 0; $i < 1000; $i++) {
            $writer->writeRow([$i], $red);
        }
        $writer->finishFile();

        $styles = $this->extract('xl/styles.xml');
        // The fill color appears exactly once despite 1000 styled rows.
        $this->assertSame(1, substr_count($styles, 'rgb="FFFFC7CE"'));
    }

    public function test_row_style_merges_with_column_number_format()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(2, 'currency_try');
        $red = $writer->registerRowStyle(['fill' => '#FFC7CE']);
        $writer->startFile(['Name', 'Amount']);
        $writer->writeRow(['Alice', 1234.5], $red);
        $writer->finishFile();

        $sheet = $this->extract('xl/worksheets/sheet1.xml');
        $styles = $this->extract('xl/styles.xml');

        // The currency format code survives.
        $this->assertStringContainsString('₺', $styles);
        // The styled currency cell (B2 — header is row 1) uses a merged xf
        // (fill + numFmt), not the bare currency style nor a fill-only style —
        // i.e. a cellXf with BOTH a custom numFmtId and the fill's fontId/fillId.
        preg_match('#<c r="B2" s="(\d+)" t="n">#', $sheet, $m);
        $this->assertNotEmpty($m, 'B2 styled numeric cell not found');

        // The merged xf must apply both number format and fill.
        $mergedId = (int) $m[1];
        preg_match_all('#<xf [^>]*/>#', $styles, $xfs);
        $this->assertArrayHasKey($mergedId, $xfs[0], 'merged xf index out of range');
        $mergedXf = $xfs[0][$mergedId];
        $this->assertStringContainsString('applyNumberFormat="1"', $mergedXf);
        $this->assertStringContainsString('applyFill="1"', $mergedXf);
    }

    public function test_styled_file_is_valid_zip_and_opens()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $ok = $writer->registerRowStyle(['fill' => '#C6EFCE', 'color' => '#006100']);
        $bad = $writer->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006', 'bold' => true]);
        $writer->startFile(['ID', 'When', 'Amount', 'Status']);
        $writer->writeRow([1, new \DateTimeImmutable('2024-01-01 10:00:00'), 99.9, 'OK'], $ok);
        $writer->writeRow([2, new \DateTimeImmutable('2024-02-01 11:00:00'), 1.5, 'FAILED'], $bad);
        $writer->writeRow([3, new \DateTimeImmutable('2024-03-01 12:00:00'), 7.0, 'OK']);
        $stats = $writer->finishFile();

        $this->assertSame(3, $stats['rows']);

        $zip = new ZipArchive();
        $this->assertTrue($zip->open($this->testFile) === true, 'output is not a valid zip');
        $this->assertNotFalse($zip->getFromName('xl/worksheets/sheet1.xml'));
        $zip->close();
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * setAutoColumnWidth(sample: N, strict: false|true) — opt-in sample
 * pass that scans the first N data rows and derives per-column width
 * from the widest observed value.
 *
 * Pinned behaviours:
 *   - <cols> reflects max(header_len, max_data_len) + 2, clamped [8.43, 255]
 *   - finishFile() before the sample size is reached still emits a valid file
 *   - Multi-sheet workbooks re-sample per sheet
 *   - Kill switch: lenient mode catches internal failures and falls back
 *     to heuristic — strict mode propagates the exception
 */
class AutoColumnWidthSampleTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-autowidth-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_sample_width_reflects_widest_data_value(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(sample: 5);
        $writer->startFile(['id', 'description']);
        $writer->writeRow([1, 'short']);
        $writer->writeRow([2, 'this is a much longer description than the header']);
        $writer->writeRow([3, 'mid']);
        $writer->finishFile();

        $cols = $this->extractColsXml();

        // col 1 (header "id" = 2 chars) → max(8.43, 2+2) = 8.43
        $this->assertMatchesRegularExpression('/<col min="1" max="1" width="8\.43"/', $cols);

        // col 2: longest data value is 49 chars, so width = 49 + 2 = 51
        $this->assertMatchesRegularExpression('/<col min="2" max="2" width="51"/', $cols);
    }

    public function test_sample_size_zero_keeps_heuristic_behaviour(): void
    {
        // setAutoColumnWidth(0) means "heuristic mode" — same as the
        // boolean-true default. Sample mode requires sample > 0.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(sample: 0);
        $writer->startFile(['veryLongHeader', 'b']);
        $writer->writeRow(['x', 'this longer data should NOT influence width in heuristic mode']);
        $writer->finishFile();

        $cols = $this->extractColsXml();

        // col 1: header "veryLongHeader" (14) + 2 = 16
        $this->assertMatchesRegularExpression('/<col min="1" max="1" width="16"/', $cols);
        // col 2: header "b" (1) + 2 = 3, clamp to 8 (heuristic floor)
        $this->assertMatchesRegularExpression('/<col min="2" max="2" width="8"/', $cols);
    }

    public function test_finishFile_before_sample_size_still_emits_valid_file(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(sample: 1000);
        $writer->startFile(['id', 'name']);
        $writer->writeRow([1, 'Alice']);
        $writer->writeRow([2, 'Bob']);
        $writer->finishFile();

        $bytes = file_get_contents($this->testFile);
        $this->assertSame('PK', substr($bytes, 0, 2));

        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->testFile) === true);
        $sheet = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        // Header must be present
        $this->assertStringContainsString('<row r="1">', $sheet);
        $this->assertStringContainsString('<t>id</t>', $sheet);
        // Data rows must be present
        $this->assertStringContainsString('<row r="2">', $sheet);
        $this->assertStringContainsString('<row r="3">', $sheet);
        $this->assertStringContainsString('<t>Alice</t>', $sheet);
    }

    public function test_multi_sheet_resamples_per_sheet(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(sample: 10);

        $writer->startFile(['a']);
        $writer->writeRow(['short']);
        $writer->newSheet('Wide', ['a']);
        $writer->writeRow(['this is a substantially wider sheet 2 row value']);
        $writer->finishFile();

        $cols1 = $this->extractColsXml('xl/worksheets/sheet1.xml');
        $cols2 = $this->extractColsXml('xl/worksheets/sheet2.xml');

        // Sheet 1: max(header "a"=1, data "short"=5) + 2 = 7 → clamped to 8.43
        $this->assertMatchesRegularExpression('/<col min="1" max="1" width="8\.43"/', $cols1);
        // Sheet 2: data 47 chars + 2 = 49 — sample widths must NOT leak from sheet 1
        $this->assertMatchesRegularExpression('/<col min="1" max="1" width="49"/', $cols2);
    }

    public function test_explicit_setColumnWidths_wins_over_sample(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(sample: 5);
        $writer->setColumnWidths([1 => 30.0]); // explicit override
        $writer->startFile(['id']);
        $writer->writeRow([str_repeat('x', 100)]); // would suggest width ~102
        $writer->finishFile();

        $cols = $this->extractColsXml();
        $this->assertMatchesRegularExpression('/<col min="1" max="1" width="30"/', $cols);
    }

    public function test_setAutoColumnWidth_rejects_negative_sample(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxStreamException::class);
        $writer->setAutoColumnWidth(sample: -1);
    }

    public function test_strict_mode_propagates_internal_failure(): void
    {
        $writer = new class(new FileSink($this->testFile)) extends SinkableXlsxWriter
        {
            public bool $shouldFail = false;

            protected function updateAutoWidthMaxLengths(array $row): void
            {
                if ($this->shouldFail) {
                    throw new \RuntimeException('test injected failure');
                }
                parent::updateAutoWidthMaxLengths($row);
            }
        };
        $writer->setAutoColumnWidth(sample: 100, strict: true);
        $writer->startFile(['x']);
        $writer->writeRow(['ok']);
        $writer->shouldFail = true;

        $this->expectException(\RuntimeException::class);
        $this->expectExceptionMessage('test injected failure');
        $writer->writeRow(['this throws']);
    }

    public function test_lenient_mode_falls_back_to_heuristic_on_failure(): void
    {
        $writer = new class(new FileSink($this->testFile)) extends SinkableXlsxWriter
        {
            public bool $shouldFail = false;

            protected function updateAutoWidthMaxLengths(array $row): void
            {
                if ($this->shouldFail) {
                    throw new \RuntimeException('test injected failure');
                }
                parent::updateAutoWidthMaxLengths($row);
            }
        };
        $writer->setAutoColumnWidth(sample: 100, strict: false); // lenient
        $writer->startFile(['x']);
        $writer->writeRow(['short']);
        $writer->shouldFail = true;

        // Suppress error_log output during the lenient bail path.
        $previous = ini_set('error_log', '/dev/null');
        try {
            $writer->writeRow(['this is fine']);
            $writer->writeRow(['and this too']);
            $writer->finishFile();
        } finally {
            ini_set('error_log', $previous === false ? '' : $previous);
        }

        $bytes = file_get_contents($this->testFile);
        $this->assertSame('PK', substr($bytes, 0, 2), 'lenient mode must produce a valid ZIP');

        $zip = new \ZipArchive();
        $zip->open($this->testFile);
        $sheet = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        // All four rows (header + 3 data) must be in the file
        $this->assertStringContainsString('<row r="1">', $sheet);
        $this->assertStringContainsString('<row r="2">', $sheet);
        $this->assertStringContainsString('<row r="3">', $sheet);
        $this->assertStringContainsString('<row r="4">', $sheet);
    }

    public function test_round_trip_through_streaming_reader(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(sample: 3);
        $writer->startFile(['id', 'name']);
        $writer->writeRow([1, 'Alice']);
        $writer->writeRow([2, 'Bob']);
        $writer->writeRow([3, 'Charlie']);
        $writer->writeRow([4, 'Dave']);
        $writer->writeRow([5, 'Eve']);
        $writer->finishFile();

        $reader = \Kolay\XlsxStream\Readers\StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertCount(6, $rows);
        $this->assertSame(['id', 'name'], $rows[0]);
        $this->assertSame(['1', 'Alice'], $rows[1]);
        $this->assertSame(['5', 'Eve'], $rows[5]);
    }

    private function extractColsXml(string $entry = 'xl/worksheets/sheet1.xml'): string
    {
        $zip = new \ZipArchive();
        $zip->open($this->testFile);
        $sheet = $zip->getFromName($entry);
        $zip->close();
        if (preg_match('/<cols>(.*?)<\/cols>/s', $sheet, $m)) {
            return $m[0];
        }

        return '';
    }
}

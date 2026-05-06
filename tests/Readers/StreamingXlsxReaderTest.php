<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

class StreamingXlsxReaderTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-facade-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_from_file_factory_opens_archive(): void
    {
        $this->writeBasic();

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertNotEmpty($reader->sheets());
        $this->assertSame('STRATEGY_0_INLINE', $reader->strategy());
    }

    public function test_sheets_lists_every_sheet_in_workbook_order(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a']);
        $writer->writeRow(['x']);
        $writer->newSheet('Reports', ['b']);
        $writer->writeRow(['y']);
        $writer->newSheet('Audit', ['c']);
        $writer->writeRow(['z']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $sheets = $reader->sheets();

        $this->assertCount(3, $sheets);
        $this->assertSame('Reports', $sheets[1]['name']);
        $this->assertSame('Audit', $sheets[2]['name']);
    }

    public function test_default_sheet_is_the_first(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['main_a']);
        $writer->writeRow(['main-1']);
        $writer->newSheet('Other', ['other_a']);
        $writer->writeRow(['other-1']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame(['main_a'], $rows[0]);
        $this->assertSame(['main-1'], $rows[1]);
    }

    public function test_on_sheet_by_name_switches_active_sheet(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['s1']);
        $writer->writeRow(['sheet1-row']);
        $writer->newSheet('Other', ['s2']);
        $writer->writeRow(['sheet2-row']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->onSheet('Other')->rows(), false);

        $this->assertSame(['s2'], $rows[0]);
        $this->assertSame(['sheet2-row'], $rows[1]);
    }

    public function test_on_sheet_index_switches_by_position(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a']);
        $writer->writeRow(['first']);
        $writer->newSheet('Second', ['b']);
        $writer->writeRow(['second-row']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->onSheetIndex(1)->rows(), false);

        $this->assertSame(['b'], $rows[0]);
        $this->assertSame(['second-row'], $rows[1]);
    }

    public function test_on_sheet_unknown_name_throws(): void
    {
        $this->writeBasic();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(XlsxReadException::class);
        $reader->onSheet('NotARealSheet');
    }

    public function test_on_sheet_index_out_of_range_throws(): void
    {
        $this->writeBasic();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(XlsxReadException::class);
        $reader->onSheetIndex(99);
    }

    public function test_header_returns_first_row_and_caches(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'name']);
        $writer->writeRow([1, 'a']);
        $writer->writeRow([2, 'b']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $first = $reader->header();
        $second = $reader->header(); // must come from cache, not re-read

        $this->assertSame(['id', 'name'], $first);
        $this->assertSame($first, $second);
    }

    public function test_rows_skip_drops_leading_rows(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['n']);
        for ($i = 1; $i <= 5; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $data = iterator_to_array($reader->rows(skip: 1), false);

        $this->assertCount(5, $data);
        $this->assertSame(['1'], $data[0]); // first DATA row is row index 0
    }

    public function test_rows_limit_caps_emission(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['n']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $data = iterator_to_array($reader->rows(skip: 1, limit: 10), false);

        $this->assertCount(10, $data);
        $this->assertSame(['1'], $data[0]);
        $this->assertSame(['10'], $data[9]);
    }

    public function test_chunked_yields_batches_with_partial_last(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['n']);
        for ($i = 1; $i <= 25; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $batches = iterator_to_array($reader->chunked(10, skip: 1), false);

        $this->assertCount(3, $batches);
        $this->assertCount(10, $batches[0]);
        $this->assertCount(10, $batches[1]);
        $this->assertCount(5, $batches[2]); // remainder
        $this->assertSame(['1'], $batches[0][0]);
        $this->assertSame(['25'], $batches[2][4]);
    }

    public function test_chunked_invalid_size_rejected(): void
    {
        $this->writeBasic();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\InvalidArgumentException::class);
        iterator_to_array($reader->chunked(0));
    }

    public function test_row_count_reports_total_rows_including_header(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['n']);
        for ($i = 1; $i <= 137; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertSame(138, $reader->rowCount());
    }

    public function test_switching_sheet_invalidates_cached_header(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['s1_h']);
        $writer->writeRow(['x']);
        $writer->newSheet('Two', ['s2_h']);
        $writer->writeRow(['y']);
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertSame(['s1_h'], $reader->header());
        $reader->onSheet('Two');
        $this->assertSame(['s2_h'], $reader->header());
    }

    private function writeBasic(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a', 'b']);
        $writer->writeRow(['x', 'y']);
        $writer->finishFile();
    }
}

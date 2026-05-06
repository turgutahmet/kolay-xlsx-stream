<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Tests\TestCase;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

class StateGuardsTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir() . '/guards_test_' . uniqid() . '.xlsx';
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }

    public function test_write_row_before_start_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('startFile() first');

        $writer->writeRow(['x']);
    }

    public function test_finish_file_before_start_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('startFile() first');

        $writer->finishFile();
    }

    public function test_double_start_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('already been started');

        $writer->startFile(['A']);
    }

    public function test_double_finish_throws_controlled_exception()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('closed writer');

        $writer->finishFile();
    }

    public function test_write_after_finish_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('closed writer');

        $writer->writeRow([2]);
    }

    public function test_too_many_columns_in_headers_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('exceeds Excel');

        $writer->startFile(array_fill(0, 17000, 'col'));
    }

    public function test_too_many_columns_in_row_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('exceeds Excel');

        $writer->writeRow(array_fill(0, 17000, 'x'));
    }

    public function test_max_columns_at_limit_succeeds()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(array_fill(0, 16384, 'h'));
        $writer->writeRow(array_fill(0, 16384, 'v'));
        $stats = $writer->finishFile();

        $this->assertEquals(1, $stats['rows']);
        $this->assertFileExists($this->testFile);
    }
}

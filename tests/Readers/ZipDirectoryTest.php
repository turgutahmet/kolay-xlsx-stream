<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Foundation tests for the reader's ZIP container parser.
 *
 * Round-trips real XLSX files produced by SinkableXlsxWriter and asserts
 * that ZipDirectory recovers every entry the writer emitted.
 */
class ZipDirectoryTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-zipdir-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_parses_central_directory_of_minimal_xlsx(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a', 'b']);
        $writer->writeRow(['x', 'y']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        // Every XLSX must include these standard parts.
        $this->assertTrue($cd->has('[Content_Types].xml'));
        $this->assertTrue($cd->has('xl/workbook.xml'));
        $this->assertTrue($cd->has('xl/worksheets/sheet1.xml'));
        $this->assertTrue($cd->has('_rels/.rels'));
        $this->assertTrue($cd->has('xl/_rels/workbook.xml.rels'));
    }

    public function test_entry_metadata_is_consistent(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['col1', 'col2', 'col3']);
        for ($i = 0; $i < 100; $i++) {
            $writer->writeRow(["row{$i}-a", "row{$i}-b", "row{$i}-c"]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $sheet = $cd->entry('xl/worksheets/sheet1.xml');
        $this->assertNotNull($sheet);
        $this->assertGreaterThan(0, $sheet['compressed_size']);
        $this->assertGreaterThan($sheet['compressed_size'], $sheet['uncompressed_size']);
        $this->assertGreaterThan(0, $sheet['offset']);
        $this->assertSame(8, $sheet['method']); // DEFLATE
    }

    public function test_data_offset_points_to_inflatable_payload(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['greeting']);
        $writer->writeRow(['hello world']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $sheet = $cd->entry('xl/worksheets/sheet1.xml');
        $offset = $cd->dataOffset($source, 'xl/worksheets/sheet1.xml');
        $compressed = $source->range($offset, $sheet['compressed_size']);

        $inflated = inflate_init(ZLIB_ENCODING_RAW);
        $xml = inflate_add($inflated, $compressed, ZLIB_FINISH);

        $this->assertNotFalse($xml);
        $this->assertStringContainsString('<row r="1"', $xml);
        $this->assertStringContainsString('hello world', $xml);
    }

    public function test_entry_lookup_returns_null_for_missing_name(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['x']);
        $writer->writeRow(['y']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $this->assertNull($cd->entry('xl/no-such-thing.xml'));
        $this->assertFalse($cd->has('xl/no-such-thing.xml'));
    }

    public function test_data_offset_throws_for_unknown_entry(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['x']);
        $writer->writeRow(['y']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $this->expectException(XlsxReadException::class);
        $cd->dataOffset($source, 'xl/no-such-thing.xml');
    }

    public function test_rejects_non_zip_input(): void
    {
        file_put_contents($this->testFile, 'definitely not a zip archive');

        $source = new LocalFileSource($this->testFile);

        $this->expectException(XlsxReadException::class);
        ZipDirectory::fromSource($source);
    }

    public function test_local_file_source_reports_size_and_serves_ranges(): void
    {
        file_put_contents($this->testFile, str_repeat('abcdefghij', 1000)); // 10000 bytes

        $source = new LocalFileSource($this->testFile);
        $this->assertSame(10000, $source->size());

        $head = $source->range(0, 10);
        $this->assertSame('abcdefghij', $head);

        $mid = $source->range(50, 5);
        $this->assertSame('abcde', $mid);

        $tail = $source->range(9990, 10);
        $this->assertSame('abcdefghij', $tail);
    }

    public function test_local_file_source_stream_from_offset(): void
    {
        file_put_contents($this->testFile, str_repeat('0123456789', 100)); // 1000 bytes

        $source = new LocalFileSource($this->testFile);
        $stream = $source->streamFrom(990);

        $tail = '';
        while (! feof($stream)) {
            $chunk = fread($stream, 64);
            if ($chunk === false) {
                break;
            }
            $tail .= $chunk;
        }
        fclose($stream);

        $this->assertSame('0123456789', $tail);
    }
}

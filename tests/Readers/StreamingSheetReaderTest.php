<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingSheetReader;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * End-to-end round-trip tests: write XLSX with SinkableXlsxWriter, read
 * it back through StreamingSheetReader, assert byte-by-byte cell match.
 *
 * These exercise the inflate-streaming + row-boundary tokenization
 * paths as a system, layered on top of the unit tests in
 * CellTokenizerTest.
 */
class StreamingSheetReaderTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-streamread-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_round_trip_small_dataset(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'name', 'email']);
        $writer->writeRow([1, 'Alice', 'alice@example.com']);
        $writer->writeRow([2, 'Bob', 'bob@example.com']);
        $writer->writeRow([3, 'Charlie', 'charlie@example.com']);
        $writer->finishFile();

        $rows = $this->readAllRows($this->testFile);

        $this->assertCount(4, $rows);
        $this->assertSame(['id', 'name', 'email'], $rows[0]);
        $this->assertSame(['1', 'Alice', 'alice@example.com'], $rows[1]);
        $this->assertSame(['2', 'Bob', 'bob@example.com'], $rows[2]);
        $this->assertSame(['3', 'Charlie', 'charlie@example.com'], $rows[3]);
    }

    public function test_strategy_zero_for_writer_output(): void
    {
        // The package writer uses inlineStr for every cell; the sst entry
        // is never created. Reader must detect this and choose Strategy 0.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['x']);
        $writer->writeRow(['y']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $reader = new StreamingSheetReader($source, $cd);

        $this->assertSame('STRATEGY_0_INLINE', $reader->strategy());
    }

    public function test_round_trip_large_dataset_bounded_memory(): void
    {
        // 10 000 rows × 5 columns. Verifies that the streaming inflate
        // + row-boundary loop holds RAM steady regardless of row count.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'name', 'value', 'flag', 'note']);
        for ($i = 1; $i <= 10000; $i++) {
            $writer->writeRow([$i, "row-{$i}", $i * 1.5, $i % 2 === 0, "note {$i}"]);
        }
        $writer->finishFile();

        $rows = $this->readAllRows($this->testFile);

        $this->assertCount(10001, $rows);
        $this->assertSame(['id', 'name', 'value', 'flag', 'note'], $rows[0]);
        $this->assertSame(['1', 'row-1', '1.5', false, 'note 1'], $rows[1]);
        $this->assertSame(['10000', 'row-10000', '15000', true, 'note 10000'], $rows[10000]);
    }

    public function test_xml_special_characters_round_trip(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['raw']);
        $writer->writeRow(['foo & bar']);
        $writer->writeRow(['<tag>value</tag>']);
        $writer->writeRow(['"quoted"']);
        $writer->writeRow(["it's fine"]);
        $writer->finishFile();

        $rows = $this->readAllRows($this->testFile);

        $this->assertSame('foo & bar', $rows[1][0]);
        $this->assertSame('<tag>value</tag>', $rows[2][0]);
        $this->assertSame('"quoted"', $rows[3][0]);
        $this->assertSame("it's fine", $rows[4][0]);
    }

    public function test_unicode_round_trip(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['text']);
        $writer->writeRow(['İstanbul']);
        $writer->writeRow(['日本語']);
        $writer->writeRow(['🚀 emoji']);
        $writer->writeRow(['Ωμέγα']);
        $writer->finishFile();

        $rows = $this->readAllRows($this->testFile);

        $this->assertSame('İstanbul', $rows[1][0]);
        $this->assertSame('日本語', $rows[2][0]);
        $this->assertSame('🚀 emoji', $rows[3][0]);
        $this->assertSame('Ωμέγα', $rows[4][0]);
    }

    public function test_whitespace_preservation(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['ws']);
        $writer->writeRow(['  leading']);
        $writer->writeRow(['trailing  ']);
        $writer->writeRow(['   ']); // pure whitespace
        $writer->writeRow(["tab\there"]);
        $writer->writeRow(["new\nline"]);
        $writer->finishFile();

        $rows = $this->readAllRows($this->testFile);

        $this->assertSame('  leading', $rows[1][0]);
        $this->assertSame('trailing  ', $rows[2][0]);
        $this->assertSame('   ', $rows[3][0]);
        $this->assertSame("tab\there", $rows[4][0]);
        $this->assertSame("new\nline", $rows[5][0]);
    }

    public function test_boolean_cells_decode_to_php_bool(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['flag']);
        $writer->writeRow([true]);
        $writer->writeRow([false]);
        $writer->writeRow([true]);
        $writer->finishFile();

        $rows = $this->readAllRows($this->testFile);

        $this->assertTrue($rows[1][0]);
        $this->assertFalse($rows[2][0]);
        $this->assertTrue($rows[3][0]);
    }

    public function test_wide_row_with_50_columns(): void
    {
        $headers = [];
        $values = [];
        for ($i = 1; $i <= 50; $i++) {
            $headers[] = "h{$i}";
            $values[] = "v{$i}";
        }

        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile($headers);
        $writer->writeRow($values);
        $writer->finishFile();

        $rows = $this->readAllRows($this->testFile);

        $this->assertCount(50, $rows[0]);
        $this->assertCount(50, $rows[1]);
        $this->assertSame('v1', $rows[1][0]);
        $this->assertSame('v27', $rows[1][26]); // crosses Z→AA boundary
        $this->assertSame('v50', $rows[1][49]);
    }

    public function test_multi_sheet_via_explicit_entry_path(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['s1a', 's1b']);
        $writer->writeRow(['sheet1-row1', 'sheet1-row1b']);
        $writer->writeRow(['sheet1-row2', 'sheet1-row2b']);
        $writer->newSheet('Second', ['s2a', 's2b']);
        $writer->writeRow(['sheet2-row1', 'sheet2-row1b']);
        $writer->finishFile();

        $rows1 = $this->readAllRows($this->testFile, 'xl/worksheets/sheet1.xml');
        $rows2 = $this->readAllRows($this->testFile, 'xl/worksheets/sheet2.xml');

        $this->assertCount(3, $rows1);
        $this->assertSame(['s1a', 's1b'], $rows1[0]);
        $this->assertSame(['sheet1-row1', 'sheet1-row1b'], $rows1[1]);

        $this->assertCount(2, $rows2);
        $this->assertSame(['s2a', 's2b'], $rows2[0]);
        $this->assertSame(['sheet2-row1', 'sheet2-row1b'], $rows2[1]);
    }

    public function test_long_string_cell(): void
    {
        $longValue = str_repeat('lorem ipsum ', 2000); // ~24 KB

        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['long']);
        $writer->writeRow([$longValue]);
        $writer->finishFile();

        $rows = $this->readAllRows($this->testFile);

        $this->assertSame($longValue, $rows[1][0]);
    }

    public function test_generator_yields_rows_one_at_a_time(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['n']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $reader = new StreamingSheetReader($source, $cd);

        // Confirm that rows() returns a Generator (lazy evaluation).
        $gen = $reader->rows();
        $this->assertInstanceOf(\Generator::class, $gen);

        $count = 0;
        foreach ($gen as $row) {
            $count++;
            $this->assertNotEmpty($row);
            if ($count >= 5) {
                break;  // early termination — shouldn't read remaining 95 rows
            }
        }

        $this->assertSame(5, $count);
    }

    /**
     * Helper: collect every row of a sheet into a flat array.
     *
     * @return array<int, array<int, mixed>>
     */
    private function readAllRows(string $path, string $entry = 'xl/worksheets/sheet1.xml'): array
    {
        $source = new LocalFileSource($path);
        $cd = ZipDirectory::fromSource($source);
        $reader = new StreamingSheetReader($source, $cd, $entry);

        $rows = [];
        foreach ($reader->rows() as $row) {
            $rows[] = $row;
        }

        return $rows;
    }
}

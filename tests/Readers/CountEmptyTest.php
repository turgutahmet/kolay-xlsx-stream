<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * countEmpty() — data rows in a column that hold no numeric value (blank or
 * non-numeric), answered from the STAT zone maps alone. It is the missing
 * side of columnStats()['count']: the header is excluded and the two always
 * sum to the data-row count. Gates: blank-only oracle, blanks + stray text
 * (both are "non-numeric"), the all-numeric zero, and the null contract.
 */
class CountEmptyTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-empty-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_counts_blank_cells_in_a_numeric_column(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        $blanks = 0;
        for ($i = 1; $i <= 500; $i++) {
            // Every 7th row leaves amount blank.
            $blank = $i % 7 === 0;
            $blanks += $blank ? 1 : 0;
            $writer->writeRow([$i, $blank ? '' : (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSame($blanks, $reader->countEmpty('amount'));
        // Complement identity: numeric count + empty == data rows.
        $this->assertSame(500, $reader->columnStats('amount')['count'] + $reader->countEmpty('amount'));
        $reader->close();
    }

    public function test_non_numeric_text_counts_as_empty(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        $missing = 0;
        for ($i = 1; $i <= 300; $i++) {
            if ($i % 10 === 0) {
                $writer->writeRow([$i, '']);       // blank
                $missing++;
            } elseif ($i % 10 === 5) {
                $writer->writeRow([$i, 'N/A']);    // non-numeric text
                $missing++;
            } else {
                $writer->writeRow([$i, (float) $i]);
            }
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSame($missing, $reader->countEmpty(2), 'blanks and non-numeric text both count');
        $reader->close();
    }

    public function test_all_numeric_column_is_zero(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSame(0, $reader->countEmpty('amount'));
        $reader->close();
    }

    public function test_null_when_column_untracked(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->countEmpty(1), 'untracked column has no STAT to count from');
        $reader->close();
    }
}

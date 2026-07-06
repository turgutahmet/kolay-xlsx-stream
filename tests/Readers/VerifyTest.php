<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Z3 — verify(): read-side integrity check against the writer's per-block
 * (SCRC) and whole-sheet CRC pins. A clean file verifies ok; a byte
 * flipped inside a sheet's compressed data (the ZIP central directory
 * left intact, as a real storage bit-flip would leave it) is caught, and
 * on a born-indexed file the damage is localized to a block.
 */
class VerifyTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-verify-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    private function writeFixture(bool $indexed = true): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        if ($indexed) {
            $writer->withRandomAccessIndex(every: 100);
        }
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'name']);
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([$i, 'row-'.$i]);
        }
        $writer->finishFile();
    }

    /** Flip one byte inside sheet1.xml's COMPRESSED data, leaving the ZIP CD intact. */
    private function corruptSheetData(): void
    {
        $data = file_get_contents($this->testFile);
        $name = 'xl/worksheets/sheet1.xml';
        $fnPos = strpos($data, $name);          // first hit = the local file header's name field
        $this->assertNotFalse($fnPos);
        $header = $fnPos - 30;                    // local header is 30 fixed bytes + name + extra
        $this->assertSame("PK\x03\x04", substr($data, $header, 4), 'expected a local file header');
        $compSize = unpack('V', substr($data, $header + 18, 4))[1];
        $fnLen = unpack('v', substr($data, $header + 26, 2))[1];
        $extraLen = unpack('v', substr($data, $header + 28, 2))[1];
        $dataStart = $header + 30 + $fnLen + $extraLen;
        $flipAt = $dataStart + intdiv($compSize, 2); // mid-stream -> corrupts a middle block
        $data[$flipAt] = chr(ord($data[$flipAt]) ^ 0xFF);
        file_put_contents($this->testFile, $data);
    }

    public function test_clean_indexed_file_verifies_ok(): void
    {
        $this->writeFixture(indexed: true);
        $report = StreamingXlsxReader::fromFile($this->testFile)->verify();

        $this->assertTrue($report['ok']);
        $this->assertCount(1, $report['sheets']);
        $sheet = $report['sheets'][0];
        $this->assertTrue($sheet['ok']);
        $this->assertTrue($sheet['sheet_crc_ok']);
        $this->assertTrue($sheet['inflate_ok']);
        $this->assertSame([], $sheet['corrupt_blocks']);
        $this->assertGreaterThan(0, $sheet['blocks_checked']); // SCRC pins present
    }

    public function test_clean_unindexed_file_verifies_ok(): void
    {
        $this->writeFixture(indexed: false);
        $report = StreamingXlsxReader::fromFile($this->testFile)->verify();

        $this->assertTrue($report['ok']);
        $this->assertSame(0, $report['sheets'][0]['blocks_checked']); // whole-entry CRC only
        $this->assertTrue($report['sheets'][0]['sheet_crc_ok']);
    }

    public function test_corrupted_sheet_bytes_are_detected(): void
    {
        $this->writeFixture(indexed: true);
        $this->corruptSheetData();

        $report = StreamingXlsxReader::fromFile($this->testFile)->verify();

        $this->assertFalse($report['ok'], 'a flipped byte in the sheet data must fail verification');
        $sheet = $report['sheets'][0];
        $this->assertFalse($sheet['ok']);
        // Either a block CRC mismatched, the whole-sheet CRC mismatched, or
        // the stream was too damaged to inflate — any of these means "bad".
        $this->assertTrue(
            $sheet['corrupt_blocks'] !== [] || ! $sheet['sheet_crc_ok'] || ! $sheet['inflate_ok'],
            'corruption must surface as a bad block, a bad sheet CRC, or a failed inflate'
        );
    }

    public function test_corruption_is_localized_to_a_block(): void
    {
        // With SCRC pins a mid-stream flip that still inflates should name a
        // specific block rather than only failing the whole-sheet CRC.
        $this->writeFixture(indexed: true);
        $this->corruptSheetData();

        $report = StreamingXlsxReader::fromFile($this->testFile)->verify();
        $sheet = $report['sheets'][0];

        // If it inflated, at least one SCRC checkpoint must have caught it.
        if ($sheet['inflate_ok']) {
            $this->assertNotEmpty($sheet['corrupt_blocks'], 'a block should be pinpointed when the stream still inflates');
        } else {
            $this->assertTrue(true); // too damaged to inflate — still correctly reported as bad
        }
    }
}

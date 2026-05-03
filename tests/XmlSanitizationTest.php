<?php

namespace Kolay\XlsxStream\Tests;

use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * Regression coverage for the v2.2.2 control-byte fast-path bug.
 *
 * Pre-fix, the strpbrk needle was a single-quoted string literal so
 * "\x00" landed as the four characters "\", "x", "0", "0" instead of a
 * real null byte. Inputs whose characters didn't overlap with
 * "&<>\"'\\x0..9A..F" sailed past sanitization and Excel rejected the
 * resulting workbook with "Char 0x0 out of allowed range".
 */
class XmlSanitizationTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/sanitize_'.uniqid().'.xlsx';
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }

    public function test_pure_lowercase_with_control_byte_is_stripped()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['Dirty']);
        $writer->writeRow(["abc\x00def"]);
        $writer->finishFile();

        $xml = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertNoControlBytes($xml);
        $this->assertStrictXmlParses($xml);
    }

    public function test_pure_control_bytes_are_stripped()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['Dirty']);
        $writer->writeRow(["\x00\x01\x02\x03\x04\x05\x06\x07\x08"]);
        $writer->writeRow(["\x0B\x0C\x0E\x0F\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F"]);
        $writer->finishFile();

        $xml = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertNoControlBytes($xml);
        $this->assertStrictXmlParses($xml);
    }

    public function test_mixed_case_with_control_bytes_strict_parses()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['Mixed']);
        $writer->writeRow(["lowercase\x00MIXED"]);
        $writer->writeRow(["UPPER\x07lower"]);
        $writer->writeRow(["data\x01with\x02breaks"]);
        $writer->finishFile();

        $xml = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertNoControlBytes($xml);
        $this->assertStrictXmlParses($xml);
    }

    public function test_control_bytes_in_header_are_stripped()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(["Bad\x00Header"]);
        $writer->writeRow(['ok']);
        $writer->finishFile();

        $xml = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertNoControlBytes($xml);
        $this->assertStrictXmlParses($xml);
    }

    public function test_tab_newline_carriage_return_are_preserved()
    {
        // \t (0x09), \n (0x0A), \r (0x0D) are valid XML chars and should
        // pass through unchanged — we only strip the C0 control set
        // *minus* those three.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['Whitespace']);
        $writer->writeRow(["line1\nline2"]);
        $writer->writeRow(["before\ttab"]);
        $writer->writeRow(["cr\rhere"]);
        $writer->finishFile();

        $xml = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertStrictXmlParses($xml);
        $this->assertStringContainsString("line1\nline2", $xml);
        $this->assertStringContainsString("before\ttab", $xml);
    }

    public function test_special_xml_chars_still_escaped()
    {
        // Sanity: the slow path still escapes &<>"' correctly.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['Special']);
        $writer->writeRow(['<tag>&"\'']);
        $writer->finishFile();

        $xml = $this->extract('xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('&lt;tag&gt;&amp;&quot;&apos;', $xml);
        $this->assertStrictXmlParses($xml);
    }

    private function assertNoControlBytes(string $xml): void
    {
        $this->assertDoesNotMatchRegularExpression(
            '/[\x00-\x08\x0B\x0C\x0E-\x1F]/',
            $xml,
            'Output contains forbidden XML 1.0 control bytes'
        );
    }

    private function assertStrictXmlParses(string $xml): void
    {
        libxml_use_internal_errors(true);
        libxml_clear_errors();
        $parsed = simplexml_load_string($xml);
        $errors = array_map(fn ($e) => trim($e->message), libxml_get_errors());
        libxml_clear_errors();

        $this->assertNotFalse(
            $parsed,
            'libxml strict parse failed: '.implode(' | ', $errors)
        );
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

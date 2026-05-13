<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Tests\TestCase;

use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * Verifies that v2.0 data type fixes produce well-formed XLSX cells.
 */
class DataTypesTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir() . '/types_test_' . uniqid() . '.xlsx';
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }

    public function test_boolean_true_becomes_native_excel_boolean()
    {
        $row = $this->buildRowAndExtract([true]);
        $this->assertStringContainsString('t="b"><v>1</v>', $row);
        $this->assertStringNotContainsString('inlineStr', $row);
    }

    public function test_boolean_false_becomes_native_excel_boolean()
    {
        $row = $this->buildRowAndExtract([false]);
        $this->assertStringContainsString('t="b"><v>0</v>', $row);
    }

    public function test_datetime_becomes_excel_serial_with_datetime_style()
    {
        $dt = new \DateTime('2026-01-15 10:30:00', new \DateTimeZone('UTC'));
        $row = $this->buildRowAndExtract([$dt]);

        // 2026-01-15 10:30:00 UTC -> serial 46037.4375
        $this->assertStringContainsString('s="1"', $row);
        $this->assertStringContainsString('t="n"', $row);
        $this->assertStringContainsString('46037.4375', $row);
    }

    public function test_datetime_immutable_supported()
    {
        $dt = new \DateTimeImmutable('2026-01-15 00:00:00', new \DateTimeZone('UTC'));
        $row = $this->buildRowAndExtract([$dt]);

        $this->assertStringContainsString('s="1"', $row);
        $this->assertStringContainsString('46037', $row);
    }

    public function test_long_numeric_string_preserved_without_precision_loss()
    {
        $row = $this->buildRowAndExtract(['12345678901234567890']);

        $this->assertStringContainsString('inlineStr', $row);
        $this->assertStringContainsString('12345678901234567890', $row);
        $this->assertStringNotContainsString('E+', $row);
        $this->assertStringNotContainsString('e+', $row);
    }

    public function test_leading_zero_string_preserved()
    {
        $row = $this->buildRowAndExtract(['00123']);
        $this->assertStringContainsString('inlineStr', $row);
        $this->assertStringContainsString('00123', $row);
    }

    public function test_plus_prefixed_string_preserved()
    {
        $row = $this->buildRowAndExtract(['+90555']);
        $this->assertStringContainsString('inlineStr', $row);
        $this->assertStringContainsString('+90555', $row);
    }

    public function test_short_numeric_string_still_treated_as_number()
    {
        $row = $this->buildRowAndExtract(['42']);
        $this->assertStringContainsString('t="n"><v>42</v>', $row);
        $this->assertStringNotContainsString('inlineStr', $row);
    }

    public function test_decimal_string_still_treated_as_number()
    {
        $row = $this->buildRowAndExtract(['3.14']);
        $this->assertStringContainsString('t="n"><v>3.14</v>', $row);
    }

    public function test_styles_xml_includes_datetime_format()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $zip = new ZipArchive();
        $zip->open($this->testFile);
        $styles = $zip->getFromName('xl/styles.xml');
        $zip->close();

        $this->assertStringContainsString('numFmtId="164"', $styles);
        $this->assertStringContainsString('yyyy-mm-dd', $styles);
        $this->assertStringContainsString('cellXfs count="2"', $styles);
    }

    public function test_zip_remains_valid_with_mixed_types()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        $writer->startFile(['Str', 'Int', 'BigInt', 'Bool', 'DateTime']);
        $writer->writeRow([
            'hello',
            42,
            '99999999999999999999',
            true,
            new \DateTime('2026-01-15 10:30:00', new \DateTimeZone('UTC')),
        ]);
        $writer->finishFile();

        $zip = new ZipArchive();
        $this->assertTrue($zip->open($this->testFile, ZipArchive::CHECKCONS) === true);
        $zip->close();
    }

    private function buildRowAndExtract(array $row): string
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        $writer->startFile(array_fill(0, count($row), 'H'));
        $writer->writeRow($row);
        $writer->finishFile();

        $zip = new ZipArchive();
        $zip->open($this->testFile);
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        preg_match('#<row r="2">.*?</row>#s', $xml, $m);
        return $m[0] ?? '';
    }
}

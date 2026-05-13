<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\BaseXlsxWriter;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * setColumnFormat(int $col, int $builtinNumFmtId) — locale-aware
 * column formatting via Excel's reserved 0-49 numFmtId range.
 *
 * Pinned behaviours:
 *   - cellXf carries numFmtId="N" with N in [0,49]
 *   - styles.xml does NOT contain a <numFmt> entry for the built-in
 *   - Cells in that column render with s="cellXf-index"
 *   - Range validation rejects N < 0 and N > 49
 */
class BuiltinNumFmtTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-builtin-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_builtin_numFmtId_emits_cellXf_without_custom_numFmt_entry(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, BaseXlsxWriter::BUILTIN_NUMFMT_DATE); // 14
        $writer->startFile(['when']);
        $writer->writeRow([new \DateTimeImmutable('2026-05-06', new \DateTimeZone('UTC'))]);
        $writer->finishFile();

        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->testFile) === true);
        $styles = $zip->getFromName('xl/styles.xml');
        $sheet = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        // cellXf with numFmtId=14 must exist
        $this->assertMatchesRegularExpression(
            '/<xf\s+numFmtId="14"[^\/>]*applyNumberFormat="1"/',
            $styles
        );

        // No <numFmt formatCode="..." numFmtId="14"...> entry — Excel
        // reads the numFmtId straight from its built-in table.
        $this->assertStringNotContainsString('numFmtId="14" formatCode=', $styles);
        $this->assertStringNotContainsString('formatCode="" numFmtId="14"', $styles);

        // Data row references the registered cellXf via s="N"
        $this->assertMatchesRegularExpression('/<c r="A2" s="\d+" t="n"><v>/', $sheet);
    }

    public function test_builtin_numFmtId_currency_renders_with_id_5(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, BaseXlsxWriter::BUILTIN_NUMFMT_CURRENCY);
        $writer->startFile(['amount']);
        $writer->writeRow([1234.56]);
        $writer->finishFile();

        $zip = new \ZipArchive();
        $zip->open($this->testFile);
        $styles = $zip->getFromName('xl/styles.xml');
        $zip->close();

        $this->assertMatchesRegularExpression('/<xf\s+numFmtId="5"[^\/>]*applyNumberFormat="1"/', $styles);
    }

    public function test_builtin_numFmtId_below_zero_throws(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('Built-in numFmtId must be 0-49');
        $writer->setColumnFormat(1, -1);
    }

    public function test_builtin_numFmtId_above_49_throws(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('Built-in numFmtId must be 0-49');
        $writer->setColumnFormat(1, 50);
    }

    public function test_builtin_numFmtId_high_boundary_is_accepted(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, 49);
        $writer->startFile(['x']);
        $writer->writeRow([1.0]);
        $writer->finishFile();

        $zip = new \ZipArchive();
        $zip->open($this->testFile);
        $styles = $zip->getFromName('xl/styles.xml');
        $zip->close();

        $this->assertMatchesRegularExpression('/numFmtId="49"/', $styles);
    }

    public function test_string_preset_path_still_emits_custom_numFmt(): void
    {
        // Sanity check: the string overload is unchanged. A preset like
        // 'currency_try' (literal symbol) must still emit a custom <numFmt>.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, 'currency_try');
        $writer->startFile(['x']);
        $writer->writeRow([1.0]);
        $writer->finishFile();

        $zip = new \ZipArchive();
        $zip->open($this->testFile);
        $styles = $zip->getFromName('xl/styles.xml');
        $zip->close();

        $this->assertStringContainsString('<numFmt', $styles);
        $this->assertStringContainsString('₺', $styles);
    }

    public function test_repeated_builtin_assignments_share_cellXf(): void
    {
        // Idempotency: registering the same builtin twice must not grow
        // the cellXfs table — ensures style-ids stay stable across columns.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setColumnFormat(1, BaseXlsxWriter::BUILTIN_NUMFMT_DATE);
        $writer->setColumnFormat(2, BaseXlsxWriter::BUILTIN_NUMFMT_DATE);
        $writer->startFile(['a', 'b']);
        $writer->writeRow([1, 2]);
        $writer->finishFile();

        $zip = new \ZipArchive();
        $zip->open($this->testFile);
        $styles = $zip->getFromName('xl/styles.xml');
        $zip->close();

        $count = preg_match_all('/numFmtId="14"/', $styles);
        $this->assertSame(1, $count, 'BUILTIN_NUMFMT_DATE should yield exactly one cellXf');
    }
}

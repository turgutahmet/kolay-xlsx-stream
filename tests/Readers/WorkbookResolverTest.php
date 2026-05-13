<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Readers\WorkbookResolver;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

class WorkbookResolverTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-resolver-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_single_sheet_default_name(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a']);
        $writer->writeRow(['x']);
        $writer->finishFile();

        $resolved = $this->resolve($this->testFile);

        $this->assertCount(1, $resolved);
        $this->assertSame('xl/worksheets/sheet1.xml', $resolved[0]['entry']);
        $this->assertNotEmpty($resolved[0]['name']);
        $this->assertGreaterThan(0, $resolved[0]['sheetId']);
    }

    public function test_multi_sheet_preserves_workbook_order(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['s1']);
        $writer->writeRow(['a']);
        $writer->newSheet('Reports', ['s2']);
        $writer->writeRow(['b']);
        $writer->newSheet('Audit', ['s3']);
        $writer->writeRow(['c']);
        $writer->finishFile();

        $resolved = $this->resolve($this->testFile);

        $this->assertCount(3, $resolved);
        // Names preserved in workbook order
        $names = array_column($resolved, 'name');
        $this->assertSame('Reports', $names[1]);
        $this->assertSame('Audit', $names[2]);
        // Entries match the canonical worksheet path pattern
        foreach ($resolved as $i => $sheet) {
            $this->assertSame('xl/worksheets/sheet'.($i + 1).'.xml', $sheet['entry']);
        }
    }

    public function test_sheet_entries_are_present_in_archive(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['col']);
        $writer->writeRow(['v']);
        $writer->newSheet('Other', ['col2']);
        $writer->writeRow(['w']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $resolved = WorkbookResolver::resolve($source, $cd);

        foreach ($resolved as $sheet) {
            $this->assertTrue(
                $cd->has($sheet['entry']),
                "resolved entry {$sheet['entry']} must exist in CD"
            );
        }
    }

    public function test_throws_when_workbook_xml_missing(): void
    {
        // Build a minimal ZIP with no workbook.xml — use ZipArchive directly.
        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->testFile, \ZipArchive::CREATE) === true);
        $zip->addFromString('placeholder.txt', 'nothing useful');
        $zip->close();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $this->expectException(XlsxReadException::class);
        WorkbookResolver::resolve($source, $cd);
    }

    /**
     * @return list<array{name: string, sheetId: int, entry: string}>
     */
    private function resolve(string $path): array
    {
        $source = new LocalFileSource($path);
        $cd = ZipDirectory::fromSource($source);

        return WorkbookResolver::resolve($source, $cd);
    }
}

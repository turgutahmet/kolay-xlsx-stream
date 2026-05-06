<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Tests\TestCase;
use ZipArchive;

/**
 * Tests for reading XLSX archives that use the shared-strings table.
 *
 * SinkableXlsxWriter only emits inline strings, so to exercise the
 * t="s" code path each test builds a minimal OPC archive matching the
 * shape PhpSpreadsheet, openpyxl, Apache POI, etc. produce. The shared
 * strings table is small in every test (a few KB at most) — within the
 * package's bounded-RAM contract, no memory_limit jugglery.
 */
class ExternalXlsxTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-external-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_reader_resolves_shared_string_references(): void
    {
        $this->buildExternalXlsx(
            sharedStrings: ['Istanbul', 'Ankara', 'İzmir'],
            sheetRows: [
                [['t' => 's', 'v' => '0']], // header → "Istanbul"
                [['t' => 's', 'v' => '1'], ['t' => 'n', 'v' => '42']],
                [['t' => 's', 'v' => '2'], ['t' => 'n', 'v' => '7']],
            ]
        );

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertCount(3, $rows);
        $this->assertSame(['Istanbul'], $rows[0]);
        $this->assertSame(['Ankara', '42'], $rows[1]);
        $this->assertSame(['İzmir', '7'], $rows[2]);
    }

    public function test_unicode_and_entities_in_shared_strings_round_trip(): void
    {
        $this->buildExternalXlsx(
            sharedStrings: ['foo & bar', '<tag>', 'İstanbul 🌊'],
            sheetRows: [
                [['t' => 's', 'v' => '0']],
                [['t' => 's', 'v' => '1']],
                [['t' => 's', 'v' => '2']],
            ]
        );

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame(['foo & bar'], $rows[0]);
        $this->assertSame(['<tag>'], $rows[1]);
        $this->assertSame(['İstanbul 🌊'], $rows[2]);
    }

    public function test_mixed_inline_and_shared_strings_in_same_sheet(): void
    {
        // Some external writers mix inline strings (long text) with sst
        // references (short repeated values) in the same sheet. Reader
        // must handle both within a single row.
        $this->buildExternalXlsx(
            sharedStrings: ['Active', 'Pending'],
            sheetRows: [
                [
                    ['t' => 'inlineStr', 'is' => 'first user'],
                    ['t' => 's', 'v' => '0'],
                    ['t' => 'n', 'v' => '100'],
                ],
                [
                    ['t' => 'inlineStr', 'is' => 'second user'],
                    ['t' => 's', 'v' => '1'],
                    ['t' => 'n', 'v' => '200'],
                ],
            ]
        );

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(), false);

        $this->assertSame(['first user', 'Active', '100'], $rows[0]);
        $this->assertSame(['second user', 'Pending', '200'], $rows[1]);
    }

    public function test_chunked_works_on_external_xlsx(): void
    {
        $sst = [];
        for ($i = 0; $i < 50; $i++) {
            $sst[] = "row-{$i}";
        }
        $rows = [];
        for ($i = 0; $i < 50; $i++) {
            $rows[] = [['t' => 's', 'v' => (string) $i]];
        }

        $this->buildExternalXlsx(sharedStrings: $sst, sheetRows: $rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $batches = iterator_to_array($reader->chunked(20), false);

        $this->assertCount(3, $batches);
        $this->assertCount(20, $batches[0]);
        $this->assertCount(20, $batches[1]);
        $this->assertCount(10, $batches[2]);
        $this->assertSame(['row-0'], $batches[0][0]);
        $this->assertSame(['row-49'], $batches[2][9]);
    }

    /**
     * @param  list<string>  $sharedStrings
     * @param  list<list<array{t: string, v?: string, is?: string}>>  $sheetRows
     */
    private function buildExternalXlsx(array $sharedStrings, array $sheetRows): void
    {
        $zip = new ZipArchive();
        $opened = $zip->open($this->testFile, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        $this->assertTrue($opened === true);

        $zip->addFromString('[Content_Types].xml', $this->contentTypes());
        $zip->addFromString('_rels/.rels', $this->packageRels());
        $zip->addFromString('xl/workbook.xml', $this->workbookXml());
        $zip->addFromString('xl/_rels/workbook.xml.rels', $this->workbookRels());
        $zip->addFromString('xl/sharedStrings.xml', $this->buildSstXml($sharedStrings));
        $zip->addFromString('xl/worksheets/sheet1.xml', $this->buildSheetXml($sheetRows));

        $zip->close();
    }

    /**
     * @param  list<string>  $strings
     */
    private function buildSstXml(array $strings): string
    {
        $count = count($strings);
        $body = '';
        foreach ($strings as $s) {
            $encoded = htmlspecialchars($s, ENT_QUOTES | ENT_XML1, 'UTF-8');
            $body .= '<si><t xml:space="preserve">'.$encoded.'</t></si>';
        }

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'.
            ' count="'.$count.'" uniqueCount="'.$count.'">'.
            $body.'</sst>';
    }

    /**
     * @param  list<list<array{t: string, v?: string, is?: string}>>  $rows
     */
    private function buildSheetXml(array $rows): string
    {
        $body = '';
        foreach ($rows as $rowIdx => $cells) {
            $rowNum = $rowIdx + 1;
            $cellXml = '';
            foreach ($cells as $colIdx => $cell) {
                $ref = $this->columnLetters($colIdx).$rowNum;
                if ($cell['t'] === 'inlineStr') {
                    $text = htmlspecialchars($cell['is'] ?? '', ENT_QUOTES | ENT_XML1, 'UTF-8');
                    $cellXml .= '<c r="'.$ref.'" t="inlineStr"><is><t>'.$text.'</t></is></c>';
                } else {
                    $cellXml .= '<c r="'.$ref.'" t="'.$cell['t'].'"><v>'.($cell['v'] ?? '').'</v></c>';
                }
            }
            $body .= '<row r="'.$rowNum.'">'.$cellXml.'</row>';
        }

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<sheetData>'.$body.'</sheetData>'.
            '</worksheet>';
    }

    private function columnLetters(int $col0): string
    {
        $col = $col0 + 1;
        $letters = '';
        while ($col > 0) {
            $col--;
            $letters = chr(65 + ($col % 26)).$letters;
            $col = intdiv($col, 26);
        }

        return $letters;
    }

    private function contentTypes(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'.
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'.
            '<Default Extension="xml" ContentType="application/xml"/>'.
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'.
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'.
            '</Types>';
    }

    private function packageRels(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'.
            '</Relationships>';
    }

    private function workbookXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'.
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.
            '<sheets>'.
            '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'.
            '</sheets>'.
            '</workbook>';
    }

    private function workbookRels(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'.
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'.
            '</Relationships>';
    }
}

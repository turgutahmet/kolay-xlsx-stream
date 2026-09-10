<?php

namespace Kolay\XlsxStream\Tests\Templates;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Templates\Template;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Template parsing — the four-point seam's first point: cutting a foreign
 * sheet at <sheetData>. The package must NOT interpret the template; it
 * opens the seam, reads the sample rows as a style-id oracle, and keeps
 * every other byte verbatim.
 *
 * The matrix covers three producer dialects (xlsx-stream, a
 * PhpSpreadsheet-shaped sheet, an Excel-shaped namespace-prefixed sheet)
 * plus the shapes that break naive parsers: self-closing <sheetData/>,
 * rows without an r attribute, and customHeight preceding ht.
 */
class TemplateParseTest extends TestCase
{
    /** @var list<string> */
    private array $tmp = [];

    protected function tearDown(): void
    {
        foreach ($this->tmp as $f) {
            @unlink($f);
        }
        $this->tmp = [];
        parent::tearDown();
    }

    // ─── fixture plumbing ────────────────────────────────────────────────

    /** @param array<string,string> $parts entry name => contents */
    private function buildXlsx(array $parts): string
    {
        $path = sys_get_temp_dir().'/kxs-tpl-'.uniqid('', true).'.xlsx';
        $this->tmp[] = $path;
        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($path, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true);
        foreach ($parts as $name => $body) {
            $zip->addFromString($name, $body);
        }
        $zip->close();

        return file_get_contents($path);
    }

    /** Minimal workbook scaffolding around one or more sheet XML bodies. */
    private function scaffold(array $sheets, array $extra = []): array
    {
        $sheetTags = '';
        $rels = '';
        $overrides = '';
        $i = 0;
        foreach ($sheets as $name => $_) {
            $i++;
            $sheetTags .= '<sheet name="'.$name.'" sheetId="'.$i.'" r:id="rId'.$i.'"/>';
            $rels .= '<Relationship Id="rId'.$i.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'.$i.'.xml"/>';
            $overrides .= '<Override PartName="/xl/worksheets/sheet'.$i.'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }

        $parts = [
            '[Content_Types].xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                .'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                .'<Default Extension="xml" ContentType="application/xml"/>'
                .'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                .$overrides.'</Types>',
            '_rels/.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                .'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                .'</Relationships>',
            'xl/workbook.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                .'<sheets>'.$sheetTags.'</sheets></workbook>',
            'xl/_rels/workbook.xml.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.$rels.'</Relationships>',
        ];
        $i = 0;
        foreach ($sheets as $body) {
            $parts['xl/worksheets/sheet'.(++$i).'.xml'] = $body;
        }

        return $parts + $extra;
    }

    /** PhpSpreadsheet-shaped: dimension, sheetFormatPr, cols, frozen pane, merge tail. */
    private function phpSpreadsheetSheet(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<sheetPr/><dimension ref="A1:D4"/>'
            .'<sheetViews><sheetView tabSelected="1" workbookViewId="0">'
            .'<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            .'<sheetFormatPr defaultRowHeight="14.4"/>'
            .'<cols><col min="1" max="1" width="12.5" customWidth="1"/></cols>'
            .'<sheetData>'
            .'<row r="1" spans="1:4" ht="24" customHeight="1"><c r="A1" s="3" t="s"><v>0</v></c><c r="B1" s="3" t="s"><v>1</v></c></row>'
            .'<row r="2" spans="1:4"><c r="A2" s="4" t="s"><v>2</v></c></row>'
            // sample variant 0 — customHeight BEFORE ht: the naive /ht="/ trap
            .'<row r="3" spans="1:4" customHeight="1" ht="18.75"><c r="A3" s="5"/><c r="B3" s="6"/><c r="D3" s="9"/></row>'
            // sample variant 1 — zebra
            .'<row r="4" spans="1:4" customHeight="1" ht="18.75"><c r="A4" s="7"/><c r="B4" s="8"/><c r="D4" s="10"/></row>'
            .'</sheetData>'
            .'<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>'
            .'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            .'</worksheet>';
    }

    /** Excel-shaped: everything namespace-prefixed with x:. */
    private function excelPrefixedSheet(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<x:dimension ref="A1:B3"/>'
            .'<x:sheetViews><x:sheetView workbookViewId="0"/></x:sheetViews>'
            .'<x:sheetData>'
            .'<x:row r="1"><x:c r="A1" s="1" t="s"><x:v>0</x:v></x:c></x:row>'
            .'<x:row r="2"><x:c r="A2" s="2"/><x:c r="B2" s="3"/></x:row>'
            .'</x:sheetData>'
            .'<x:pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            .'</x:worksheet>';
    }

    // ─── tests ───────────────────────────────────────────────────────────

    public function test_parses_our_own_writer_output(): void
    {
        $path = sys_get_temp_dir().'/kxs-tpl-own-'.uniqid('', true).'.xlsx';
        $this->tmp[] = $path;
        $w = new SinkableXlsxWriter(new FileSink($path));
        $w->setHeaderStyle(['bold' => true, 'fill' => '#1F4E78', 'color' => '#FFFFFF']);
        $w->setColumnFormat(3, 'currency_try');
        $w->startFile(['id', 'name', 'amount']);
        $w->writeRow([1, 'sample', 10.5]);   // becomes the sample row
        $w->finishFile();

        $sheet = Template::open($path)->sheet('Report', dataStartRow: 2);

        $this->assertSame(1, $sheet->variantCount(), 'one sample row → one variant');
        $this->assertStringEndsWith('<sheetData>', $sheet->head());
        $this->assertStringStartsWith('</sheetData>', $sheet->tail());
        // Header row 1 is kept verbatim, the sample row is not emitted.
        $this->assertStringContainsString('r="1"', $sheet->headerRowsXml());
        $this->assertStringNotContainsString('r="2"', $sheet->headerRowsXml());
        // The formatted column carries a style id the writer chose.
        $this->assertArrayHasKey(2, $sheet->styleMap(0));
    }

    public function test_parses_phpspreadsheet_shaped_sheet(): void
    {
        $bytes = $this->buildXlsx($this->scaffold(['Leave' => $this->phpSpreadsheetSheet()]));
        $sheet = Template::fromString($bytes)->sheet('Leave', dataStartRow: 3);

        // head keeps layout, ends at the seam.
        $head = $sheet->head();
        $this->assertStringContainsString('<sheetViews>', $head);
        $this->assertStringContainsString('state="frozen"', $head);
        $this->assertStringContainsString('<cols>', $head);
        $this->assertStringEndsWith('<sheetData>', $head);

        // Two header rows kept, two sample rows consumed.
        $this->assertStringContainsString('r="1"', $sheet->headerRowsXml());
        $this->assertStringContainsString('r="2"', $sheet->headerRowsXml());
        $this->assertStringNotContainsString('r="3"', $sheet->headerRowsXml());
        $this->assertSame(2, $sheet->variantCount());

        // Style oracle: 0-based column => cellXfs id, gaps allowed (C absent).
        $this->assertSame([0 => 5, 1 => 6, 3 => 9], $sheet->styleMap(0));
        $this->assertSame([0 => 7, 1 => 8, 3 => 10], $sheet->styleMap(1));

        // tail byte-identical from the closing tag on.
        $this->assertStringStartsWith('</sheetData>', $sheet->tail());
        $this->assertStringContainsString('<mergeCell ref="A1:D1"/>', $sheet->tail());
        $this->assertStringContainsString('<pageMargins', $sheet->tail());
        $this->assertStringEndsWith('</worksheet>', $sheet->tail());
    }

    public function test_custom_height_before_ht_does_not_fool_the_row_height_parser(): void
    {
        $bytes = $this->buildXlsx($this->scaffold(['Leave' => $this->phpSpreadsheetSheet()]));
        $sheet = Template::fromString($bytes)->sheet('Leave', dataStartRow: 3);

        $attrs = $sheet->rowAttributes(0);
        // Naive /ht="([^"]*)"/ would read "1" out of customHeight="1".
        $this->assertSame('18.75', $attrs['ht']);
        $this->assertTrue($attrs['customHeight']);
    }

    public function test_namespace_prefixed_excel_dialect(): void
    {
        $bytes = $this->buildXlsx($this->scaffold(['Data' => $this->excelPrefixedSheet()]));
        $sheet = Template::fromString($bytes)->sheet('Data', dataStartRow: 2);

        $this->assertStringEndsWith('<x:sheetData>', $sheet->head());
        $this->assertStringStartsWith('</x:sheetData>', $sheet->tail());
        $this->assertSame(1, $sheet->variantCount());
        $this->assertSame([0 => 2, 1 => 3], $sheet->styleMap(0));
    }

    public function test_self_closing_sheet_data(): void
    {
        $body = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<dimension ref="A1:A1"/><sheetData/>'
            .'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            .'</worksheet>';
        $bytes = $this->buildXlsx($this->scaffold(['Empty' => $body]));
        $sheet = Template::fromString($bytes)->sheet('Empty', dataStartRow: 1);

        // The self-closing tag is normalised into an open/close pair.
        $this->assertStringEndsWith('<sheetData>', $sheet->head());
        $this->assertStringStartsWith('</sheetData>', $sheet->tail());
        $this->assertStringContainsString('<pageMargins', $sheet->tail());
        $this->assertSame('', $sheet->headerRowsXml());
        $this->assertSame(1, $sheet->variantCount(), 'no samples → one default variant');
        $this->assertSame([], $sheet->styleMap(0));
    }

    public function test_no_sample_rows_yields_one_unstyled_variant(): void
    {
        // dataStartRow past the last row: everything is a header row.
        $bytes = $this->buildXlsx($this->scaffold(['Leave' => $this->phpSpreadsheetSheet()]));
        $sheet = Template::fromString($bytes)->sheet('Leave', dataStartRow: 99);

        $this->assertSame(1, $sheet->variantCount());
        $this->assertSame([], $sheet->styleMap(0));
        $this->assertStringContainsString('r="4"', $sheet->headerRowsXml());
    }

    public function test_rows_without_r_attribute_use_positional_index(): void
    {
        $body = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
            .'<row><c s="1" t="inlineStr"><is><t>h</t></is></c></row>'
            .'<row><c s="4"/><c s="5"/></row>'
            .'</sheetData></worksheet>';
        $bytes = $this->buildXlsx($this->scaffold(['Compact' => $body]));
        $sheet = Template::fromString($bytes)->sheet('Compact', dataStartRow: 2);

        $this->assertSame(1, $sheet->variantCount());
        // No r on the cells either → positional columns.
        $this->assertSame([0 => 4, 1 => 5], $sheet->styleMap(0));
    }

    public function test_dimension_is_stripped_from_head(): void
    {
        // <dimension> cannot be rewritten in a streaming write (the last row
        // is unknown when head goes out) and is optional in the schema —
        // exactly what our own preamble does. Strip it.
        $bytes = $this->buildXlsx($this->scaffold(['Leave' => $this->phpSpreadsheetSheet()]));
        $sheet = Template::fromString($bytes)->sheet('Leave', dataStartRow: 3);

        $this->assertStringNotContainsString('<dimension', $sheet->head());
        $this->assertStringContainsString('<sheetPr/>', $sheet->head(), 'only dimension is removed');
    }

    public function test_merge_inside_the_data_region_is_rejected(): void
    {
        $body = str_replace(
            '<mergeCell ref="A1:D1"/>',
            '<mergeCell ref="A3:B4"/>',
            $this->phpSpreadsheetSheet()
        );
        $bytes = $this->buildXlsx($this->scaffold(['Leave' => $body]));

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/merge/i');
        Template::fromString($bytes)->sheet('Leave', dataStartRow: 3);
    }

    public function test_unknown_sheet_name_throws(): void
    {
        $bytes = $this->buildXlsx($this->scaffold(['Leave' => $this->phpSpreadsheetSheet()]));
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/Nope/');
        Template::fromString($bytes)->sheet('Nope', dataStartRow: 2);
    }

    public function test_sheet_names_and_entries_are_listed(): void
    {
        $bytes = $this->buildXlsx($this->scaffold([
            'Leave' => $this->phpSpreadsheetSheet(),
            'Data' => $this->excelPrefixedSheet(),
        ]));
        $t = Template::fromString($bytes);

        $this->assertSame(['Leave', 'Data'], $t->sheetNames());
        $this->assertSame('xl/worksheets/sheet2.xml', $t->entryFor('Data'));
        // Every zip entry is visible so finishFile() can copy the untouched ones.
        $this->assertContains('xl/workbook.xml', $t->entryNames());
        $this->assertContains('[Content_Types].xml', $t->entryNames());
    }

    public function test_template_is_parsed_once_and_reusable(): void
    {
        $bytes = $this->buildXlsx($this->scaffold(['Leave' => $this->phpSpreadsheetSheet()]));
        $t = Template::fromString($bytes);

        $a = $t->sheet('Leave', dataStartRow: 3);
        $b = $t->sheet('Leave', dataStartRow: 3);
        $this->assertSame($a->head(), $b->head());
        $this->assertSame($a->styleMap(1), $b->styleMap(1));
        // Different cut point on the same template object is allowed.
        $c = $t->sheet('Leave', dataStartRow: 4);
        $this->assertSame(1, $c->variantCount());
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Templates;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Templates\Template;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Template mode, second and third seam points: rows streamed into the cut,
 * everything else carried across untouched.
 *
 * The gates here are the ones that decide whether the feature is a fence or
 * a stretch. Rows must land at the template's own row numbers wearing the
 * sample row's style ids; every entry the writer did not author must arrive
 * CRC-identical, proven through this package's own ZipDirectory rather than
 * ext-zip so copy fidelity is attested by the reader that will consume it;
 * and every classic operation that would author layout the template already
 * owns must fail loudly instead of half-applying.
 */
class TemplateWriteTest extends TestCase
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

    // ─── fixtures ────────────────────────────────────────────────────────

    private function tmpPath(string $suffix = '.xlsx'): string
    {
        $path = sys_get_temp_dir().'/kxs-tw-'.uniqid('', true).$suffix;
        $this->tmp[] = $path;

        return $path;
    }

    /**
     * @param  array<string,string>  $parts
     * @param  list<string>  $stored  entries to write without compression
     */
    private function buildXlsx(array $parts, array $stored = []): string
    {
        $path = $this->tmpPath();
        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($path, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true);
        foreach ($parts as $name => $body) {
            $zip->addFromString($name, $body);
            if (in_array($name, $stored, true)) {
                $zip->setCompressionName($name, \ZipArchive::CM_STORE);
            }
        }
        $zip->close();

        return $path;
    }

    private function stylesXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<numFmts count="1"><numFmt numFmtId="180" formatCode="dd/mm/yyyy"/></numFmts>'
            .'<fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts>'
            .'<fills count="2"><fill><patternFill patternType="none"/></fill>'
            .'<fill><patternFill patternType="gray125"/></fill></fills>'
            .'<borders count="1"><border/></borders>'
            .'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            .'<cellXfs count="8">'
            .'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'      // 0 default
            .'<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>'      // 1 bold
            .'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'      // 2
            .'<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0"/>'      // 3 header
            .'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'      // 4 body text
            .'<xf numFmtId="180" fontId="0" fillId="0" borderId="0" xfId="0"/>'    // 5 body date
            .'<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0"/>'      // 6 body money
            .'<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0"/>'      // 7 zebra text
            .'</cellXfs>'
            .'<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'
            .'</styleSheet>';
    }

    private function sharedStringsXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">'
            .'<si><t>Personnel</t></si><si><t>Start</t></si><si><t>Amount</t></si>'
            .'</sst>';
    }

    /**
     * Two header rows, then two sample rows: variant 0 plain, variant 1 zebra.
     * Column A text, column B date, column C money. The tail carries a merge
     * that stays above the data region plus page setup.
     */
    private function mainSheet(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<dimension ref="A1:C5"/>'
            .'<sheetViews><sheetView tabSelected="1" workbookViewId="0">'
            .'<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            .'<sheetFormatPr defaultRowHeight="14.4"/>'
            .'<cols><col min="1" max="1" width="28.5" customWidth="1"/></cols>'
            .'<sheetData>'
            .'<row r="1" ht="30" customHeight="1"><c r="A1" s="3" t="s"><v>0</v></c></row>'
            .'<row r="2"><c r="A2" s="1" t="s"><v>0</v></c><c r="B2" s="1" t="s"><v>1</v></c>'
            .'<c r="C2" s="1" t="s"><v>2</v></c></row>'
            .'<row r="3" customHeight="1" ht="18.75"><c r="A3" s="4"/><c r="B3" s="5"/><c r="C3" s="6"/></row>'
            .'<row r="4" customHeight="1" ht="18.75" s="7" customFormat="1"><c r="A4" s="7"/><c r="B4" s="5"/>'
            .'<c r="C4" s="6"/></row>'
            .'</sheetData>'
            .'<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>'
            .'<conditionalFormatting sqref="A1:C2"><cfRule type="expression" dxfId="0" priority="1">'
            .'<formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>'
            .'<autoFilter ref="A2:C2"/>'
            .'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            .'</worksheet>';
    }

    private function summarySheet(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<sheetData><row r="1"><c r="A1" s="1" t="s"><v>2</v></c></row></sheetData>'
            .'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            .'</worksheet>';
    }

    /**
     * @param  array<string,string>  $sheets  sheet name => sheet XML
     * @param  array<string,string>  $extra   additional parts
     */
    private function templateFile(array $sheets, array $extra = [], array $stored = []): string
    {
        $sheetTags = '';
        $rels = '';
        $overrides = '';
        $i = 0;
        $parts = [];
        foreach ($sheets as $name => $body) {
            $i++;
            $sheetTags .= '<sheet name="'.$name.'" sheetId="'.$i.'" r:id="rId'.$i.'"/>';
            $rels .= '<Relationship Id="rId'.$i.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'.$i.'.xml"/>';
            $overrides .= '<Override PartName="/xl/worksheets/sheet'.$i.'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
            $parts['xl/worksheets/sheet'.$i.'.xml'] = $body;
        }

        $head = [
            '[Content_Types].xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                .'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                .'<Default Extension="xml" ContentType="application/xml"/>'
                .'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                .'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
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
            'xl/styles.xml' => $this->stylesXml(),
        ];

        return $this->buildXlsx($head + $parts + $extra, $stored);
    }

    /** The standard fixture: two sheets, a shared string table, a theme blob. */
    private function standardTemplate(): string
    {
        return $this->templateFile(
            ['Leaves' => $this->mainSheet(), 'Summary' => $this->summarySheet()],
            [
                'xl/sharedStrings.xml' => $this->sharedStringsXml(),
                'xl/theme/theme1.xml' => '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements/></a:theme>',
            ],
            ['xl/theme/theme1.xml']
        );
    }

    private function readOut(string $path, string $entry): string
    {
        $source = new LocalFileSource($path);
        $zip = ZipDirectory::fromSource($source);

        return $zip->readEntry($source, $entry);
    }

    private function sheetBody(string $xml): string
    {
        $start = strpos($xml, '<sheetData>');
        $end = strpos($xml, '</sheetData>');
        $this->assertNotFalse($start);
        $this->assertNotFalse($end);

        return substr($xml, $start + 11, $end - $start - 11);
    }

    // ─── the seam ────────────────────────────────────────────────────────

    public function test_rows_land_at_the_templates_row_numbers_wearing_its_styles(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', 42.5]);
        $writer->writeRow(['Grace', 17.0], variant: 1);
        $result = $writer->finishFile();

        $xml = $this->readOut($out, 'xl/worksheets/sheet1.xml');
        $body = $this->sheetBody($xml);

        // Header rows kept verbatim, sample rows dropped.
        $this->assertStringContainsString('<row r="1" ht="30" customHeight="1">', $body);
        $this->assertStringContainsString('<row r="2">', $body);
        $this->assertStringNotContainsString('r="A3" s="4"/>', $body);

        // Data begins at dataStartRow and wears the variant's style ids.
        $this->assertStringContainsString('r="A3"', $body);
        $this->assertStringContainsString('s="4"', $body, 'variant 0 column A style');
        $this->assertStringContainsString('<c r="B3" s="5" t="n"><v>42.5</v></c>', $body);
        $this->assertStringContainsString('<c r="A4" s="7"', $body, 'variant 1 column A style');
        $this->assertStringNotContainsString('r="A5"', $body, 'only two rows written');

        $this->assertSame(2, $result['rows']);
    }

    public function test_the_head_and_the_tail_survive_byte_for_byte(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', 1.0]);
        $writer->finishFile();

        $xml = $this->readOut($out, 'xl/worksheets/sheet1.xml');

        $this->assertStringContainsString('<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>', $xml);
        $this->assertStringContainsString('<cols><col min="1" max="1" width="28.5" customWidth="1"/></cols>', $xml);
        $this->assertStringContainsString(
            '</sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>',
            $xml
        );
        $this->assertStringContainsString(
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>',
            $xml
        );
        // The streaming writer cannot honour a stale range, so it is dropped.
        $this->assertStringNotContainsString('<dimension', $xml);
    }

    public function test_untouched_entries_are_copied_crc_identical(): void
    {
        $tpl = $this->standardTemplate();
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $tpl);
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', 1.0]);
        $writer->finishFile();

        // Fidelity attested through this package's own reader, not ext-zip.
        $inSrc = new LocalFileSource($tpl);
        $in = ZipDirectory::fromSource($inSrc);
        $outSrc = new LocalFileSource($out);
        $result = ZipDirectory::fromSource($outSrc);

        $untouched = [
            '[Content_Types].xml',
            '_rels/.rels',
            'xl/workbook.xml',
            'xl/_rels/workbook.xml.rels',
            'xl/worksheets/sheet2.xml',
            'xl/theme/theme1.xml',
        ];
        foreach ($untouched as $name) {
            $this->assertTrue($result->has($name), "{$name} is present in the output");
            $this->assertSame(
                $in->entry($name)['crc32'],
                $result->entry($name)['crc32'],
                "{$name} kept its CRC"
            );
            $this->assertSame(
                $in->entry($name)['compressed_size'],
                $result->entry($name)['compressed_size'],
                "{$name} was moved, not recompressed"
            );
            $this->assertSame(
                $in->readEntry($inSrc, $name),
                $result->readEntry($outSrc, $name),
                "{$name} still inflates to the same bytes"
            );
        }

        // The STORED entry kept its method — a copy, not a re-encode.
        $this->assertSame(0, $result->entry('xl/theme/theme1.xml')['method']);
        $this->assertSame(8, $result->entry('xl/workbook.xml')['method']);
    }

    public function test_styles_are_copied_verbatim_when_nothing_was_registered(): void
    {
        $tpl = $this->standardTemplate();
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $tpl);
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', 1.0]);
        $writer->finishFile();

        $this->assertSame($this->stylesXml(), $this->readOut($out, 'xl/styles.xml'));
    }

    public function test_a_second_sheet_can_be_streamed_after_the_first(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', 1.0]);
        $writer->sheet('Summary', 2);
        $writer->writeRow(['Total']);
        $result = $writer->finishFile();

        $this->assertSame(2, $result['rows']);
        $this->assertStringContainsString('r="A3"', $this->readOut($out, 'xl/worksheets/sheet1.xml'));
        $this->assertStringContainsString('r="A2"', $this->readOut($out, 'xl/worksheets/sheet2.xml'));
    }

    // ─── shared strings ──────────────────────────────────────────────────

    public function test_strings_join_the_templates_shared_table(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', 1.0]);
        $writer->writeRow(['Ada', 2.0]);
        $writer->finishFile();

        $body = $this->sheetBody($this->readOut($out, 'xl/worksheets/sheet1.xml'));
        $this->assertStringContainsString('<c r="A3" s="4" t="s"><v>3</v></c>', $body);
        $this->assertStringContainsString('<c r="A4" s="4" t="s"><v>3</v></c>', $body, 'repeat interns once');

        $sst = $this->readOut($out, 'xl/sharedStrings.xml');
        $this->assertStringContainsString('<si><t>Personnel</t></si>', $sst, 'seed kept');
        $this->assertStringContainsString('<si><t>Ada</t></si>', $sst);
        $this->assertStringContainsString('uniqueCount="4"', $sst);
    }

    public function test_strings_stay_inline_when_the_template_has_no_shared_table(): void
    {
        $tpl = $this->templateFile(['Leaves' => $this->mainSheet()]);
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $tpl);
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', 1.0]);
        $writer->finishFile();

        $body = $this->sheetBody($this->readOut($out, 'xl/worksheets/sheet1.xml'));
        $this->assertStringContainsString('t="inlineStr"><is><t>Ada</t></is>', $body);

        $result = ZipDirectory::fromSource(new LocalFileSource($out));
        $this->assertFalse(
            $result->has('xl/sharedStrings.xml'),
            'adding the part would need a new relationship and a new content type — a fifth seam'
        );
    }

    public function test_the_shared_table_is_copied_verbatim_when_no_string_was_added(): void
    {
        $tpl = $this->standardTemplate();
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $tpl);
        $writer->sheet('Leaves', 3);
        $writer->writeRow([null, 1.0]);
        $writer->finishFile();

        $inSrc = new LocalFileSource($tpl);
        $in = ZipDirectory::fromSource($inSrc);
        $outSrc = new LocalFileSource($out);
        $result = ZipDirectory::fromSource($outSrc);

        $this->assertSame(
            $in->entry('xl/sharedStrings.xml')['crc32'],
            $result->entry('xl/sharedStrings.xml')['crc32']
        );
    }

    // ─── type rules ──────────────────────────────────────────────────────

    public function test_dates_take_the_templates_format_not_the_classic_datetime_style(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', new \DateTimeImmutable('2026-03-01 00:00:00 UTC')]);
        $writer->finishFile();

        $body = $this->sheetBody($this->readOut($out, 'xl/worksheets/sheet1.xml'));
        // Column B's sample cell carries xf 5 (numFmtId 180, dd/mm/yyyy). The
        // classic STYLE_DATETIME constant is xf 1 here — the template's bold
        // text style — so applying it would silently mis-format the column.
        $this->assertStringContainsString('<c r="B3" s="5" t="n">', $body);
        $this->assertStringNotContainsString('<c r="B3" s="1"', $body);
        // And the template's style table was not touched.
        $this->assertSame($this->stylesXml(), $this->readOut($out, 'xl/styles.xml'));
    }

    public function test_a_date_beyond_the_samples_width_gets_an_appended_datetime_format(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        // Column D is past the sample row's last styled cell (C), so the
        // template holds no opinion and the classic date rule applies.
        $writer->writeRow(['Ada', 1.0, 2.0, new \DateTimeImmutable('2026-03-01 00:00:00 UTC')]);
        $writer->finishFile();

        $body = $this->sheetBody($this->readOut($out, 'xl/worksheets/sheet1.xml'));
        $this->assertMatchesRegularExpression('/<c r="D3" s="(\d+)" t="n">/', $body);
        preg_match('/<c r="D3" s="(\d+)" t="n">/', $body, $m);
        $this->assertGreaterThanOrEqual(8, (int) $m[1], 'appended past the seed table');

        $styles = $this->readOut($out, 'xl/styles.xml');
        $this->assertStringContainsString('<dxfs count="1">', $styles, 'unmodelled blocks survive');
        $this->assertStringContainsString('numFmtId="180"', $styles, 'the seed numFmt survives');
    }

    public function test_data_types_mirror_the_classic_writer(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->templateFile(['S' => $this->mainSheet()]));
        $writer->sheet('S', 3);
        $writer->writeRow([null, 7, 3.25, true, false, '', '00123', 'x y']);
        $writer->finishFile();

        $body = $this->sheetBody($this->readOut($out, 'xl/worksheets/sheet1.xml'));
        $this->assertStringContainsString('<c r="A3" s="4"/>', $body, 'null → empty cell, style kept');
        $this->assertStringContainsString('<c r="B3" s="5" t="n"><v>7</v></c>', $body);
        $this->assertStringContainsString('<c r="C3" s="6" t="n"><v>3.25</v></c>', $body);
        $this->assertStringContainsString('<c r="D3" t="b"><v>1</v></c>', $body);
        $this->assertStringContainsString('<c r="E3" t="b"><v>0</v></c>', $body);
        $this->assertStringContainsString('<c r="F3"/>', $body, 'empty string → empty cell');
        $this->assertStringContainsString('<t>00123</t>', $body, 'leading zero preserved as text');
    }

    // ─── the fence ───────────────────────────────────────────────────────

    public function test_an_unknown_sheet_name_throws(): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/no sheet named .Nope./');
        $writer->sheet('Nope', 2);
    }

    public function test_streaming_the_same_sheet_twice_throws(): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['Ada', 1.0]);
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/already been streamed/');
        $writer->sheet('Leaves', 3);
    }

    public function test_a_prefix_only_dialect_is_refused(): void
    {
        $sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<x:sheetData><x:row r="1"><x:c r="A1" s="1"/></x:row></x:sheetData></x:worksheet>';
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->templateFile(['P' => $sheet]));
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/default namespace/');
        $writer->sheet('P', 2);
    }

    public function test_compact_mode_and_template_mode_are_mutually_exclusive(): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/compact\(\)/');
        $writer->compact();
    }

    public function test_a_template_cannot_be_attached_to_a_compact_writer(): void
    {
        $writer = SinkableXlsxWriter::createForFile($this->tmpPath());
        $writer->compact();
        $this->expectException(XlsxStreamException::class);
        $writer->useTemplate(Template::open($this->standardTemplate()));
    }

    public function test_startFile_and_newSheet_are_refused_in_template_mode(): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        try {
            $writer->startFile(['a']);
            $this->fail('startFile() should be refused');
        } catch (XlsxStreamException $e) {
            $this->assertStringContainsString('startFile()', $e->getMessage());
        }

        $writer->sheet('Leaves', 3);
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/newSheet\(\)/');
        $writer->newSheet('Other');
    }

    #[\PHPUnit\Framework\Attributes\DataProvider('layoutOperations')]
    public function test_layout_operations_the_template_owns_are_refused(callable $op, string $needle): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/'.preg_quote($needle, '/').'/');
        $op($writer);
    }

    public static function layoutOperations(): array
    {
        return [
            'column format' => [fn ($w) => $w->setColumnFormat(2, 'date'), 'setColumnFormat()'],
            'column widths' => [fn ($w) => $w->setColumnWidths([1 => 20]), 'setColumnWidths()'],
            'auto width' => [fn ($w) => $w->setAutoColumnWidth(), 'setAutoColumnWidth()'],
            'freeze' => [fn ($w) => $w->freezeFirstRow(), 'freezeFirstRow()'],
            'freeze rc' => [fn ($w) => $w->freezeRowsAndColumns(1, 1), 'freezeRowsAndColumns()'],
            'auto filter' => [fn ($w) => $w->enableAutoFilter(), 'enableAutoFilter()'],
            'header style' => [fn ($w) => $w->setHeaderStyle(['bold' => true]), 'setHeaderStyle()'],
        ];
    }

    public function test_a_registered_row_style_is_refused_because_the_variant_owns_the_row(): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/variant/');
        $writer->writeRow(['Ada', 1.0], 5);
    }

    public function test_writing_before_a_sheet_is_chosen_throws(): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/sheet\(\)/');
        $writer->writeRow(['Ada']);
    }

    public function test_finishing_without_a_sheet_throws(): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $this->expectException(XlsxStreamException::class);
        $writer->finishFile();
    }

    public function test_sheet_is_refused_on_a_classic_writer(): void
    {
        $writer = SinkableXlsxWriter::createForFile($this->tmpPath());
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/useTemplate\(\)/');
        $writer->sheet('Leaves', 2);
    }

    // ─── bulk API ────────────────────────────────────────────────────────

    public function test_writeRows_can_pick_the_variant_per_row(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $writer->writeRows(
            [['a', 1.0], ['b', 2.0], ['c', 3.0], ['d', 4.0]],
            fn (array $row, int $i): int => $i % 2
        );
        $writer->finishFile();

        $body = $this->sheetBody($this->readOut($out, 'xl/worksheets/sheet1.xml'));
        $this->assertStringContainsString('<c r="A3" s="4"', $body, 'row 0 → variant 0');
        $this->assertStringContainsString('<c r="A4" s="7"', $body, 'row 1 → variant 1');
        $this->assertStringContainsString('<c r="A5" s="4"', $body);
        $this->assertStringContainsString('<c r="A6" s="7"', $body);
        // Variant 1's row-level attributes come from its sample row.
        $this->assertStringContainsString('<row r="4" ht="18.75" customHeight="1" s="7" customFormat="1">', $body);
        $this->assertStringContainsString('<row r="3" ht="18.75" customHeight="1">', $body);
    }

    public function test_an_out_of_range_variant_falls_back_to_the_first(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['a', 1.0], variant: 9);
        $writer->finishFile();

        $body = $this->sheetBody($this->readOut($out, 'xl/worksheets/sheet1.xml'));
        $this->assertStringContainsString('<c r="A3" s="4"', $body);
    }

    /**
     * A range the template drew over its header block is copied verbatim and
     * keeps meaning exactly what it meant. Ranges that reach into the data
     * region are refused at parse time instead (see TemplateParseTest), so
     * no output can carry a filter or format that silently stops short.
     */
    public function test_header_scoped_tail_ranges_are_carried_over_verbatim(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        for ($i = 0; $i < 20; $i++) {
            $writer->writeRow(['row '.$i, (float) $i]);
        }
        $writer->finishFile();

        $xml = $this->readOut($out, 'xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<autoFilter ref="A2:C2"/>', $xml);
        $this->assertStringContainsString('<conditionalFormatting sqref="A1:C2">', $xml);
        $this->assertStringContainsString('r="A22"', $xml, 'the data itself reaches row 22');
    }

    public function test_the_random_access_index_is_refused_rather_than_dropped(): void
    {
        $writer = SinkableXlsxWriter::createForFile($this->tmpPath());
        $writer->withRandomAccessIndex(100);
        try {
            $writer->useTemplate(Template::open($this->standardTemplate()));
            $this->fail('the sidecar and template mode should not combine yet');
        } catch (XlsxStreamException $e) {
            $this->assertStringContainsString('withRandomAccessIndex()', $e->getMessage());
        }

        // And turned on between attaching the template and choosing a sheet.
        $late = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $late->withRandomAccessIndex(100);
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/withRandomAccessIndex\(\)/');
        $late->sheet('Leaves', 3);
    }

    public function test_registering_a_row_style_is_refused(): void
    {
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($this->tmpPath()), $this->standardTemplate());
        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/registerRowStyle\(\)/');
        $writer->registerRowStyle(['bold' => true]);
    }

    public function test_a_full_dictionary_falls_back_to_inline_strings(): void
    {
        $out = $this->tmpPath();
        // The seed already holds three entries, so the fourth string saturates.
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate(), 4);
        $writer->sheet('Leaves', 3);
        $writer->writeRow(['first', 1.0]);
        $writer->writeRow(['second', 2.0]);
        $writer->writeRow(['first', 3.0]);
        $writer->finishFile();

        $body = $this->sheetBody($this->readOut($out, 'xl/worksheets/sheet1.xml'));
        $this->assertStringContainsString('<c r="A3" s="4" t="s"><v>3</v></c>', $body, 'the fourth entry still fits');
        $this->assertStringContainsString('t="inlineStr"><is><t>second</t></is>', $body, 'past the ceiling, inline');
        $this->assertStringContainsString('<c r="A5" s="4" t="s"><v>3</v></c>', $body, 'an interned repeat still resolves');

        $sst = $this->readOut($out, 'xl/sharedStrings.xml');
        $this->assertStringContainsString('uniqueCount="4"', $sst);
        $this->assertStringNotContainsString('second', $sst);
        $this->assertInstanceOf(\SimpleXMLElement::class, simplexml_load_string($sst));
    }

    public function test_a_caller_supplied_template_survives_the_writer_and_can_be_reused(): void
    {
        $template = Template::open($this->standardTemplate());

        foreach (['Ada', 'Grace'] as $name) {
            $out = $this->tmpPath();
            $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $template);
            $writer->sheet('Leaves', 3);
            $writer->writeRow([$name, 1.0]);
            $writer->finishFile();

            $this->assertStringContainsString($name, $this->readOut($out, 'xl/sharedStrings.xml'));
        }

        $template->close();
    }

    public function test_the_output_is_well_formed_xml(): void
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->standardTemplate());
        $writer->sheet('Leaves', 3);
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow(['name '.$i, (float) $i], variant: $i % 2);
        }
        $writer->finishFile();

        foreach (['xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/workbook.xml'] as $entry) {
            $this->assertInstanceOf(
                \SimpleXMLElement::class,
                simplexml_load_string($this->readOut($out, $entry)),
                "{$entry} is well-formed"
            );
        }
    }
}

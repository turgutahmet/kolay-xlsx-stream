<?php

namespace Kolay\XlsxStream\Tests\Templates;

use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Templates\Template;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\RandomAccessIndex;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Template mode and the random-access sidecar together.
 *
 * The sidecar is the fourth seam point: one Override in the template's
 * [Content_Types].xml, without which Excel drops into repair mode. The
 * harder gate is soundness — the template's header rows are real rows in
 * the streamed sheet's first index block, so a zone map built from the
 * streamed data alone could prune a block that an un-indexed scan would
 * have matched. The two paths must agree row for row.
 *
 * Only streamed sheets get an index section; the sheets carried across
 * untouched are not described, which the reader handles by falling back to
 * a scan.
 */
class TemplateIndexTest extends TestCase
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

    private function tmpPath(): string
    {
        $path = sys_get_temp_dir().'/kxs-ti-'.uniqid('', true).'.xlsx';
        $this->tmp[] = $path;

        return $path;
    }

    /**
     * A layout with two header rows: a title, then a column-name row whose
     * "2026" cell is numeric-looking on purpose — that is the cell an
     * unfolded zone map would let pruning hide.
     */
    private function templateFile(array $extra = [], string $contentTypesExtra = ''): string
    {
        $sheet = static function (string $tail): string {
            return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                .'<sheetData>'
                .'<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Yearly report</t></is></c><c r="B1" s="1"/></row>'
                .'<row r="2"><c r="A2" s="1" t="inlineStr"><is><t>Personnel</t></is></c>'
                .'<c r="B2" s="1" t="n"><v>2026</v></c></row>'
                .'<row r="3"><c r="A3" s="2"/><c r="B3" s="3"/></row>'
                .'</sheetData>'.$tail.'</worksheet>';
        };

        $parts = [
            '[Content_Types].xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                .'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                .'<Default Extension="xml" ContentType="application/xml"/>'
                .'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                .'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                .'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                .'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                .$contentTypesExtra.'</Types>',
            '_rels/.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                .'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                .'</Relationships>',
            'xl/workbook.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                .'<sheets><sheet name="Notes" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>',
            'xl/_rels/workbook.xml.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                .'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                .'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
                .'</Relationships>',
            'xl/styles.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                .'<fonts count="1"><font><sz val="11"/></font></fonts>'
                .'<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
                .'<borders count="1"><border/></borders>'
                .'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                .'<cellXfs count="4">'
                .'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                .'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                .'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                .'<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                .'</cellXfs></styleSheet>',
            // Sheet 1 is the untouched one; the data is streamed into sheet 2,
            // so a sidecar keyed by sheet index instead of entry would file
            // everything under a sheet the reader never opens.
            'xl/worksheets/sheet1.xml' => $sheet(''),
            'xl/worksheets/sheet2.xml' => $sheet(''),
        ];

        $path = $this->tmpPath();
        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($path, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true);
        foreach ($parts + $extra as $name => $body) {
            $zip->addFromString($name, $body);
        }
        $zip->close();

        return $path;
    }

    /** Stream $rows rows into the "Data" sheet, index on unless told otherwise. */
    private function write(string $template, int $rows, bool $index = true, int $every = 50): string
    {
        $out = $this->tmpPath();
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $template);
        if ($index) {
            $writer->withRandomAccessIndex($every)->withColumnStats([2])->withStringStats([1]);
        }
        $writer->setBufferFlushInterval(25);
        $writer->sheet('Data', 3);
        for ($i = 1; $i <= $rows; $i++) {
            $writer->writeRow(['Employee '.$i, (float) ($i * 10)]);
        }
        $writer->finishFile();

        return $out;
    }

    // ─── the sidecar itself ──────────────────────────────────────────────

    public function test_the_sidecar_is_written_and_declared_in_the_content_types(): void
    {
        $out = $this->write($this->templateFile(), 200);

        $source = new LocalFileSource($out);
        $zip = ZipDirectory::fromSource($source);
        $this->assertTrue($zip->has(RandomAccessIndex::ENTRY_PATH));

        $types = $zip->readEntry($source, '[Content_Types].xml');
        $this->assertSame(
            1,
            substr_count($types, 'PartName="/'.RandomAccessIndex::ENTRY_PATH.'"'),
            'declared exactly once'
        );
        $this->assertStringContainsString('ContentType="application/octet-stream"/></Types>', $types);
        // The rest of the template's declarations are untouched.
        $this->assertStringContainsString('PartName="/xl/worksheets/sheet2.xml"', $types);
        $this->assertStringContainsString('<Default Extension="rels"', $types);
    }

    public function test_the_content_types_is_copied_untouched_when_the_index_is_off(): void
    {
        $template = $this->templateFile();
        $out = $this->write($template, 50, index: false);

        $inSource = new LocalFileSource($template);
        $in = ZipDirectory::fromSource($inSource);
        $outSource = new LocalFileSource($out);
        $result = ZipDirectory::fromSource($outSource);

        $this->assertFalse($result->has(RandomAccessIndex::ENTRY_PATH));
        $this->assertSame(
            $in->entry('[Content_Types].xml')['crc32'],
            $result->entry('[Content_Types].xml')['crc32']
        );
    }

    public function test_an_already_declared_content_type_is_not_duplicated(): void
    {
        $template = $this->templateFile(
            ['xl/_kxs/index.bin' => 'stale sidecar bytes from an earlier export'],
            '<Override PartName="/xl/_kxs/index.bin" ContentType="application/octet-stream"/>'
        );
        $out = $this->write($template, 120);

        $source = new LocalFileSource($out);
        $zip = ZipDirectory::fromSource($source);
        $types = $zip->readEntry($source, '[Content_Types].xml');
        $this->assertSame(1, substr_count($types, 'PartName="/'.RandomAccessIndex::ENTRY_PATH.'"'));

        // The stale payload was replaced, not carried across.
        $payload = $zip->readEntry($source, RandomAccessIndex::ENTRY_PATH);
        $this->assertStringNotContainsString('stale sidecar bytes', $payload);
        $this->assertSame(120, StreamingXlsxReader::fromFile($out)->onSheet('Data')->rowCount() - 2);
    }

    // ─── the streamed sheet is queryable ─────────────────────────────────

    public function test_random_access_reads_the_streamed_sheet(): void
    {
        $out = $this->write($this->templateFile(), 300);
        $reader = StreamingXlsxReader::fromFile($out)->onSheet('Data');

        // Physical sheet rows: 1-2 are the template's header block, data
        // starts at 3.
        $this->assertSame(302, $reader->rowCount());
        $this->assertSame(['Employee 1', '10'], array_values($reader->rowAt(3)));
        $this->assertSame(['Employee 250', '2500'], array_values($reader->rowAt(252)));
        $this->assertNull($reader->rowAt(303));

        $range = [];
        foreach ($reader->rowRange(100, 104) as $rowNumber => $row) {
            $range[$rowNumber] = $row[0];
        }
        $this->assertSame(
            [100 => 'Employee 98', 101 => 'Employee 99', 102 => 'Employee 100', 103 => 'Employee 101', 104 => 'Employee 102'],
            $range
        );
        $reader->close();
    }

    /**
     * The sidecar is keyed by zip entry, and a template sheet keeps the
     * entry the template gave it — here the data is streamed into the
     * workbook's SECOND sheet, so an index that derived the path from the
     * writer's own sheet counter would file every sync point and zone map
     * under sheet1.xml, a sheet the reader never opens for this query. The
     * file would still read correctly, by scanning, which is exactly why
     * this is asserted rather than left to be noticed as slowness.
     */
    public function test_sync_points_are_filed_under_the_templates_own_entry(): void
    {
        $out = $this->write($this->templateFile(), 400);

        $source = new LocalFileSource($out);
        $zip = ZipDirectory::fromSource($source);
        $index = ReaderIndex::decode($zip->readEntry($source, RandomAccessIndex::ENTRY_PATH));

        $this->assertNotEmpty($index->syncPoints('xl/worksheets/sheet2.xml'), 'the streamed sheet has seek points');
        $this->assertSame([], $index->syncPoints('xl/worksheets/sheet1.xml'), 'the copied sheet has none');
        $this->assertNotNull($index->sheetCrc32('xl/worksheets/sheet2.xml'));

        // And the reader actually seeks: no query on the streamed sheet
        // falls back to a full scan.
        $reasons = [];
        $reader = StreamingXlsxReader::fromFile($out)->onSheet('Data');
        $reader->onFullScan(function (...$args) use (&$reasons) { $reasons[] = $args; });
        $reader->rowAt(300);
        iterator_to_array($reader->rowsWhere(2, '=', 1500));
        $reader->close();
        $this->assertSame([], $reasons, 'every read used the index');
    }

    public function test_the_workbook_verifies(): void
    {
        $out = $this->write($this->templateFile(), 400);
        $reader = StreamingXlsxReader::fromFile($out);
        $result = $reader->verify();

        $this->assertTrue($result['ok'], 'every sheet verifies, streamed and copied alike');
        $reader->close();
    }

    public function test_profile_reads_the_streamed_sheet(): void
    {
        $out = $this->write($this->templateFile(), 500);
        $reader = StreamingXlsxReader::fromFile($out)->onSheet('Data');
        $profile = $reader->profile([2], histogram: false);

        $this->assertArrayHasKey(2, $profile['columns']);
        $this->assertSame(5000.0, $profile['columns'][2]['max']);
        $reader->close();
    }

    // ─── soundness ───────────────────────────────────────────────────────

    /**
     * The gate this part exists for. The template's second header row holds
     * a numeric 2026 in the tracked column and the word "Personnel" in the
     * tracked string column, both far outside the streamed values' range. If
     * those cells were not folded into the first block's zone maps, pruning
     * would drop the block that holds them and the indexed answer would be
     * missing a row the plain scan returns.
     */
    public function test_zone_map_pruning_never_hides_a_template_header_row(): void
    {
        $indexed = $this->write($this->templateFile(), 400);
        $plain = $this->write($this->templateFile(), 400, index: false);

        foreach ([[2, '=', 2026], [2, '>=', 2026], [2, 'between', 2000]] as $predicate) {
            [$column, $op] = $predicate;
            $value = $predicate[2];

            $withIndex = $this->rowNumbers($indexed, $column, $op, $value, $op === 'between' ? 3000 : null);
            $withoutIndex = $this->rowNumbers($plain, $column, $op, $value, $op === 'between' ? 3000 : null);

            $this->assertSame($withoutIndex, $withIndex, "numeric {$op} agrees with the un-indexed scan");
            $this->assertContains(2, $withIndex, 'the header row itself is returned');
        }

        $withIndex = $this->rowNumbers($indexed, 1, '=', 'Personnel');
        $withoutIndex = $this->rowNumbers($plain, 1, '=', 'Personnel');
        $this->assertSame($withoutIndex, $withIndex, 'string equality agrees with the un-indexed scan');
        $this->assertSame([2], $withIndex);
    }

    /** @return list<int> */
    private function rowNumbers(string $file, int $column, string $op, mixed $value, mixed $value2 = null): array
    {
        $reader = StreamingXlsxReader::fromFile($file)->onSheet('Data');
        $rows = [];
        foreach ($reader->rowsWhere($column, $op, $value, $value2) as $rowNumber => $_) {
            $rows[] = $rowNumber;
        }
        $reader->close();

        return $rows;
    }

    // ─── what the sidecar does not cover ─────────────────────────────────

    public function test_only_the_streamed_sheet_is_described(): void
    {
        $out = $this->write($this->templateFile(), 200);
        $reader = StreamingXlsxReader::fromFile($out);

        // The copied sheet still reads correctly, by scanning rather than
        // seeking — it carries no section, which is the documented rule.
        $this->assertSame(3, $reader->onSheet('Notes')->rowCount());
        $this->assertSame(202, $reader->onSheet('Data')->rowCount());
        $reader->close();
    }
}

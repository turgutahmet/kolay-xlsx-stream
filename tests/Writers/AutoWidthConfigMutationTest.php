<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\BaseXlsxWriter as W;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * v3.2.2 manual-Excel-test findings, written RED against the pre-fix
 * writer:
 *
 * 1. Sample auto-width defers the active sheet's preamble, so config
 *    mutations issued while a sample is pending used to act
 *    RETROACTIVELY on the sheet being written: clearColumnFormats()
 *    before newSheet() wiped the previous sheet's explicit
 *    setColumnWidths() entry (observed 40 -> 131), and a header style
 *    set for the NEXT sheet repainted the PREVIOUS sheet's header.
 *    The fix pins the principle: config mutations never act
 *    retroactively — a pending sample is finalized (with the state it
 *    was sampled under) before the mutation lands.
 *
 * 2. minWidthForFormat() only knew NAMED format presets; columns
 *    formatted via BUILTIN_NUMFMT_* ids fell back to width 8 in
 *    heuristic mode and rendered as ###### in Excel. Builtin ids now
 *    map to the same minimum-width table — <cols> hints only, cell
 *    bytes are pinned unchanged.
 */
class AutoWidthConfigMutationTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-widthmut-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * The documented multi-sheet order — clear + configure the NEXT
     * sheet, then newSheet() — must not corrupt the sheet whose sample
     * is still pending: its explicit width survives, its header keeps
     * the style that was active while it was written, and the next
     * sheet gets the new style/formats.
     */
    public function test_pre_new_sheet_cleanup_does_not_corrupt_pending_sample_sheet(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(sample: 100); // stays pending past every row below
        $writer->setColumnWidths([2 => 10]);
        $writer->setHeaderStyle(['fill' => '#1F4E78', 'color' => '#FFFFFF']);
        $writer->setColumnFormat(3, W::BUILTIN_NUMFMT_CURRENCY);
        $writer->startFile(['id', 'desc', 'price']);

        for ($i = 1; $i <= 5; $i++) {
            $writer->writeRow([$i, 'a deliberately long description well past ten chars', 12.5 * $i]);
        }

        // Prepare sheet 2 (different layout) the documented way.
        $writer->clearColumnFormats();
        $writer->setColumnFormat(2, W::BUILTIN_NUMFMT_DATE);
        $writer->setHeaderStyle(['fill' => '#375623', 'color' => '#FFFFFF']);
        $writer->newSheet('S2', ['when', 'date']);
        $writer->writeRow(['x', new \DateTimeImmutable('2026-01-15')]);
        $writer->finishFile();

        // (ii) sheet 1 keeps the user width 10 on col 2 — pre-fix the
        // pending sample won with the long-description width.
        $this->assertStringContainsString(
            '<col min="2" max="2" width="10"',
            $this->sheetXml(1),
            'explicit setColumnWidths entry must survive a pre-newSheet clearColumnFormats'
        );

        // (iii) sheet 1's header keeps ITS style; sheet 2 gets the new
        // one — pre-fix both carried the second style id.
        $s1 = $this->headerStyleIdOf(1);
        $s2 = $this->headerStyleIdOf(2);
        $this->assertNotNull($s1);
        $this->assertNotNull($s2);
        $this->assertNotSame($s1, $s2, 'a header style set for the NEXT sheet must not repaint the pending sheet');

        // (iv) each sheet's own column format landed on its data cells.
        $this->assertMatchesRegularExpression('/<c r="C2" s="\d+" t="n">/', $this->sheetXml(1));
        $this->assertMatchesRegularExpression('/<c r="B2" s="\d+" t="n">/', $this->sheetXml(2));
    }

    /**
     * Changing the header style mid-sheet (while the sample is pending)
     * finalizes the pending sheet with the OLD style; the new style
     * only applies from the next sheet on. Rows written after the
     * early finalize keep flowing normally.
     */
    public function test_mid_sheet_header_style_change_does_not_rewrite_pending_header(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(sample: 100);
        $writer->setHeaderStyle(['fill' => '#1F4E78', 'color' => '#FFFFFF']);
        $writer->startFile(['a', 'b']);
        $writer->writeRow([1, 'x']);
        $writer->writeRow([2, 'y']);

        $writer->setHeaderStyle(['fill' => '#7030A0', 'color' => '#FFFFFF']); // mid-sheet

        $writer->writeRow([3, 'z']);
        $writer->writeRow([4, 'w']);
        $writer->newSheet('S2');
        $writer->writeRow([5, 'v']);
        $writer->finishFile();

        $this->assertNotSame(
            $this->headerStyleIdOf(1),
            $this->headerStyleIdOf(2),
            'mid-sheet setHeaderStyle must not repaint the sheet whose sample was pending'
        );

        // 4 data rows + header all landed on sheet 1.
        $this->assertSame(5, preg_match_all('/<row /', $this->sheetXml(1)));
    }

    /**
     * FIX B: builtin numFmt ids participate in the heuristic minimum-
     * width table — date/datetime/currency columns stop rendering as
     * ###### — while the cell bytes stay byte-identical to the pre-fix
     * writer (the literal below was captured from the pre-fix build;
     * only the <cols> hints may differ).
     */
    public function test_builtin_numfmt_ids_get_min_width_in_heuristic_mode(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth(); // heuristic
        $writer->setColumnFormat(1, W::BUILTIN_NUMFMT_DATE);
        $writer->setColumnFormat(2, W::BUILTIN_NUMFMT_DATETIME);
        $writer->setColumnFormat(3, W::BUILTIN_NUMFMT_CURRENCY);
        $writer->startFile(['d', 'dt', 'c']);
        $writer->writeRow([
            new \DateTimeImmutable('2026-01-15 10:30:00'),
            new \DateTimeImmutable('2026-01-15 10:30:00'),
            1234.56,
        ]);
        $writer->finishFile();

        $xml = $this->sheetXml(1);

        $this->assertStringContainsString('<col min="1" max="1" width="12"', $xml); // date
        $this->assertStringContainsString('<col min="2" max="2" width="20"', $xml); // datetime
        $this->assertStringContainsString('<col min="3" max="3" width="14"', $xml); // currency

        // Cell bytes pinned to the pre-fix writer output.
        $this->assertStringContainsString(
            '<row r="2"><c r="A2" s="2" t="n"><v>46037.4375</v></c>'
            .'<c r="B2" s="3" t="n"><v>46037.4375</v></c>'
            .'<c r="C2" s="4" t="n"><v>1234.56</v></c></row>',
            $xml
        );
    }

    /**
     * Named presets keep their existing minimums (regression guard for
     * the shared table) and unformatted columns stay on the header
     * heuristic.
     */
    public function test_named_presets_and_plain_columns_unchanged(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setAutoColumnWidth();
        $writer->setColumnFormat(1, 'date');
        $writer->startFile(['d', 'plain']);
        $writer->writeRow([new \DateTimeImmutable('2026-01-15'), 'x']);
        $writer->finishFile();

        $xml = $this->sheetXml(1);
        $this->assertStringContainsString('<col min="1" max="1" width="12"', $xml);
        $this->assertStringContainsString('<col min="2" max="2" width="8"', $xml);
    }

    private function sheetXml(int $sheet): string
    {
        $zip = new \ZipArchive();
        $zip->open($this->testFile);
        $xml = (string) $zip->getFromName("xl/worksheets/sheet{$sheet}.xml");
        $zip->close();

        return $xml;
    }

    /** s= attribute of the sheet's A1 header cell, or null. */
    private function headerStyleIdOf(int $sheet): ?string
    {
        if (preg_match('/<c r="A1" s="(\d+)"/', $this->sheetXml($sheet), $m)) {
            return $m[1];
        }

        return null;
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\BaseXlsxWriter;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * v3.3 compact mode (H1): opt-in `->compact()` emits rows and cells
 * WITHOUT the optional r attributes (ECMA-376: c/@r and row/@r are
 * optional; readers assign positions sequentially).
 *
 * The load-bearing proof is the TRANSFORM BYTE-ORACLE: for the same
 * input, the compact writer's sheet XML must equal the classic
 * writer's sheet XML with every ` r="A1"`-style attribute stripped —
 * i.e. compact is the exact r-less projection of the classic path,
 * not a reimplementation that could drift. Every cell-shape branch
 * (strings incl. ws-preserve/numeric-preserve, int/float with and
 * without column styles, null/'' placeholders, bool, DateTime,
 * Stringable, styled rows incl. merged column-format styles, the
 * preamble header row, sample-mode deferred preambles, multi-sheet)
 * rides through the corpus below.
 */
class CompactModeTest extends TestCase
{
    /** @var list<string> */
    private array $files = [];

    protected function tearDown(): void
    {
        foreach ($this->files as $f) {
            @unlink($f);
        }
        $this->files = [];
        parent::tearDown();
    }

    private function tmpFile(string $tag): string
    {
        $path = sys_get_temp_dir()."/kxs-compact-{$tag}-".uniqid('', true).'.xlsx';
        $this->files[] = $path;

        return $path;
    }

    /**
     * Rich corpus exercising every emission branch, written with the
     * exact same calls in both modes.
     */
    private function writeCorpus(string $path, bool $compact): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($path));
        if ($compact) {
            $writer->compact();
        }

        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([1, 3]);
        $writer->withColumnSketches([1, 3]);
        $writer->setAutoColumnWidth(50);           // sample mode -> deferred preamble path
        $writer->setColumnWidths([6 => 30]);
        $writer->freezeFirstRow();
        $writer->enableAutoFilter();
        $writer->setHeaderStyle(['fill' => '#1F4E78', 'color' => '#FFFFFF', 'bold' => true]);
        $writer->setColumnFormat(3, BaseXlsxWriter::BUILTIN_NUMFMT_CURRENCY);
        $writer->setColumnFormat(4, 'dd"."mm"."yyyy hh:mm', raw: true);

        $writer->startFile(['id', 'ad', 'tutar', 'tarih', 'aktif', 'not']);

        $red = $writer->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
        $blue = $writer->registerRowStyle(['fill' => '#1F4E78', 'color' => '#FFFFFF', 'bold' => true]);

        $base = new \DateTimeImmutable('2026-03-01 09:00:00');
        $stringable = new class () {
            public function __toString(): string
            {
                return '  stringable ws  ';
            }
        };

        for ($i = 1; $i <= 600; $i++) {
            $row = [
                $i,
                'müşteri <'.$i.'> & "özel"',
                round($i * 1.37, 2),
                $base->modify('+'.$i.' minutes'),
                $i % 2 === 0,
                'not-'.$i,
            ];
            if ($i % 7 === 0) {          // sparse: mid-row nulls
                $row[1] = null;
                $row[3] = null;
            }
            if ($i % 11 === 0) {         // empty strings
                $row[2] = '';
                $row[4] = '';
            }
            if ($i % 13 === 0) {         // numeric-string preservation + big int
                $row[1] = '007';
                $row[5] = '12345678901234567890';
            }
            if ($i % 17 === 0) {         // ws-preserve + stringable
                $row[1] = '  boşluklu  ';
                $row[5] = $stringable;
            }
            if ($i % 19 === 0) {         // short row (trailing columns absent)
                $row = [$i, 'kısa'];
            }

            $style = $i % 5 === 0 ? ($i % 10 === 0 ? $blue : $red) : null;
            $writer->writeRow($row, $style);
        }

        // Second sheet: different schema, clean formats (v3.2.2 rule:
        // prepare BEFORE newSheet is safe — pending sample finalizes).
        $writer->clearColumnFormats();
        $writer->setHeaderStyle(['fill' => '#375623', 'color' => '#FFFFFF']);
        $writer->newSheet('Özet', ['ay', 'toplam']);
        for ($m = 1; $m <= 12; $m++) {
            $writer->writeRow([$m, $m * 1000.5]);
        }

        $writer->finishFile();
    }

    /** The PoC transform: strip cell refs, then row refs. */
    private function stripRefs(string $sheetXml): string
    {
        $out = preg_replace('~ r="[A-Z]+\d+"~', '', $sheetXml);

        return preg_replace('~<row r="\d+"~', '<row', $out);
    }

    private function sheetXml(string $path, int $sheet): string
    {
        $zip = new ZipArchive();
        $this->assertTrue($zip->open($path));
        $xml = $zip->getFromName("xl/worksheets/sheet{$sheet}.xml");
        $zip->close();
        $this->assertNotFalse($xml, "sheet{$sheet}.xml missing in {$path}");

        return $xml;
    }

    public function test_compact_output_is_the_exact_r_less_projection_of_classic_output(): void
    {
        $classic = $this->tmpFile('classic');
        $compact = $this->tmpFile('compact');
        $this->writeCorpus($classic, compact: false);
        $this->writeCorpus($compact, compact: true);

        foreach ([1, 2] as $sheet) {
            $expected = $this->stripRefs($this->sheetXml($classic, $sheet));
            $actual = $this->sheetXml($compact, $sheet);
            $this->assertSame($expected, $actual, "sheet{$sheet}: compact XML is not the exact r-less projection");
        }
    }

    public function test_reader_reads_compact_file_identically_to_classic(): void
    {
        $classic = $this->tmpFile('classic-r');
        $compact = $this->tmpFile('compact-r');
        $this->writeCorpus($classic, compact: false);
        $this->writeCorpus($compact, compact: true);

        $rA = StreamingXlsxReader::fromFile($classic);
        $rB = StreamingXlsxReader::fromFile($compact);
        foreach ([0, 1] as $idx) {
            $rA->onSheetIndex($idx);
            $rB->onSheetIndex($idx);
            $a = iterator_to_array($rA->rows(), false);
            $b = iterator_to_array($rB->rows(), false);
            $this->assertSame($a, $b, "sheet index {$idx} rows differ between classic and compact");
            $this->assertGreaterThan(2, count($a));
        }
    }

    public function test_sidecar_random_access_works_on_compact_files(): void
    {
        $path = $this->tmpFile('sidecar');
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->compact();
        $writer->withRandomAccessIndex(every: 500);
        $writer->withColumnStats([1, 2]);
        $writer->withColumnSketches([1, 2]);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 5000; $i++) {
            $writer->writeRow([$i, 2 * $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($path);
        $this->assertSame(5001, $reader->rowCount());

        $row = $reader->rowAt(4001);            // sync-point seek into an r-less block
        $this->assertNotNull($row);
        $this->assertSame('4000', (string) $row[0]);

        $stats = $reader->columnStats(1);
        $this->assertNotNull($stats);
        $this->assertSame(5000, $stats['count']);
        $this->assertSame(5000.0, $stats['max']);
        $this->assertSame('asc', $stats['sorted']);

        $this->assertSame(5000.0, $reader->quantile(1, 1.0)); // exact max from TDIG

        $hits = iterator_to_array($reader->rowsWhere(1, 'between', 4000, 4002), true);
        $this->assertSame([4001, 4002, 4003], array_keys($hits));
    }

    public function test_compact_after_start_file_throws(): void
    {
        $path = $this->tmpFile('late');
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->startFile(['a']);

        $this->expectException(XlsxStreamException::class);
        $writer->compact();
    }

    public function test_sparse_rows_keep_positions_in_compact_mode(): void
    {
        $path = $this->tmpFile('sparse');
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->compact();
        $writer->startFile(['a', 'b', 'c', 'd']);
        $writer->writeRow([1, null, 3, null]);   // mid + trailing null
        $writer->writeRow([null, '', null, 4]);  // leading null + empty string
        $writer->writeRow([5]);                  // short row
        $writer->finishFile();

        $rows = iterator_to_array(StreamingXlsxReader::fromFile($path)->rows(skip: 1), false);
        $this->assertSame(['1', '', '3', ''], $rows[0]);
        $this->assertSame(['', '', '', '4'], $rows[1]);
        $this->assertSame(['5'], $rows[2]);
    }

    public function test_compact_with_sample_auto_width_keeps_user_widths(): void
    {
        $path = $this->tmpFile('widths');
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->compact();
        $writer->setAutoColumnWidth(10);
        $writer->setColumnWidths([2 => 33]);
        $writer->startFile(['x', 'y']);
        for ($i = 0; $i < 20; $i++) {
            $writer->writeRow(['aa', str_repeat('uzun', 10)]);
        }
        $writer->finishFile();

        $xml = $this->sheetXml($path, 1);
        $this->assertStringContainsString('<col min="2" max="2" width="33"', $xml);
        $this->assertStringNotContainsString(' r="', preg_replace('~<worksheet[^>]*>~', '', $xml));
    }
}

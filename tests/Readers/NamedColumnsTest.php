<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * v3.3 E1 — named column addressing on the reader's query surface.
 *
 * Written RED-FIRST against the int-only API. The contract under test:
 *
 *   - Every query API (columnStats, quantile/median, countDistinct,
 *     rowsWhere, findRow, groupStats) and castColumn/castColumns accept
 *     a header NAME wherever they accepted a column number, and the
 *     named call returns EXACTLY what the equivalent numeric call
 *     returns — names are sugar, never a second code path.
 *   - A string is ALWAYS a name — '2024' looks up a header called
 *     "2024", it is never coerced to column 2024 (ambiguity ban).
 *   - Duplicate header names throw (naming both positions); unknown
 *     names throw listing the available headers (pit-of-success).
 *   - The 1-based (query APIs) vs 0-based (castColumn) split becomes
 *     invisible when addressing by name — the exact trap E1 removes.
 */
class NamedColumnsTest extends TestCase
{
    private string $file;

    protected function setUp(): void
    {
        parent::setUp();
        $this->file = sys_get_temp_dir().'/kxs-named-'.uniqid('', true).'.xlsx';

        $writer = new SinkableXlsxWriter(new FileSink($this->file));
        $writer->withRandomAccessIndex(every: 500);
        $writer->withColumnStats([1, 2]);
        $writer->withColumnSketches([1, 2]);
        $writer->startFile(['ID', 'Amount', 'City', '2024', '', 'Tarih']);

        for ($i = 1; $i <= 2000; $i++) {
            $writer->writeRow([$i, $i * 2.5, 'c'.($i % 7), $i % 100, 'x'.$i, 45000 + $i]);
        }
        $writer->finishFile();
    }

    protected function tearDown(): void
    {
        @unlink($this->file);
        parent::tearDown();
    }

    private function reader(): StreamingXlsxReader
    {
        return StreamingXlsxReader::fromFile($this->file);
    }

    public function test_query_apis_accept_names_and_match_numeric_calls(): void
    {
        $byInt = $this->reader();
        $byName = $this->reader();

        $this->assertSame($byInt->columnStats(2), $byName->columnStats('Amount'));
        $this->assertSame($byInt->quantile(2, 0.9), $byName->quantile('Amount', 0.9));
        $this->assertSame($byInt->median(1), $byName->median('ID'));
        $this->assertSame($byInt->countDistinct(3), $byName->countDistinct('City'));
        $this->assertSame($byInt->findRow(1, 1500), $byName->findRow('ID', 1500));

        $intRows = iterator_to_array($byInt->rowsWhere(2, 'between', 100.0, 110.0));
        $nameRows = iterator_to_array($byName->rowsWhere('Amount', 'between', 100.0, 110.0));
        $this->assertNotEmpty($intRows);
        $this->assertSame($intRows, $nameRows);

        $bucket = fn (float $v) => intdiv((int) $v, 500);
        $this->assertSame(
            $byInt->groupStats(1, 2, $bucket),
            $byName->groupStats('ID', 'Amount', $bucket)
        );
    }

    public function test_cast_column_by_name_matches_zero_based_numeric(): void
    {
        $byInt = $this->reader()->castColumn(1, 'float');   // 0-based -> Amount
        $byName = $this->reader()->castColumn('Amount', 'float');

        $a = iterator_to_array($byInt->rows(skip: 1, limit: 3), false);
        $b = iterator_to_array($byName->rows(skip: 1, limit: 3), false);

        $this->assertSame($a, $b);
        $this->assertSame(2.5, $a[0][1]);
    }

    public function test_cast_columns_accepts_name_keys(): void
    {
        $reader = $this->reader()->castColumns(['Amount' => 'float', 'ID' => 'int']);

        $row = iterator_to_array($reader->rows(skip: 1, limit: 1), false)[0];
        $this->assertSame(1, $row[0]);
        $this->assertSame(2.5, $row[1]);
    }

    public function test_numeric_looking_string_is_a_name_not_an_index(): void
    {
        // '2024' is the 4th header — NOT column 2024.
        $stats = $this->reader()->rowsWhere('2024', '=', 42.0);
        $first = null;
        foreach ($stats as $rn => $row) {
            $first = [$rn, $row];
            break;
        }

        $this->assertNotNull($first);
        $this->assertSame('42', (string) $first[1][3]);
    }

    public function test_unknown_name_throws_listing_headers(): void
    {
        try {
            $this->reader()->columnStats('Amonut');
            $this->fail('expected unknown-column exception');
        } catch (XlsxReadException $e) {
            $this->assertStringContainsString('Amonut', $e->getMessage());
            $this->assertStringContainsString('Amount', $e->getMessage()); // available list
            $this->assertStringContainsString('City', $e->getMessage());
        }
    }

    public function test_duplicate_header_name_throws_with_both_positions(): void
    {
        $dup = sys_get_temp_dir().'/kxs-named-dup-'.uniqid('', true).'.xlsx';
        $writer = new SinkableXlsxWriter(new FileSink($dup));
        $writer->startFile(['ID', 'Amount', 'Amount']);
        $writer->writeRow([1, 10, 20]);
        $writer->finishFile();

        try {
            $reader = StreamingXlsxReader::fromFile($dup);
            try {
                $reader->countDistinct('Amount');
                $this->fail('expected ambiguous-column exception');
            } catch (XlsxReadException $e) {
                $this->assertStringContainsString('Amount', $e->getMessage());
                $this->assertStringContainsString('2', $e->getMessage());
                $this->assertStringContainsString('3', $e->getMessage());
            }
            // Unambiguous names on the same sheet keep working.
            $this->assertNull($reader->columnStats('ID'));
        } finally {
            @unlink($dup);
        }
    }

    public function test_empty_header_cell_is_not_addressable(): void
    {
        // Column 5's header is '' — no name can reach it, but the
        // sheet's other names still resolve (the empty cell is skipped,
        // not an error).
        $this->assertNotNull($this->reader()->columnStats('Amount'));

        $this->expectException(XlsxReadException::class);
        $this->reader()->columnStats('');
    }

    public function test_named_cast_and_auto_detect_dates_interact_like_numeric(): void
    {
        // castColumn by name must land in the same skip-set slot the
        // numeric form uses, so autoDetectDates() exempts the column
        // identically on both paths.
        $byInt = $this->reader()->autoDetectDates()->castColumn(5, fn ($v) => 'RAW:'.$v);
        $byName = $this->reader()->autoDetectDates()->castColumn('Tarih', fn ($v) => 'RAW:'.$v);

        $a = iterator_to_array($byInt->rows(skip: 1, limit: 2), false);
        $b = iterator_to_array($byName->rows(skip: 1, limit: 2), false);

        $this->assertSame($a, $b);
        $this->assertSame('RAW:45001', $a[0][5]);
    }

    public function test_names_follow_the_active_sheet(): void
    {
        $multi = sys_get_temp_dir().'/kxs-named-multi-'.uniqid('', true).'.xlsx';
        $writer = new SinkableXlsxWriter(new FileSink($multi));
        $writer->withRandomAccessIndex();
        $writer->withColumnStats([1]);
        $writer->startFile(['ID', 'Amount']);
        for ($i = 1; $i <= 10; $i++) {
            $writer->writeRow([$i, $i * 2]);
        }
        $writer->newSheet('Totals', ['Region', 'Sum']);
        for ($i = 1; $i <= 5; $i++) {
            $writer->writeRow([$i * 100, $i]);
        }
        $writer->finishFile();

        try {
            $reader = StreamingXlsxReader::fromFile($multi);
            $this->assertSame(10.0, $reader->columnStats('ID')['max']);

            $reader->onSheet('Totals');
            $this->assertSame(500.0, $reader->columnStats('Region')['max']);

            // Sheet 1's names no longer resolve on Totals.
            $this->expectException(XlsxReadException::class);
            $reader->columnStats('Amount');
        } finally {
            @unlink($multi);
        }
    }
}

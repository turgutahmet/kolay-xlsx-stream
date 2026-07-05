<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * v3.2.2 correctness fix (BRAINSTORM F2): on auto-split workbooks the
 * whole query surface silently answered from the ACTIVE sheet only —
 * a 2.1M-row export reported rowCount 1,048,575, findRow() returned
 * null for rows past the first sheet, columnStats() summed less than
 * half the data. These tests were written RED against v3.2.1 behaviour
 * and pin the spanning semantics:
 *
 *   - A continuation CHAIN is a maximal run of consecutive sheets where
 *     every non-final member is exactly full (1,048,575 rows including
 *     its header) and every member carries an identical header row.
 *     Only born-indexed files are chain-detected (the sidecar provides
 *     the O(1) row totals the detection needs).
 *   - Chain = ONE logical table. Global row numbers count sheet 1 in
 *     full and continuation members' data rows only (their repeated
 *     header rows do not exist logically): data row i sits at global
 *     row i + 1.
 *   - Explicitly different tables (newSheet with different headers, or
 *     non-full sheets) are never chained — per-sheet semantics stay.
 *
 * The fixture is written once per class (2 full sheets + a 5,000-row
 * tail, ~2.1M rows) with closed-form cell values so every oracle below
 * is arithmetic, not a re-scan: id column = data row index, amount =
 * 2*id, city = 'c' . (id % 5).
 */
class AutoSplitSpanTest extends TestCase
{
    private const FULL_SHEET_ROWS = 1_048_575;      // rows per full sheet incl. header

    private const DATA_PER_FULL_SHEET = self::FULL_SHEET_ROWS - 1;

    private const TAIL_DATA_ROWS = 5_000;

    private const DATA_ROWS = 2 * self::DATA_PER_FULL_SHEET + self::TAIL_DATA_ROWS;

    private const GLOBAL_ROWS = self::DATA_ROWS + 1; // one logical header

    private static ?string $file = null;

    public static function setUpBeforeClass(): void
    {
        parent::setUpBeforeClass();

        $path = sys_get_temp_dir().'/kxs-autosplit-span-'.uniqid('', true).'.xlsx';
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->withRandomAccessIndex(every: 10_000);
        $writer->withColumnStats([1, 2]);
        $writer->withColumnSketches([1, 2]);
        $writer->startFile(['id', 'amount', 'city', 'flag']);

        for ($i = 1; $i <= self::DATA_ROWS; $i++) {
            $writer->writeRow([$i, 2 * $i, 'c'.($i % 5), $i % 2 === 0]);
        }
        $writer->finishFile();

        self::$file = $path;
    }

    public static function tearDownAfterClass(): void
    {
        if (self::$file !== null) {
            @unlink(self::$file);
            self::$file = null;
        }
        parent::tearDownAfterClass();
    }

    private function reader(): StreamingXlsxReader
    {
        return StreamingXlsxReader::fromFile(self::$file);
    }

    /** The writer really did split: 3 sheets, first two exactly full. */
    public function test_fixture_shape(): void
    {
        $sheets = $this->reader()->sheets();

        $this->assertCount(3, $sheets);
    }

    public function test_row_count_spans_the_chain(): void
    {
        $this->assertSame(self::GLOBAL_ROWS, $this->reader()->rowCount());
    }

    public function test_column_stats_cover_every_data_row(): void
    {
        $stats = $this->reader()->columnStats(1);

        $this->assertNotNull($stats);
        $this->assertSame(self::DATA_ROWS, $stats['count']);
        $this->assertSame(1.0, $stats['min']);
        $this->assertSame((float) self::DATA_ROWS, $stats['max']);
        // Σ 1..D — exact in float64 well past 2^53 only in sum of ~2.2e12;
        // compare against the closed form computed the same way.
        $d = (float) self::DATA_ROWS;
        $this->assertSame($d * ($d + 1.0) / 2.0, $stats['sum']);
        // 'other' carries one text header cell per physical sheet.
        $this->assertSame(3, $stats['other']);
        $this->assertSame('asc', $stats['sorted']);
    }

    public function test_find_row_reaches_past_the_first_sheet(): void
    {
        $hit = $this->reader()->findRow(1, 2_000_000);

        $this->assertNotNull($hit);
        $this->assertSame(2_000_001, $hit['row']);
        $this->assertSame('2000000', (string) $hit['values'][0]);
        $this->assertSame('4000000', (string) $hit['values'][1]);
    }

    public function test_row_at_uses_global_numbering(): void
    {
        $row = $this->reader()->rowAt(2_000_001);

        $this->assertNotNull($row);
        $this->assertSame('2000000', (string) $row[0]);

        // Global row 1 stays the (single logical) header.
        $this->assertSame('id', $this->reader()->rowAt(1)[0]);

        // Past-the-end stays null.
        $this->assertNull($this->reader()->rowAt(self::GLOBAL_ROWS + 1));
    }

    public function test_row_range_crosses_the_sheet_boundary_seamlessly(): void
    {
        // Global rows 1,048,574..1,048,578 straddle sheet1 -> sheet2.
        // Data ids run 1,048,573..1,048,577 with no gap and no repeated
        // header row in between.
        $got = [];
        foreach ($this->reader()->rowRange(1_048_574, 1_048_578) as $rn => $row) {
            $got[$rn] = (int) $row[0];
        }

        $this->assertSame([
            1_048_574 => 1_048_573,
            1_048_575 => 1_048_574,
            1_048_576 => 1_048_575,
            1_048_577 => 1_048_576,
            1_048_578 => 1_048_577,
        ], $got);
    }

    public function test_rows_skip_lands_in_a_later_sheet(): void
    {
        $got = [];
        foreach ($this->reader()->rows(skip: self::GLOBAL_ROWS - 2) as $row) {
            $got[] = (int) $row[0];
        }

        $this->assertSame([self::DATA_ROWS - 1, self::DATA_ROWS], $got);
    }

    public function test_rows_where_matches_in_a_later_sheet_with_global_keys(): void
    {
        $got = [];
        foreach ($this->reader()->rowsWhere(1, 'between', 2_000_000, 2_000_004) as $rn => $row) {
            $got[$rn] = (int) $row[0];
        }

        $this->assertSame([
            2_000_001 => 2_000_000,
            2_000_002 => 2_000_001,
            2_000_003 => 2_000_002,
            2_000_004 => 2_000_003,
            2_000_005 => 2_000_004,
        ], $got);
    }

    public function test_group_stats_span_the_chain(): void
    {
        // bucket = intdiv(id, 500_000) — monotone over the sorted id
        // column; group g covers ids max(1, 500_000g)..min(D, 500_000g+499_999).
        $groups = $this->reader()->groupStats(1, 2, fn (float $v) => intdiv((int) $v, 500_000));

        $byKey = [];
        foreach ($groups as $g) {
            $byKey[(int) $g['group']] = $g;
        }

        // Ids reach 2,102,148 -> buckets 0..4 must ALL be present
        // (bucket 4 starts at id 2,000,000 — entirely inside sheet 3's
        // chain territory, invisible without spanning).
        $this->assertSame([0, 1, 2, 3, 4], array_keys($byKey));

        foreach ($byKey as $bucket => $g) {
            $lo = max(1, $bucket * 500_000);
            $hi = min(self::DATA_ROWS, $bucket * 500_000 + 499_999);
            $n = $hi - $lo + 1;
            $this->assertSame($n, $g['count'], "bucket {$bucket} count");
            // amount = 2*id -> sum = 2 * Σ lo..hi
            $expected = (float) ($lo + $hi) * $n; // 2 * (lo+hi)*n/2
            $this->assertSame($expected, $g['sum'], "bucket {$bucket} sum");
            $this->assertSame((float) (2 * $lo), $g['min']);
            $this->assertSame((float) (2 * $hi), $g['max']);
        }
    }

    public function test_sketches_cover_the_chain(): void
    {
        $reader = $this->reader();

        // q=1 is the digest's exact max — chain merge must surface the
        // global maximum, not sheet 1's.
        $this->assertSame((float) self::DATA_ROWS, $reader->quantile(1, 1.0));
        $this->assertSame(1.0, $reader->quantile(1, 0.0));

        // Median of 1..D ≈ D/2 within the documented rank error.
        $median = $reader->median(1);
        $this->assertNotNull($median);
        $this->assertEqualsWithDelta(self::DATA_ROWS / 2, $median, self::DATA_ROWS * 0.01);

        // Distinct ids ≈ D (±2.3% documented; ±5% CI bound). Sheet-1-only
        // behaviour (~1.05M) sits 50% off and can never pass this.
        $distinct = $reader->countDistinct(1);
        $this->assertNotNull($distinct);
        $this->assertEqualsWithDelta(self::DATA_ROWS, $distinct, self::DATA_ROWS * 0.05);
    }

    public function test_shards_cover_every_chain_sheet(): void
    {
        $shards = $this->reader()->shards(6);

        $sheetsCovered = array_unique(array_column($shards, 'sheet'));
        sort($sheetsCovered);

        $this->assertSame([
            'xl/worksheets/sheet1.xml',
            'xl/worksheets/sheet2.xml',
            'xl/worksheets/sheet3.xml',
        ], $sheetsCovered);

        // Shards must tile each sheet's rows exactly (local numbering,
        // headers ride at local row 1 of their first shard).
        $bySheet = [];
        foreach ($shards as $shard) {
            $bySheet[$shard['sheet']][] = $shard;
        }
        foreach ($bySheet as $sheet => $list) {
            $this->assertSame(1, $list[0]['first_row'], "{$sheet} first shard");
            for ($i = 1; $i < count($list); $i++) {
                $this->assertSame(
                    $list[$i - 1]['last_row'] + 1,
                    $list[$i]['first_row'],
                    "{$sheet} shard {$i} continuity"
                );
            }
        }
        $this->assertSame(self::FULL_SHEET_ROWS, end($bySheet['xl/worksheets/sheet1.xml'])['last_row']);
        $this->assertSame(self::TAIL_DATA_ROWS + 1, end($bySheet['xl/worksheets/sheet3.xml'])['last_row']);
    }

    public function test_selecting_a_chain_member_still_answers_for_the_whole_chain(): void
    {
        // A chain is one logical table; onSheet() into a member does not
        // shrink the query surface back to a physical sheet.
        $reader = $this->reader()->onSheetIndex(1);

        $this->assertSame(self::GLOBAL_ROWS, $reader->rowCount());
        $this->assertSame((float) self::DATA_ROWS, $reader->columnStats(1)['max']);
    }

    // ------------------------------------------------------------------
    // Negative controls: chains must NOT form here.
    // ------------------------------------------------------------------

    public function test_distinct_tables_are_not_chained(): void
    {
        $path = sys_get_temp_dir().'/kxs-span-neg1-'.uniqid('', true).'.xlsx';
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->withRandomAccessIndex();
        $writer->withColumnStats([1]);
        $writer->startFile(['id', 'name']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i, 'a'.$i]);
        }
        $writer->newSheet('Totals', ['region', 'sum']);
        for ($i = 1; $i <= 10; $i++) {
            $writer->writeRow([$i * 1000, 'r'.$i]);
        }
        $writer->finishFile();

        try {
            $reader = StreamingXlsxReader::fromFile($path);
            // Sheet 1 is not full -> no chain; per-sheet semantics intact.
            $this->assertSame(101, $reader->rowCount());
            $this->assertSame(100.0, $reader->columnStats(1)['max']);

            $reader->onSheet('Totals');
            $this->assertSame(11, $reader->rowCount());
            $this->assertSame(10000.0, $reader->columnStats(1)['max']);
        } finally {
            @unlink($path);
        }
    }

    public function test_same_header_partial_sheets_are_not_chained(): void
    {
        $path = sys_get_temp_dir().'/kxs-span-neg2-'.uniqid('', true).'.xlsx';
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->withRandomAccessIndex();
        $writer->withColumnStats([1]);
        $writer->startFile(['id', 'name']);
        for ($i = 1; $i <= 50; $i++) {
            $writer->writeRow([$i, 'a'.$i]);
        }
        // Same header layout, explicitly separate sheet — sheet 1 is not
        // full, so this is two tables that happen to share a schema.
        $writer->newSheet('Q2');
        for ($i = 51; $i <= 80; $i++) {
            $writer->writeRow([$i, 'a'.$i]);
        }
        $writer->finishFile();

        try {
            $reader = StreamingXlsxReader::fromFile($path);
            $this->assertSame(51, $reader->rowCount());
            $this->assertSame(50.0, $reader->columnStats(1)['max']);

            $reader->onSheet('Q2');
            $this->assertSame(31, $reader->rowCount());
            $this->assertSame(80.0, $reader->columnStats(1)['max']);
        } finally {
            @unlink($path);
        }
    }
}

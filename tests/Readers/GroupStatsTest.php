<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Oracle tests for groupStats(): every pruned (sorted-group pushdown)
 * result must equal a brute-force GROUP BY computed row-by-row over
 * rows() with the same semantics — header excluded, non-numeric group
 * cells excluded, booleans folded as 0/1, aggregates over numeric
 * aggregate cells only. The brute implementation below is deliberately
 * independent of the production code so a shared bug can't vouch for
 * itself.
 *
 * A counting-spy Source additionally pins the PLAN: group-pure interior
 * blocks must never be inflated (the whole point of the pushdown), which
 * identical results alone cannot prove.
 */
class GroupStatsTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-group-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * @param  list<array<int, mixed>>  $rows
     * @param  list<string>  $header
     */
    private function writeRows(array $rows, array $header = ['grp', 'amount'], bool $indexed = true): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        if ($indexed) {
            $writer->withRandomAccessIndex(every: 100);
            $writer->withColumnStats([1, 2]);
            $writer->setBufferFlushInterval(100);
        }
        $writer->startFile($header);
        foreach ($rows as $row) {
            $writer->writeRow($row);
        }
        $writer->finishFile();
    }

    /**
     * Independent brute-force GROUP BY over rows() — the oracle.
     *
     * @return list<array{group: int|float, sum: float, count: int, min: float|null, max: float|null}>
     */
    private function bruteGroups(StreamingXlsxReader $reader, int $groupBy, int $aggregate, ?callable $bucket = null): array
    {
        $bucket ??= static fn (float $v): float => $v;
        $num = static function (mixed $cell): ?float {
            if (is_int($cell) || is_float($cell)) {
                return (float) $cell;
            }
            if (is_string($cell)) {
                return $cell !== '' && is_numeric($cell) ? (float) $cell : null;
            }
            if (is_bool($cell)) {
                return $cell ? 1.0 : 0.0;
            }

            return null;
        };

        $groups = [];
        $isHeader = true;
        foreach ($reader->rows() as $row) {
            if ($isHeader) {
                $isHeader = false;

                continue;
            }
            $g = $num($row[$groupBy - 1] ?? null);
            if ($g === null) {
                continue;
            }
            $g = $bucket($g);
            $key = (string) $g;
            if (! isset($groups[$key])) {
                $groups[$key] = ['group' => $g, 'sum' => 0.0, 'count' => 0, 'min' => null, 'max' => null];
            }
            $a = $num($row[$aggregate - 1] ?? null);
            if ($a === null) {
                continue;
            }
            $groups[$key]['sum'] += $a;
            $groups[$key]['count']++;
            $groups[$key]['min'] = $groups[$key]['min'] === null ? $a : min($groups[$key]['min'], $a);
            $groups[$key]['max'] = $groups[$key]['max'] === null ? $a : max($groups[$key]['max'], $a);
        }

        return array_values($groups);
    }

    /**
     * Exact on group identity, order and count; delta on the float
     * aggregates (block sums and the brute sum add the same values in
     * different association order, which is not bit-identical for
     * non-representable decimals).
     *
     * @param  list<array{group: int|float, sum: float, count: int, min: float|null, max: float|null}>  $expected
     * @param  list<array{group: int|float, sum: float, count: int, min: float|null, max: float|null}>  $actual
     */
    private function assertGroupsMatch(array $expected, array $actual, string $label): void
    {
        $this->assertSame(
            array_map(fn ($g) => $g['group'], $expected),
            array_map(fn ($g) => $g['group'], $actual),
            "{$label}: group keys/order diverged"
        );
        foreach ($expected as $i => $exp) {
            $act = $actual[$i];
            $this->assertSame($exp['count'], $act['count'], "{$label}: count diverged for group {$exp['group']}");
            $this->assertEqualsWithDelta($exp['sum'], $act['sum'], 1e-6, "{$label}: sum diverged for group {$exp['group']}");
            foreach (['min', 'max'] as $k) {
                if ($exp[$k] === null) {
                    $this->assertNull($act[$k], "{$label}: {$k} not null for group {$exp['group']}");
                } else {
                    $this->assertNotNull($act[$k], "{$label}: {$k} null for group {$exp['group']}");
                    $this->assertEqualsWithDelta($exp[$k], $act[$k], 1e-9, "{$label}: {$k} diverged for group {$exp['group']}");
                }
            }
        }
    }

    private function assertOracleHolds(int $groupBy, int $aggregate, ?callable $bucket, string $label): void
    {
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $pruned = $reader->groupStats($groupBy, $aggregate, $bucket);
        $brute = $this->bruteGroups($reader, $groupBy, $aggregate, $bucket);
        $this->assertGroupsMatch($brute, $pruned, $label);
        $reader->close();
    }

    public function test_sorted_duplicate_groups_spanning_block_boundaries(): void
    {
        // 650 rows, groups of 130 — every group straddles a 100-row
        // block boundary, so the plan mixes interior blocks and
        // boundary scans. Decimal amounts exercise the float-delta
        // comparison honestly.
        $rows = [];
        for ($i = 1; $i <= 650; $i++) {
            $rows[] = [intdiv($i - 1, 130), $i * 1.25];
        }
        $this->writeRows($rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSame('asc', $reader->columnStats(1)['sorted']);
        $result = $reader->groupStats(1, 2);
        $this->assertCount(5, $result);
        $this->assertSame([0.0, 1.0, 2.0, 3.0, 4.0], array_map(fn ($g) => $g['group'], $result));
        $this->assertSame(130, $result[0]['count']);
        $reader->close();

        $this->assertOracleHolds(1, 2, null, 'duplicate groups');
    }

    public function test_group_column_with_text_and_empty_cells_mixed_in(): void
    {
        // Non-numeric group cells (text, empty, missing) are excluded
        // rows. They don't break the sorted flag (only numeric values
        // feed the order tracker) but they DO dirty their block's
        // `other` count, forcing a row-level scan there. Booleans in
        // the AGGREGATE column check the 0/1 fold on both paths.
        $rows = [];
        for ($i = 1; $i <= 500; $i++) {
            $grp = intdiv($i - 1, 100);
            if ($i % 37 === 0) {
                $grp = '';                      // empty cell -> excluded row
            } elseif ($i % 41 === 0) {
                $grp = 'n/a';                   // text cell  -> excluded row
            }
            $amount = $i % 29 === 0 ? ($i % 2 === 0) : $i * 2.5; // sprinkle bools
            $rows[] = [$grp, $amount];
        }
        $this->writeRows($rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSame('asc', $reader->columnStats(1)['sorted'], 'text cells must not break sortedness');
        $reader->close();

        $this->assertOracleHolds(1, 2, null, 'text/empty mixed');
    }

    public function test_single_group_sheet(): void
    {
        $rows = [];
        for ($i = 1; $i <= 500; $i++) {
            $rows[] = [7, $i];
        }
        $this->writeRows($rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $result = $reader->groupStats(1, 2);
        $this->assertCount(1, $result);
        $this->assertSame(7.0, $result[0]['group']);
        $this->assertSame(500, $result[0]['count']);
        $this->assertSame((float) (500 * 501 / 2), $result[0]['sum']);
        $this->assertSame(1.0, $result[0]['min']);
        $this->assertSame(500.0, $result[0]['max']);
        $reader->close();

        $this->assertOracleHolds(1, 2, null, 'single group');
    }

    public function test_every_row_distinct_groups(): void
    {
        // Identity bucket + strictly increasing values: every block
        // straddles a group boundary, so the pushdown degenerates to
        // scanning everything — same results, no wrong pruning.
        $rows = [];
        for ($i = 1; $i <= 300; $i++) {
            $rows[] = [$i, $i * 3.0];
        }
        $this->writeRows($rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $result = $reader->groupStats(1, 2);
        $this->assertCount(300, $result);
        $this->assertSame(1, $result[0]['count']);
        $reader->close();

        $this->assertOracleHolds(1, 2, null, 'distinct groups');
    }

    public function test_descending_sorted_group_column(): void
    {
        $rows = [];
        for ($i = 1; $i <= 650; $i++) {
            $rows[] = [1000 - intdiv($i - 1, 130), $i * 0.5];
        }
        $this->writeRows($rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSame('desc', $reader->columnStats(1)['sorted']);
        $result = $reader->groupStats(1, 2);
        // First-encounter sheet order = descending group values.
        $this->assertSame([1000.0, 999.0, 998.0, 997.0, 996.0], array_map(fn ($g) => $g['group'], $result));
        $reader->close();

        $this->assertOracleHolds(1, 2, null, 'descending sort');
    }

    public function test_bucket_closure_merges_values_into_coarser_groups(): void
    {
        // Monotone threshold bucket (floor of a division) — the shape
        // date-serial -> month bucketing takes. Group values 0.1 .. 60.0
        // bucket into floor(v/5): 13 groups with boundary rows landing
        // mid-block.
        $rows = [];
        for ($i = 1; $i <= 600; $i++) {
            $rows[] = [$i / 10, $i];
        }
        $this->writeRows($rows);

        $bucket = static fn (float $v): int => (int) floor($v / 5);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $result = $reader->groupStats(1, 2, $bucket);
        $this->assertSame(range(0, 12), array_map(fn ($g) => $g['group'], $result));
        $reader->close();

        $this->assertOracleHolds(1, 2, $bucket, 'bucket closure');
    }

    public function test_unsorted_group_column_falls_back_to_full_scan(): void
    {
        mt_srand(1234);
        $rows = [];
        for ($i = 1; $i <= 400; $i++) {
            // A handful of repeating groups in random order; bool group
            // cells fold as 0/1 on the scan path (group 1 merges them).
            $grp = $i % 53 === 0 ? true : mt_rand(0, 5);
            $rows[] = [$grp, $i * 1.5];
        }
        $this->writeRows($rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->columnStats(1)['sorted']);
        $reader->close();

        $this->assertOracleHolds(1, 2, null, 'unsorted fallback');
    }

    public function test_without_sidecar_falls_back_to_full_scan(): void
    {
        $rows = [];
        for ($i = 1; $i <= 200; $i++) {
            $rows[] = [intdiv($i - 1, 50), $i];
        }
        $this->writeRows($rows, indexed: false);

        $this->assertOracleHolds(1, 2, null, 'no sidecar');
    }

    public function test_group_whose_aggregate_cells_are_all_text_still_appears(): void
    {
        $rows = [];
        for ($i = 1; $i <= 300; $i++) {
            $grp = intdiv($i - 1, 100);
            $rows[] = [$grp, $grp === 1 ? 'pending' : $i * 1.0]; // group 1: no numeric aggregates
        }
        $this->writeRows($rows);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $result = $reader->groupStats(1, 2);
        $this->assertCount(3, $result);
        $this->assertSame(0, $result[1]['count']);
        $this->assertSame(0.0, $result[1]['sum']);
        $this->assertNull($result[1]['min']);
        $this->assertNull($result[1]['max']);
        $reader->close();

        $this->assertOracleHolds(1, 2, null, 'all-text aggregate group');
    }

    /**
     * Header caution (the v3.1 lesson): header cells fold into block
     * 0's stats. A TEXT header is invisible (non-numeric), and a
     * NUMERIC-LOOKING header must not become a group — block 0 always
     * takes the scan path, where row 1 is excluded by row number.
     */
    public function test_numeric_looking_header_does_not_become_a_group(): void
    {
        $rows = [];
        for ($i = 1; $i <= 300; $i++) {
            $rows[] = [intdiv($i - 1, 100) + 1, $i * 2.0];
        }
        // Header '999999' in the group column, '888888' in the aggregate.
        $this->writeRows($rows, header: ['999999', '888888']);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        // The numeric header must not poison sortedness (writer folds it
        // without order tracking) — pushdown stays active.
        $this->assertSame('asc', $reader->columnStats(1)['sorted']);

        $result = $reader->groupStats(1, 2);
        $this->assertSame([1.0, 2.0, 3.0], array_map(fn ($g) => $g['group'], $result));
        foreach ($result as $g) {
            $this->assertSame(100, $g['count'], 'header leaked into a group aggregate');
        }
        $reader->close();

        $this->assertOracleHolds(1, 2, null, 'numeric header');
    }

    public function test_rejects_non_positive_columns(): void
    {
        $this->writeRows([[1, 2]]);
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\InvalidArgumentException::class);
        $reader->groupStats(0, 2);
    }

    /**
     * The plan pin: interior (group-pure) blocks must never be
     * inflated. Fixture: 500 rows, group A for rows 1..250 and group B
     * for 251..500, sync every 100. Expected plan:
     *
     *   block 0 (rows 1..~101)     scan  (header rides here)
     *   block 1                    interior A — from block stats
     *   block 2 (A/B boundary)     scan
     *   blocks 3, 4                interior B — from block stats
     *   tail                       empty/skip
     *
     * = exactly TWO stream opens: one at the sheet start (block-0 run)
     * and one at a sync-point offset past it (boundary run). A
     * regression that inflates interior blocks shows up as extra opens;
     * results-only oracles can't see this.
     */
    public function test_interior_blocks_are_not_read(): void
    {
        $rows = [];
        for ($i = 1; $i <= 500; $i++) {
            $rows[] = [$i <= 250 ? 10 : 20, $i];
        }
        $this->writeRows($rows);

        $spy = new class(new LocalFileSource($this->testFile)) implements Source
        {
            /** @var list<int> */
            public array $streamOffsets = [];

            public function __construct(private LocalFileSource $inner) {}

            public function size(): int
            {
                return $this->inner->size();
            }

            public function range(int $offset, int $length): string
            {
                return $this->inner->range($offset, $length);
            }

            public function streamFrom(int $offset, ?int $length = null)
            {
                $this->streamOffsets[] = $offset;

                return $this->inner->streamFrom($offset, $length);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };

        $reader = StreamingXlsxReader::from($spy);

        // Where does a plain sequential read open the sheet stream?
        $reader->rows()->current();
        $sheetStart = min($spy->streamOffsets);
        $spy->streamOffsets = [];

        $result = $reader->groupStats(1, 2);

        $this->assertCount(2, $spy->streamOffsets, 'pushdown plan should open exactly two scan runs');
        $this->assertSame($sheetStart, min($spy->streamOffsets), 'block-0 run must start at the sheet start');
        $this->assertGreaterThan($sheetStart, max($spy->streamOffsets), 'boundary run must seek past the sheet start');

        // And the pruned result is still exact.
        $this->assertGroupsMatch($this->bruteGroups($reader, 1, 2), $result, 'spy fixture');
        $reader->close();
    }
}

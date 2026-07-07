<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * groupQuantile(groupBy, aggregate, q) — percentile GROUP BY, composed
 * from the row-space building blocks: a sorted group column yields each
 * group's contiguous sheet-row span (whole group-pure blocks from the
 * zone maps, boundary blocks scanned), and each group's quantile is then
 * answered by the range-quantile path over that span — large groups merge
 * TDGB superblocks, small groups scan. An unsorted or unindexed group
 * column degrades to one honest scan routing rows into per-group digests.
 * Gates: pushdown result == exact per-group oracle without a full scan,
 * unsorted fallback == same oracle with a full scan announced, bucketing,
 * empty-aggregate groups.
 */
class GroupQuantileTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-gq-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * Writes a two-column sheet: col 1 the group id, col 2 the amount.
     *
     * @param  callable(int): array{0: int|float, 1: mixed}  $row  data-row ordinal → [group, amount]
     * @return array<string, list<float>>  group key => amounts (numeric only)
     */
    private function writeFixture(int $rows, callable $row, bool $withTdgb = true): array
    {
        /** @var array<string, list<float>> $oracle */
        $oracle = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([1]);
        if ($withTdgb) {
            $writer->withRangeQuantiles([2]);
        }
        $writer->withRandomAccessIndex(every: 1_000)->setBufferFlushInterval(1_000);
        $writer->startFile(['group', 'amount']);
        for ($i = 1; $i <= $rows; $i++) {
            [$g, $amount] = $row($i);
            $writer->writeRow([$g, $amount]);
            if (is_int($amount) || is_float($amount)) {
                $oracle[(string) $g][] = (float) $amount;
            } else {
                $oracle[(string) $g] ??= [];
            }
        }
        $writer->finishFile();

        return $oracle;
    }

    /**
     * @param  array<string, list<float>>  $oracle
     * @param  list<array{group: int|float, count: int, quantile: float|null}>  $result
     */
    private function assertGroupsMatch(array $oracle, array $result, float $q, float $delta = 0.02): void
    {
        $this->assertCount(count($oracle), $result, 'group count mismatch');
        foreach ($result as $g) {
            $key = (string) $g['group'];
            $this->assertArrayHasKey($key, $oracle, "unexpected group {$key}");
            $vals = $oracle[$key];
            sort($vals);
            $n = count($vals);
            $this->assertSame($n, $g['count'], "count mismatch for group {$key}");
            if ($n === 0) {
                $this->assertNull($g['quantile'], "group {$key} has no values but a quantile");

                continue;
            }
            $below = 0;
            foreach ($vals as $v) {
                if ($v < $g['quantile']) {
                    $below++;
                }
            }
            $this->assertEqualsWithDelta($q, $below / $n, $delta, "rank error group {$key} q={$q}");
        }
    }

    public function test_sorted_groups_pushdown_matches_oracle_without_full_scan(): void
    {
        // Groups of mixed size: group 1 spans 25000 rows (> one 16384-row
        // superblock → merge path), group 2 spans 15000 rows (< superblock
        // → scan path). Both sorted ascending → contiguous.
        $oracle = $this->writeFixture(40_000, function (int $i): array {
            $g = $i <= 25_000 ? 1 : 2;

            return [$g, round((($i * 7919) % 10_000) + $i * 0.001, 3)];
        });

        $scans = [];
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $reader->onFullScan(function (array $ctx) use (&$scans): void {
            $scans[] = $ctx;
        });

        foreach ([0.5, 0.9] as $q) {
            $result = $reader->groupQuantile('group', 'amount', $q);
            $this->assertGroupsMatch($oracle, $result, $q);
        }

        $this->assertSame([], $scans, 'sorted+TDGB group quantile must not full-scan');
        $reader->close();
    }

    public function test_unsorted_group_column_falls_back_to_honest_scan(): void
    {
        // Non-monotone group ids → not contiguous → no pushdown basis.
        $oracle = $this->writeFixture(6_000, function (int $i): array {
            $g = ($i * 7) % 5; // 0..4, shuffled across the sheet

            return [$g, round((($i * 13) % 997) + 0.5, 3)];
        });

        $scans = [];
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $reader->onFullScan(function (array $ctx) use (&$scans): void {
            $scans[] = $ctx;
        });

        $result = $reader->groupQuantile('group', 'amount', 0.5);
        $this->assertGroupsMatch($oracle, $result, 0.5);

        $this->assertNotEmpty($scans, 'an unsorted group column must announce the honest scan');
        $this->assertSame('groupby-not-sorted', $scans[0]['reason']);
        $reader->close();
    }

    public function test_bucketed_groups(): void
    {
        // Sorted id column bucketed into bands of 10000 → groups 0,1,2,3.
        $oracle = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([1])->withRangeQuantiles([2]);
        $writer->withRandomAccessIndex(every: 1_000)->setBufferFlushInterval(1_000);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 35_000; $i++) {
            $amount = round((($i * 6151) % 8_000) + 0.25, 3);
            $writer->writeRow([$i, $amount]);
            $oracle[(string) intdiv($i, 10_000)][] = $amount;
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $result = $reader->groupQuantile('id', 'amount', 0.5, fn (float $v): int => intdiv((int) $v, 10_000));
        $this->assertGroupsMatch($oracle, $result, 0.5);
        $reader->close();
    }

    public function test_group_with_no_numeric_aggregate_reports_null_quantile(): void
    {
        // Group 2's amount cells are all text → count 0, quantile null.
        $oracle = $this->writeFixture(4_000, function (int $i): array {
            $g = $i <= 2_000 ? 1 : 2;
            $amount = $g === 1 ? round($i * 0.5, 2) : 'n/a';

            return [$g, $amount];
        });

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $result = $reader->groupQuantile('group', 'amount', 0.5);
        $this->assertGroupsMatch($oracle, $result, 0.5);

        $g2 = array_values(array_filter($result, fn ($r) => $r['group'] == 2))[0];
        $this->assertSame(0, $g2['count']);
        $this->assertNull($g2['quantile']);
        $reader->close();
    }
}

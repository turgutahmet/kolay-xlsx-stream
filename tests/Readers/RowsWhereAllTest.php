<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * G1 — rowsWhereAll(): multi-predicate AND with zone-map INTERSECTION.
 *
 * The surviving-block set of each predicate is intersected: a row that
 * satisfies every predicate must live in a block that survives every
 * predicate (zone maps are sound — they widen, never narrow), so the
 * intersection can only shrink the candidate set, never drop a match.
 * The pruned AND-scan must return exactly the full-scan oracle's rows.
 *
 * Correctness oracle is always the brute force: full scan + per-row
 * filter by all predicates.
 */
class RowsWhereAllTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-whereall-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * 2000 rows, sync every 100. col1 = id ascending (clusters tightly),
     * col2 = (id*7 % 2000) a permutation (scatters across blocks), col3
     * = untracked text.
     */
    private function writeFixture(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([1, 2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'perm', 'name']);
        for ($i = 1; $i <= 2000; $i++) {
            $writer->writeRow([$i, ($i * 7) % 2000, 'name-'.$i]);
        }
        $writer->finishFile();
    }

    /** @param list<array{0:int|string,1:string,2:int|float,3?:int|float}> $predicates */
    private function oracle(StreamingXlsxReader $reader, array $predicates): array
    {
        $rows = [];
        foreach ($reader->rows() as $row) {
            $ok = true;
            foreach ($predicates as $p) {
                $col0 = ((int) $p[0]) - 1;
                $v = is_numeric($row[$col0] ?? null) ? (float) $row[$col0] : null;
                $match = match ($p[1]) {
                    '=' => $v !== null && $v == $p[2],
                    '>=' => $v !== null && $v >= $p[2],
                    '<=' => $v !== null && $v <= $p[2],
                    '>' => $v !== null && $v > $p[2],
                    '<' => $v !== null && $v < $p[2],
                    'between' => $v !== null && $v >= $p[2] && $v <= $p[3],
                    default => false,
                };
                if (! $match) {
                    $ok = false;
                    break;
                }
            }
            if ($ok) {
                $rows[] = $row;
            }
        }

        return $rows;
    }

    public function test_and_scan_matches_full_scan_oracle(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $predicates = [[1, 'between', 400, 600], [2, 'between', 800, 1200]];
        $oracle = $this->oracle($reader, $predicates);
        $got = iterator_to_array($reader->rowsWhereAll($predicates), false);

        $this->assertSame($oracle, $got);
        $this->assertNotEmpty($got, 'fixture should produce some AND matches');
        $reader->close();
    }

    public function test_result_is_subset_of_each_single_predicate(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $p1 = [1, 'between', 400, 600];
        $p2 = [2, 'between', 800, 1200];

        $single1 = iterator_to_array($reader->rowsWhere($p1[0], $p1[1], $p1[2], $p1[3]), false);
        $single2 = iterator_to_array($reader->rowsWhere($p2[0], $p2[1], $p2[2], $p2[3]), false);
        $both = iterator_to_array($reader->rowsWhereAll([$p1, $p2]), false);

        // AND ⊆ each side.
        $this->assertLessThanOrEqual(count($single1), count($both));
        $this->assertLessThanOrEqual(count($single2), count($both));
        foreach ($both as $row) {
            $this->assertContains($row, $single1);
            $this->assertContains($row, $single2);
        }
        $reader->close();
    }

    public function test_predicate_without_stats_still_filters_per_row(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // col3 (name) carries no zone maps — it cannot prune blocks but
        // its predicate must still filter rows. Combine with a statful
        // col1 band so blocks are pruned by col1 and rows by col3.
        // name column is text → numeric predicates never match it, so
        // the whole AND is empty. Pair a matchable col1 with an
        // unmatchable col3 predicate.
        $got = iterator_to_array($reader->rowsWhereAll([[1, 'between', 400, 600], [3, '>=', 0]]), false);
        $this->assertSame([], $got, 'text column never satisfies a numeric predicate');
        $reader->close();
    }

    public function test_predicates_address_columns_by_name(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $byIndex = iterator_to_array($reader->rowsWhereAll([[1, 'between', 400, 600], [2, 'between', 800, 1200]]), false);
        $byName = iterator_to_array($reader->rowsWhereAll([['id', 'between', 400, 600], ['perm', 'between', 800, 1200]]), false);

        $this->assertSame($byIndex, $byName);
        $reader->close();
    }

    public function test_no_stats_file_falls_back_to_full_scan(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id', 'perm']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i, ($i * 7) % 200]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $predicates = [[1, '>=', 100], [2, '<=', 50]];
        $oracle = $this->oracle($reader, $predicates);
        $got = iterator_to_array($reader->rowsWhereAll($predicates), false);

        $this->assertSame($oracle, $got);
        $reader->close();
    }

    public function test_empty_predicates_throws(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\InvalidArgumentException::class);
        iterator_to_array($reader->rowsWhereAll([]));
    }
}

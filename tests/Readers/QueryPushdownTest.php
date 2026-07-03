<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * End-to-end tests for the KXSI query surface: columnStats(),
 * rowsWhere() (zone-map block pruning), findRow() and shards().
 *
 * The correctness oracle is always the brute-force path — every pruned
 * query must yield exactly what a full scan + per-row filter yields.
 * One test additionally pins that pruning really skips ahead (the
 * sheet stream starts at a sync-point offset, not at 0) so a regression
 * back to full scans can't hide behind identical results.
 */
class QueryPushdownTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-query-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * 500 data rows, sync every 100 -> ~5 blocks. Col 1 ascending id,
     * col 2 amount with a known outlier, col 3 untracked text.
     */
    private function writeFixture(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([1, 2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount', 'name']);
        for ($i = 1; $i <= 500; $i++) {
            $amount = $i === 321 ? 99999.5 : round($i * 1.25, 2);
            $writer->writeRow([$i, $amount, 'name-'.$i]);
        }
        $writer->finishFile();
    }

    public function test_column_stats_aggregates_from_sidecar(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $id = $reader->columnStats(1);
        $this->assertNotNull($id);
        $this->assertSame(1.0, $id['min']);
        $this->assertSame(500.0, $id['max']);
        $this->assertSame((float) (500 * 501 / 2), $id['sum']);
        $this->assertSame(500, $id['count']);
        $this->assertSame('asc', $id['sorted']);
        $this->assertSame(250.5, $id['avg']);

        $amount = $reader->columnStats(2);
        $this->assertSame(99999.5, $amount['max']);
        $this->assertNull($amount['sorted']); // the outlier breaks both orders

        $this->assertNull($reader->columnStats(3)); // untracked
        $reader->close();
    }

    public function test_rows_where_equals_matches_brute_force(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $hits = iterator_to_array($reader->rowsWhere(1, '=', 321));
        $this->assertCount(1, $hits);
        // id 321 was written as the 321st data row -> sheet row 322.
        $this->assertArrayHasKey(322, $hits);
        $this->assertSame('321', $hits[322][0]);
        $this->assertSame('99999.5', $hits[322][1]);
        $reader->close();
    }

    public function test_rows_where_between_spans_block_boundary(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // Rows with id 95..105 straddle the first sync point (~row 101).
        $hits = iterator_to_array($reader->rowsWhere(1, 'between', 95, 105));
        $this->assertCount(11, $hits);
        $this->assertSame(range(95, 105), array_map(fn ($r) => (int) $r[0], array_values($hits)));

        // Strictness of open ops.
        $lt = iterator_to_array($reader->rowsWhere(1, '<', 3));
        $this->assertSame([1, 2], array_map(fn ($r) => (int) $r[0], array_values($lt)));
        $lte = iterator_to_array($reader->rowsWhere(1, '<=', 3));
        $this->assertSame([1, 2, 3], array_map(fn ($r) => (int) $r[0], array_values($lte)));
        $reader->close();
    }

    public function test_rows_where_on_untracked_column_falls_back_to_scan(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // Column 3 carries no stats; text cells never match numeric
        // predicates, so this exercises the fallback path end to end.
        $this->assertSame([], iterator_to_array($reader->rowsWhere(3, '=', 42)));

        // And a numeric query on an untracked NUMERIC column (simulate by
        // querying col 2 pretending it were untracked — instead verify
        // tracked/untracked parity by comparing col 1 pruned results with
        // a brute-force filter over rows()).
        // rows() re-keys sequentially, rowsWhere() yields 1-based sheet
        // row numbers — compare row VALUES, which must match exactly.
        $brute = [];
        foreach ($reader->rows() as $row) {
            if (is_numeric($row[0] ?? null) && (float) $row[0] >= 490) {
                $brute[] = $row;
            }
        }
        $pruned = iterator_to_array($reader->rowsWhere(1, '>=', 490), false);
        $this->assertSame($brute, $pruned);
        $reader->close();
    }

    public function test_find_row_point_lookup(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $hit = $reader->findRow(1, 417);
        $this->assertNotNull($hit);
        $this->assertSame(418, $hit['row']);
        $this->assertSame('417', $hit['values'][0]);

        $this->assertNull($reader->findRow(1, 501));   // past the end
        $this->assertNull($reader->findRow(1, 0));     // before the start
        $reader->close();
    }

    public function test_pruned_query_seeks_instead_of_scanning(): void
    {
        $this->writeFixture();

        $spy = new class (new LocalFileSource($this->testFile)) implements Source {
            /** @var list<int> */
            public array $streamOffsets = [];

            public function __construct(private LocalFileSource $inner)
            {
            }

            public function size(): int
            {
                return $this->inner->size();
            }

            public function range(int $offset, int $length): string
            {
                return $this->inner->range($offset, $length);
            }

            public function streamFrom(int $offset)
            {
                $this->streamOffsets[] = $offset;

                return $this->inner->streamFrom($offset);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };

        $reader = StreamingXlsxReader::from($spy);

        // Target id 490 lives in the LAST block; a pruned read must open
        // the sheet stream well past its start.
        $reader->rows()->current(); // full scan: opens the stream at the sheet start
        $fullScanStart = min($spy->streamOffsets);
        $spy->streamOffsets = [];

        iterator_to_array($reader->rowsWhere(1, '=', 490));
        $this->assertNotEmpty($spy->streamOffsets);
        $this->assertGreaterThan($fullScanStart, min($spy->streamOffsets));
        $reader->close();
    }

    public function test_shards_cover_every_row_exactly_once(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $shards = $reader->shards(4);
        $this->assertGreaterThan(1, count($shards));

        // Shards must tile [1, totalRows] with no gap and no overlap.
        $expectedFirst = 1;
        foreach ($shards as $shard) {
            $this->assertSame($expectedFirst, $shard['first_row']);
            $expectedFirst = $shard['last_row'] + 1;
        }
        $this->assertSame($reader->rowCount(), end($shards)['last_row']);

        // JSON round-trip (the queue fan-out transport) + union check:
        // concatenating every shard's rows reproduces rows() exactly.
        // rowsForShard keys are 1-based sheet rows (no duplicates across
        // shards); rows() re-keys sequentially, so compare values.
        $union = [];
        foreach ($shards as $shard) {
            $shard = json_decode(json_encode($shard), true);
            foreach ($reader->rowsForShard($shard) as $rn => $row) {
                $this->assertArrayNotHasKey($rn, $union, 'row yielded by two shards');
                $union[$rn] = $row;
            }
        }
        $all = iterator_to_array($reader->rows(), false);
        ksort($union);
        $this->assertSame($all, array_values($union));
        $reader->close();
    }

    public function test_shards_degrade_to_single_shard_without_index(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id']);
        for ($i = 1; $i <= 50; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $shards = $reader->shards(4);

        $this->assertCount(1, $shards);
        $rows = iterator_to_array($reader->rowsForShard($shards[0]));
        $this->assertCount(51, $rows); // header + 50 data rows
        $reader->close();
    }

    public function test_rows_where_rejects_unknown_op(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->expectException(\InvalidArgumentException::class);
        iterator_to_array($reader->rowsWhere(1, '!=', 5));
    }
}

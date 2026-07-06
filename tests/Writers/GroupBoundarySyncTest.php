<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/** Read-filter that tallies every compressed byte the reader pulls. */
class GroupSyncByteCounter extends \php_user_filter
{
    public static int $bytes = 0;

    public function filter($in, $out, &$consumed, $closing): int
    {
        while ($bucket = stream_bucket_make_writeable($in)) {
            self::$bytes += $bucket->datalen;
            $consumed += $bucket->datalen;
            stream_bucket_append($out, $bucket);
        }

        return PSFS_PASS_ON;
    }
}

/**
 * D1 — syncAtGroupBoundaries(): align the writer's ZLIB_FULL_FLUSH sync
 * points (= index block boundaries) to GROUP changes, so every index
 * block holds exactly one group. groupStats() already folds a
 * group-pure block straight from the sidecar's per-block aggregates
 * without reading a row (foldGroupStatsFor's stats path); group-aligned
 * blocks make EVERY block pure, so a grouped aggregate reads ZERO rows.
 *
 * The sheet's decompressed bytes are unchanged (flush positions move,
 * XML content does not); only where deflate flushes and where the sidecar
 * records sync points differ. Default (no group sync) stays byte-identical.
 */
class GroupBoundarySyncTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-groupsync-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * 2000 rows in ascending groups of 137 (irregular vs any every-N
     * block size). Large sync period so WITHOUT group sync the whole
     * sheet is one mixed block; WITH it each group is its own pure block.
     */
    private function writeFixture(string $path, bool $groupSync): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->withRandomAccessIndex(every: 100000); // effectively never by row count
        $writer->withColumnStats([1, 2]);
        $writer->setBufferFlushInterval(200);
        if ($groupSync) {
            $writer->syncAtGroupBoundaries(1);
        }
        $writer->startFile(['grp', 'amount']);
        for ($i = 1; $i <= 2000; $i++) {
            $writer->writeRow([intdiv($i - 1, 137), $i * 1.0]);
        }
        $writer->finishFile();
    }

    private function makeSpy(string $path): object
    {
        return new class (new LocalFileSource($path)) implements Source {
            public int $streamCalls = 0;

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
                $this->streamCalls++;

                return $this->inner->streamFrom($offset, $length);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };
    }

    /** Brute-force GROUP BY grp: [grp => [sum, count, min, max]]. */
    private function oracle(): array
    {
        $g = [];
        for ($i = 1; $i <= 2000; $i++) {
            $key = intdiv($i - 1, 137);
            $v = $i * 1.0;
            if (! isset($g[$key])) {
                $g[$key] = ['sum' => 0.0, 'count' => 0, 'min' => $v, 'max' => $v];
            }
            $g[$key]['sum'] += $v;
            $g[$key]['count']++;
            $g[$key]['min'] = min($g[$key]['min'], $v);
            $g[$key]['max'] = max($g[$key]['max'], $v);
        }

        return $g;
    }

    public function test_group_aligned_blocks_answer_group_stats_from_the_sidecar(): void
    {
        $this->writeFixture($this->testFile, groupSync: true);

        $spy = $this->makeSpy($this->testFile);
        $reader = StreamingXlsxReader::from($spy);
        $reader->rowCount();       // warm index (uses range(), not streamFrom)
        $spy->streamCalls = 0;

        $stats = $reader->groupStats(1, 2);

        // Every GROUP block folds straight from the sidecar; only block 0
        // (the header rides in it) is ever scanned -> exactly one stream,
        // no matter how many groups there are.
        $this->assertSame(1, $spy->streamCalls, 'only the header block should be scanned');

        // ...and the answer is exact.
        $oracle = $this->oracle();
        $this->assertCount(count($oracle), $stats);
        foreach ($stats as $row) {
            $g = (int) $row['group'];
            $this->assertArrayHasKey($g, $oracle);
            $this->assertSame($oracle[$g]['count'], $row['count'], "count for group {$g}");
            $this->assertEqualsWithDelta($oracle[$g]['sum'], $row['sum'], 1e-6, "sum for group {$g}");
            $this->assertEqualsWithDelta($oracle[$g]['min'], $row['min'], 1e-6, "min for group {$g}");
            $this->assertEqualsWithDelta($oracle[$g]['max'], $row['max'], 1e-6, "max for group {$g}");
        }
        $reader->close();
    }

    /** Compressed bytes the reader pulls while answering groupStats(). */
    private function groupStatsBytes(string $path): int
    {
        $inner = new LocalFileSource($path);
        $source = new class ($inner) implements Source {
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
                $h = $this->inner->streamFrom($offset, $length);
                stream_filter_append($h, 'group_sync_byte_counter', STREAM_FILTER_READ);

                return $h;
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };

        $reader = StreamingXlsxReader::from($source);
        $reader->rowCount();
        GroupSyncByteCounter::$bytes = 0;
        $reader->groupStats(1, 2);
        $reader->close();

        return GroupSyncByteCounter::$bytes;
    }

    public function test_group_sync_reads_far_fewer_bytes_than_without(): void
    {
        stream_filter_register('group_sync_byte_counter', GroupSyncByteCounter::class);

        $synced = $this->testFile;
        $plain = $this->testFile.'.plain.xlsx';
        $this->writeFixture($synced, groupSync: true);
        $this->writeFixture($plain, groupSync: false);

        try {
            $syncedBytes = $this->groupStatsBytes($synced);
            $plainBytes = $this->groupStatsBytes($plain);

            // Group sync scans only the small header block; without it the
            // one mixed block IS the whole sheet, so groupStats tokenizes
            // every row. Expect a large reduction, not a marginal one.
            $this->assertGreaterThan(0, $syncedBytes);
            $this->assertLessThan($plainBytes / 2, $syncedBytes, 'group sync must read far fewer compressed bytes');
        } finally {
            @unlink($plain);
        }
    }

    public function test_reader_reads_group_synced_file_row_for_row(): void
    {
        $this->writeFixture($this->testFile, groupSync: true);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $rows = iterator_to_array($reader->rows(skip: 1), false);
        $this->assertCount(2000, $rows);
        $this->assertSame(['0', '1'], $rows[0]);              // i=1: grp 0, amount 1
        $this->assertSame((string) intdiv(2000 - 1, 137), $rows[1999][0]);
        $reader->close();
    }

    public function test_enables_random_access_index(): void
    {
        // syncAtGroupBoundaries without an explicit withRandomAccessIndex
        // must still produce a queryable sidecar (it enables the index).
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([1, 2]);
        $writer->syncAtGroupBoundaries(1);
        $writer->startFile(['grp', 'amount']);
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([intdiv($i - 1, 50), $i * 1.0]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNotNull($reader->columnStats(2));
        $this->assertSame(10, count($reader->groupStats(1, 2)));
        $reader->close();
    }

    public function test_throws_after_start_file(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['grp', 'amount']);

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxStreamException::class);
        $writer->syncAtGroupBoundaries(1);
    }
}

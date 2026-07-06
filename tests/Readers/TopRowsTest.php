<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * F1 — topRows(): "ORDER BY column [DESC] LIMIT k" over the sidecar.
 *
 * On a column the writer observed to be sorted, the k extremes sit at
 * one end of the sheet, so topRows reads only the blocks at that end and
 * early-exits — the spreadsheet form of an indexed ORDER BY … LIMIT. On
 * an unsorted column it degrades to a single full scan holding an O(k)
 * heap, still correct. The oracle is always the brute-force sort of a
 * full scan.
 */
class TopRowsTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-toprows-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /** col1 = id ascending, col2 = a permutation (unsorted), col3 text. */
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

    /** Brute-force ORDER BY col LIMIT k, returning [rn => row] in rank order. */
    private function oracle(StreamingXlsxReader $reader, int $col, int $k, bool $desc): array
    {
        $idx = $col - 1;
        $all = [];
        foreach ($reader->rowRange(2, $reader->rowCount()) as $rn => $row) {
            if (is_numeric($row[$idx] ?? null)) {
                $all[$rn] = (float) $row[$idx];
            }
        }
        uasort($all, fn ($a, $b) => $desc ? ($b <=> $a) : ($a <=> $b));

        return array_slice($all, 0, $k, true);
    }

    private function makeSpy(): object
    {
        return new class (new LocalFileSource($this->testFile)) implements Source {
            /** @var list<int> */
            public array $offsets = [];

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
                $this->offsets[] = $offset;

                return $this->inner->streamFrom($offset, $length);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };
    }

    public function test_top_largest_on_sorted_column(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $top = $reader->topRows(1, 5, desc: true);
        $this->assertSame([2001, 2000, 1999, 1998, 1997], array_keys($top));
        $this->assertSame([2000, 1999, 1998, 1997, 1996], array_map(fn ($r) => (int) $r[0], array_values($top)));

        $this->assertSame($this->oracle($reader, 1, 5, true), array_map(fn ($r) => (float) $r[0], $top));
        $reader->close();
    }

    public function test_top_smallest_on_sorted_column(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $top = $reader->topRows(1, 3, desc: false);
        $this->assertSame([2, 3, 4], array_keys($top));
        $this->assertSame([1, 2, 3], array_map(fn ($r) => (int) $r[0], array_values($top)));
        $reader->close();
    }

    public function test_sorted_top_reads_only_the_end_blocks(): void
    {
        $this->writeFixture();
        $spy = $this->makeSpy();
        $reader = StreamingXlsxReader::from($spy);

        // Warm the index + establish the full-scan start offset.
        $reader->rowCount();
        $spy->offsets = [];

        $reader->topRows(1, 5, desc: true);

        // The largest ids live in the LAST block — the read must seek well
        // past the sheet start (not a full scan from offset ~0).
        $this->assertNotEmpty($spy->offsets);
        $this->assertGreaterThan(0, min($spy->offsets), 'sorted topRows must seek to the end, not scan from the start');
        $reader->close();
    }

    public function test_top_on_unsorted_column_matches_oracle(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        foreach ([true, false] as $desc) {
            $top = $reader->topRows(2, 7, desc: $desc);
            $oracle = $this->oracle($reader, 2, 7, $desc);
            $this->assertSame(array_keys($oracle), array_keys($top), 'row order mismatch (desc='.var_export($desc, true).')');
            $this->assertSame(
                array_values($oracle),
                array_map(fn ($r) => (float) $r[1], array_values($top))
            );
        }
        $reader->close();
    }

    public function test_k_larger_than_row_count_returns_all(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $top = $reader->topRows(1, 999999, desc: true);
        $this->assertCount(2000, $top);
        $reader->close();
    }

    public function test_k_zero_returns_empty(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertSame([], $reader->topRows(1, 0));
        $reader->close();
    }

    /**
     * A sorted column may carry non-numeric cells (optional amount/discount
     * columns) — the sorted flag orders only numeric cells. The extremes
     * must be the true numeric extremes, never the null/text cells that
     * happen to sit at the sheet end. Regression: the fast path used to
     * read the fixed k rows at the end and return them verbatim, so a
     * null-tailed asc column returned empties instead of the real maxima.
     */
    private function writeNullEndFixture(bool $nullsAtTail): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'score']);
        // Asc-sorted numeric core; a run of null-score rows at one end.
        if ($nullsAtTail) {
            for ($i = 1; $i <= 1000; $i++) {
                $writer->writeRow([$i, $i * 1.5]);
            }
            for ($i = 1001; $i <= 1015; $i++) {
                $writer->writeRow([$i, null]);
            }
        } else {
            for ($i = 1; $i <= 15; $i++) {
                $writer->writeRow([$i, null]);
            }
            for ($i = 16; $i <= 1015; $i++) {
                $writer->writeRow([$i, ($i - 15) * 1.5]);
            }
        }
        $writer->finishFile();
    }

    public function test_largest_ignores_null_tail_on_sorted_column(): void
    {
        $this->writeNullEndFixture(nullsAtTail: true);
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $top = $reader->topRows(2, 3, desc: true);
        $this->assertSame(
            [1500.0, 1498.5, 1497.0],
            array_map(fn ($r) => (float) $r[1], array_values($top))
        );
        $this->assertSame($this->oracle($reader, 2, 3, true), array_map(fn ($r) => (float) $r[1], $top));
        $reader->close();
    }

    public function test_smallest_ignores_null_head_on_sorted_column(): void
    {
        $this->writeNullEndFixture(nullsAtTail: false);
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $top = $reader->topRows(2, 3, desc: false);
        $this->assertSame(
            [1.5, 3.0, 4.5],
            array_map(fn ($r) => (float) $r[1], array_values($top))
        );
        $this->assertSame($this->oracle($reader, 2, 3, false), array_map(fn ($r) => (float) $r[1], $top));
        $reader->close();
    }

    public function test_null_scattered_including_extreme_window_matches_oracle(): void
    {
        // Nulls scattered through the column, some landing in the last k
        // rows — the exact shape that slipped a null into the result.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100)->withColumnStats([2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'score']);
        for ($i = 1; $i <= 1000; $i++) {
            $writer->writeRow([$i, $i % 25 === 0 ? null : $i * 1.5]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        foreach ([true, false] as $desc) {
            $top = $reader->topRows(2, 8, desc: $desc);
            $oracle = $this->oracle($reader, 2, 8, $desc);
            $this->assertSame(array_keys($oracle), array_keys($top), 'desc='.var_export($desc, true));
            $this->assertSame($oracle, array_map(fn ($r) => (float) $r[1], $top));
            // No non-numeric cell may appear in the result.
            foreach ($top as $row) {
                $this->assertIsNumeric($row[1]);
            }
        }
        $reader->close();
    }

    public function test_addresses_column_by_name(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $this->assertSame(
            $reader->topRows(2, 5, desc: true),
            $reader->topRows('perm', 5, desc: true)
        );
        $reader->close();
    }
}

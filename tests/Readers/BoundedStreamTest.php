<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Contracts\SupportsBoundedStream;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * G6 — bounded streamFrom(offset, ?length).
 *
 * A pruned query that scans a NON-terminal run of blocks must open the
 * sheet stream with a byte length that stops at the run's trailing sync
 * point instead of ranging to the entry's end. On S3 this turns an
 * open-ended [offset, EOF] GET into an exact [offset, offset+length-1]
 * fetch; on local files it caps the inflate read. The rows returned must
 * be byte-for-byte the full-scan oracle's rows — bounding is an I/O
 * optimization, never a correctness change.
 *
 * The load-bearing gate is the byte-oracle: bounded reads end at a
 * ZLIB_FULL_FLUSH sync boundary, so the inflate stream is fed to that
 * boundary with NO_FLUSH and stopped WITHOUT a ZLIB_FINISH step (the
 * flush already emitted every row up to it — the exact inverse of the
 * seek-to-sync-point entry the random-access index already relies on).
 */
class BoundedStreamTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-bounded-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /** 2000 rows, sync every 100 -> ~20 blocks; col 1 ascending. */
    private function writeFixture(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([1, 2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount', 'name']);
        for ($i = 1; $i <= 2000; $i++) {
            $writer->writeRow([$i, round($i * 1.25, 2), 'name-'.$i]);
        }
        $writer->finishFile();
    }

    /** Spy Source recording each open: unbounded streamFrom (length null) vs bounded streamFromRange. */
    private function makeSpy(): object
    {
        return new class (new LocalFileSource($this->testFile)) implements Source, SupportsBoundedStream {
            /** @var list<array{offset: int, length: int|null}> */
            public array $streamCalls = [];

            public function __construct(private LocalFileSource $inner) {}

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
                $this->streamCalls[] = ['offset' => $offset, 'length' => null];

                return $this->inner->streamFrom($offset);
            }

            public function streamFromRange(int $offset, int $length)
            {
                $this->streamCalls[] = ['offset' => $offset, 'length' => $length];

                return $this->inner->streamFromRange($offset, $length);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };
    }

    public function test_non_terminal_run_is_fetched_with_a_bounded_length(): void
    {
        $this->writeFixture();
        $spy = $this->makeSpy();
        $reader = StreamingXlsxReader::from($spy);

        // id in [200, 300] lives in early blocks — the run does NOT reach
        // the last block, so the fetch must be bounded.
        $hits = iterator_to_array($reader->rowsWhere(1, 'between', 200, 300));

        $this->assertSame(range(200, 300), array_map(fn ($r) => (int) $r[0], array_values($hits)));

        $bounded = array_filter($spy->streamCalls, fn ($c) => $c['length'] !== null);
        $this->assertNotEmpty($bounded, 'a non-terminal pruned run must pass a bounded length to streamFrom');

        // The bound must be strictly smaller than "from offset to entry end".
        $size = $spy->size();
        foreach ($bounded as $c) {
            $this->assertLessThan($size - $c['offset'], $c['length'], 'bounded length should stop before EOF');
            $this->assertGreaterThan(0, $c['length']);
        }
        $reader->close();
    }

    public function test_bounded_rows_equal_unbounded_oracle(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // Oracle: brute-force filter over a full scan.
        $oracle = [];
        foreach ($reader->rows() as $row) {
            $v = (float) ($row[0] ?? 0);
            if ($v >= 200 && $v <= 300) {
                $oracle[] = $row;
            }
        }

        $pruned = iterator_to_array($reader->rowsWhere(1, 'between', 200, 300), false);
        $this->assertSame($oracle, $pruned, 'bounded pruned scan must equal the full-scan oracle');
        $reader->close();
    }

    public function test_tail_query_stays_correct_under_bounding(): void
    {
        // A run reaching the last DATA block is still bounded to the final
        // sync point (the trailing zero-count block fails the count>0
        // survivor test, so the run stops just before it) — which is only
        // ever a tighter fetch. Correctness must hold regardless: the tail
        // rows come back exactly, byte-for-byte with the full-scan oracle.
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $oracle = [];
        foreach ($reader->rows() as $row) {
            if ((float) ($row[0] ?? 0) >= 1990) {
                $oracle[] = $row;
            }
        }

        $pruned = iterator_to_array($reader->rowsWhere(1, '>=', 1990), false);
        $this->assertSame($oracle, $pruned);
        $this->assertSame(range(1990, 2000), array_map(fn ($r) => (int) $r[0], $pruned));
        $reader->close();
    }

    public function test_full_scan_reads_to_eof_unbounded(): void
    {
        // The FINISH path (read to entry end) must be preserved for full
        // scans: rows() passes no offset and no length, so every stream is
        // opened unbounded exactly as before this feature.
        $this->writeFixture();
        $spy = $this->makeSpy();
        $reader = StreamingXlsxReader::from($spy);

        $count = iterator_count($reader->rows());
        $this->assertSame(2001, $count); // header + 2000 rows

        $this->assertNotEmpty($spy->streamCalls);
        foreach ($spy->streamCalls as $c) {
            $this->assertNull($c['length'], 'full scan must open streams unbounded');
        }
        $reader->close();
    }
}

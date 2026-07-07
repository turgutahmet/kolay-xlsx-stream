<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * STRZ (G3) reader query — string rowsWhere/findRow over the STRZ zone
 * maps. The correctness oracle is always the brute-force full scan: every
 * pruned string query must yield exactly what a full scan + per-row
 * strcmp filter yields, across =, prefix, between, and the open ops.
 * "string findRow on S3" is the headline; soundness is the gate.
 */
class StringQueryTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-strq-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * col1 = sorted invoice no (common prefix), col2 = a shuffled tag
     * (unsorted string), col3 = untracked text. 2000 rows, sync every 100.
     */
    private function writeFixture(): void
    {
        mt_srand(7);
        $tags = [];
        for ($i = 1; $i <= 2000; $i++) {
            $tags[$i] = 'TAG-'.str_pad((string) mt_rand(0, 300), 4, '0', STR_PAD_LEFT);
        }
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withStringStats([1, 2])->withRandomAccessIndex(every: 100)->setBufferFlushInterval(100);
        $writer->startFile(['invoice', 'tag', 'note']);
        for ($i = 1; $i <= 2000; $i++) {
            $writer->writeRow([sprintf('INV-2024-%08d', $i), $tags[$i], 'note-'.$i]);
        }
        $writer->finishFile();
    }

    /** Brute-force oracle: full scan, per-row strcmp filter; [rn => cellValue]. */
    private function oracle(StreamingXlsxReader $reader, int $col, string $op, string $v, ?string $v2 = null): array
    {
        $idx = $col - 1;
        $hits = [];
        // From row 1 (the header): rowsWhere scans every row including the
        // header, and the block-0 fold makes pruned == unpruned — so a
        // header cell that satisfies the predicate (e.g. "invoice" >= a
        // data value) is a real match, exactly as in the numeric path.
        foreach ($reader->rowRange(1, $reader->rowCount()) as $rn => $row) {
            $s = (string) ($row[$idx] ?? '');
            if ($s === '') {
                continue;
            }
            $lo = $v2 !== null && strcmp($v, $v2) > 0 ? $v2 : $v;
            $hi = $v2 !== null && strcmp($v, $v2) > 0 ? $v : $v2;
            $match = match ($op) {
                '=' => strcmp($s, $v) === 0,
                '<' => strcmp($s, $v) < 0,
                '<=' => strcmp($s, $v) <= 0,
                '>' => strcmp($s, $v) > 0,
                '>=' => strcmp($s, $v) >= 0,
                'prefix' => str_starts_with($s, $v),
                'between' => strcmp($s, $lo) >= 0 && strcmp($s, (string) $hi) <= 0,
                default => false,
            };
            if ($match) {
                $hits[$rn] = $s;
            }
        }

        return $hits;
    }

    private function pruned(StreamingXlsxReader $reader, int|string $col, string $op, string $v, ?string $v2 = null): array
    {
        $out = [];
        $idx = (is_int($col) ? $col : 1) - 1;
        foreach ($reader->rowsWhere($col, $op, $v, $v2) as $rn => $row) {
            $out[$rn] = (string) $row[is_int($col) ? $col - 1 : $idx];
        }

        return $out;
    }

    public function test_equals_prefix_between_open_ops_match_oracle(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // Sorted common-prefix column (col1) across every op.
        $cases = [
            [1, '=', 'INV-2024-00000137', null],
            [1, 'prefix', 'INV-2024-000012', null],   // 12xx band
            [1, 'between', 'INV-2024-00000500', 'INV-2024-00000600'],
            [1, '<', 'INV-2024-00000003', null],
            [1, '>=', 'INV-2024-00001998', null],
            // Unsorted tag column (col2) — pruning weaker but soundness holds.
            [2, '=', 'TAG-0150', null],
            [2, 'prefix', 'TAG-02', null],
            [2, 'between', 'TAG-0100', 'TAG-0110'],
        ];
        foreach ($cases as [$col, $op, $v, $v2]) {
            $this->assertSame(
                $this->oracle($reader, $col, $op, $v, $v2),
                $this->pruned($reader, $col, $op, $v, $v2),
                "op={$op} col={$col} v={$v}"
            );
        }
        $reader->close();
    }

    public function test_find_row_by_string_key(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $hit = $reader->findRow('invoice', 'INV-2024-00000417');
        $this->assertNotNull($hit);
        $this->assertSame(418, $hit['row']); // header row 1 + data row 417
        $this->assertSame('INV-2024-00000417', $hit['values'][0]);

        $this->assertNull($reader->findRow('invoice', 'INV-2024-99999999'));
        $reader->close();
    }

    public function test_addresses_string_column_by_name(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $byName = iterator_to_array($reader->rowsWhere('invoice', 'prefix', 'INV-2024-000005'));
        $byIndex = iterator_to_array($reader->rowsWhere(1, 'prefix', 'INV-2024-000005'));
        $this->assertSame(array_keys($byIndex), array_keys($byName));
        $this->assertNotEmpty($byName);
        $reader->close();
    }

    public function test_sorted_string_findrow_prunes_to_few_blocks(): void
    {
        $this->writeFixture();

        $spy = new class (new LocalFileSource($this->testFile)) implements Source {
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

        $reader = StreamingXlsxReader::from($spy);
        $reader->rowCount(); // warm index
        $spy->offsets = [];

        // A sorted-column point lookup deep in the sheet must seek past the
        // start (pruned), not scan from row 1.
        $reader->findRow('invoice', 'INV-2024-00001500');
        $this->assertNotEmpty($spy->offsets);
        $this->assertGreaterThan(0, min($spy->offsets), 'sorted string findRow did not prune (scanned from start)');
        $reader->close();
    }

    public function test_soundness_random_probes(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        mt_srand(99);
        for ($p = 0; $p < 60; $p++) {
            $v = sprintf('INV-2024-%08d', mt_rand(1, 2000));
            $this->assertSame(
                $this->oracle($reader, 1, '=', $v),
                $this->pruned($reader, 1, '=', $v),
                "probe {$v}"
            );
        }
        $reader->close();
    }

    public function test_untracked_string_column_falls_back_to_scan(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // col3 (note) carries no STRZ — must still answer correctly via scan.
        $hits = iterator_to_array($reader->rowsWhere(3, '=', 'note-42'));
        $this->assertCount(1, $hits);
        $this->assertSame('note-42', array_values($hits)[0][2]);
        $reader->close();
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * End-to-end random-access tests.
 *
 * Cover both paths:
 *   - Indexed file (writer's withRandomAccessIndex opt-in) → seeks via
 *     the xl/_kxs/index.bin sidecar, fresh-init inflate from sync point.
 *   - Plain file (no opt-in) → falls back to a sequential scan.
 *
 * Both paths must yield identical row content and identical rowCount;
 * only the cost differs. The tests assert correctness, not performance.
 */
class RandomAccessReadTest extends TestCase
{
    private string $indexed;
    private string $plain;

    protected function setUp(): void
    {
        parent::setUp();
        $tmp = sys_get_temp_dir();
        $token = uniqid('', true);
        $this->indexed = "{$tmp}/kxs-ra-indexed-{$token}.xlsx";
        $this->plain = "{$tmp}/kxs-ra-plain-{$token}.xlsx";
    }

    protected function tearDown(): void
    {
        @unlink($this->indexed);
        @unlink($this->plain);
        parent::tearDown();
    }

    public function test_row_at_returns_targeted_row_with_index(): void
    {
        $this->writeIndexed($this->indexed, syncEvery: 100, dataRows: 1000);

        $reader = StreamingXlsxReader::fromFile($this->indexed);

        $this->assertSame(['id', 'value'], $reader->rowAt(1));         // header
        $this->assertSame(['1', 'val-1'], $reader->rowAt(2));          // first data
        $this->assertSame(['250', 'val-250'], $reader->rowAt(251));    // mid-stream
        $this->assertSame(['1000', 'val-1000'], $reader->rowAt(1001)); // last
        $this->assertNull($reader->rowAt(1002));                       // past end
    }

    public function test_row_at_returns_targeted_row_without_index(): void
    {
        $this->writePlain($this->plain, dataRows: 500);

        $reader = StreamingXlsxReader::fromFile($this->plain);

        $this->assertSame(['id', 'value'], $reader->rowAt(1));
        $this->assertSame(['1', 'val-1'], $reader->rowAt(2));
        $this->assertSame(['200', 'val-200'], $reader->rowAt(201));
        $this->assertSame(['500', 'val-500'], $reader->rowAt(501));
        $this->assertNull($reader->rowAt(502));
    }

    public function test_row_at_indexed_and_plain_agree(): void
    {
        $this->writeIndexed($this->indexed, syncEvery: 100, dataRows: 600);
        $this->writePlain($this->plain, dataRows: 600);

        $a = StreamingXlsxReader::fromFile($this->indexed);
        $b = StreamingXlsxReader::fromFile($this->plain);

        foreach ([1, 2, 50, 100, 250, 500, 601] as $rn) {
            $this->assertSame(
                $a->rowAt($rn),
                $b->rowAt($rn),
                "row {$rn} should match between indexed and plain"
            );
        }
    }

    public function test_row_at_invalid_inputs_return_null(): void
    {
        $this->writeIndexed($this->indexed, syncEvery: 100, dataRows: 100);
        $reader = StreamingXlsxReader::fromFile($this->indexed);

        $this->assertNull($reader->rowAt(0));
        $this->assertNull($reader->rowAt(-5));
        $this->assertNull($reader->rowAt(999_999));
    }

    public function test_row_range_inclusive_bounds_with_index(): void
    {
        $this->writeIndexed($this->indexed, syncEvery: 100, dataRows: 1000);

        $reader = StreamingXlsxReader::fromFile($this->indexed);
        $rows = iterator_to_array($reader->rowRange(250, 253), true);

        $this->assertCount(4, $rows);
        $this->assertSame([250, 251, 252, 253], array_keys($rows));
        $this->assertSame(['249', 'val-249'], $rows[250]);
        $this->assertSame(['252', 'val-252'], $rows[253]);
    }

    public function test_row_range_inclusive_bounds_without_index(): void
    {
        $this->writePlain($this->plain, dataRows: 500);

        $reader = StreamingXlsxReader::fromFile($this->plain);
        $rows = iterator_to_array($reader->rowRange(2, 5), true);

        $this->assertSame([2, 3, 4, 5], array_keys($rows));
        $this->assertSame(['1', 'val-1'], $rows[2]);
        $this->assertSame(['4', 'val-4'], $rows[5]);
    }

    public function test_row_range_indexed_and_plain_agree(): void
    {
        $this->writeIndexed($this->indexed, syncEvery: 100, dataRows: 800);
        $this->writePlain($this->plain, dataRows: 800);

        $a = iterator_to_array(
            StreamingXlsxReader::fromFile($this->indexed)->rowRange(150, 320),
            true
        );
        $b = iterator_to_array(
            StreamingXlsxReader::fromFile($this->plain)->rowRange(150, 320),
            true
        );

        $this->assertSame($a, $b);
    }

    public function test_row_range_handles_empty_or_inverted_ranges(): void
    {
        $this->writeIndexed($this->indexed, syncEvery: 100, dataRows: 100);
        $reader = StreamingXlsxReader::fromFile($this->indexed);

        // Inverted bounds — no rows.
        $this->assertSame([], iterator_to_array($reader->rowRange(50, 10), true));

        // Range fully past EOF — no rows.
        $this->assertSame([], iterator_to_array($reader->rowRange(500, 600), true));

        // Range that overlaps EOF — emits up to the last row.
        $rows = iterator_to_array($reader->rowRange(99, 200), true);
        $this->assertCount(3, $rows);              // rows 99, 100, 101
        $this->assertSame([99, 100, 101], array_keys($rows));
    }

    public function test_row_count_is_constant_time_with_index(): void
    {
        // Functional check — both paths must produce the same total
        // including the header row. Performance isn't asserted here;
        // the index removes the iterate-and-count work from rowCount().
        $this->writeIndexed($this->indexed, syncEvery: 100, dataRows: 1000);
        $this->writePlain($this->plain, dataRows: 1000);

        $a = StreamingXlsxReader::fromFile($this->indexed);
        $b = StreamingXlsxReader::fromFile($this->plain);

        $this->assertSame(1001, $a->rowCount());
        $this->assertSame(1001, $b->rowCount());
        $this->assertSame($a->rowCount(), $b->rowCount());
    }

    public function test_random_access_after_row_iteration_preserves_state(): void
    {
        // Generators in rows() open and close their own stream; rowAt
        // must work afterwards without holding stale handles.
        $this->writeIndexed($this->indexed, syncEvery: 100, dataRows: 200);

        $reader = StreamingXlsxReader::fromFile($this->indexed);
        $first10 = iterator_to_array($reader->rowRange(2, 11), true);

        $this->assertCount(10, $first10);
        $this->assertSame(['180', 'val-180'], $reader->rowAt(181));
        $this->assertSame(1, $reader->rowAt(1)[0] === 'id' ? 1 : 0);
    }

    public function test_row_at_in_multi_sheet_indexed_workbook(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->indexed));
        $writer->withRandomAccessIndex(every: 50);
        $writer->setBufferFlushInterval(50);
        $writer->startFile(['s1']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow(["sheet1-row-{$i}"]);
        }
        $writer->newSheet('Other', ['s2']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow(["sheet2-row-{$i}"]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->indexed);

        $this->assertSame(['s1'], $reader->rowAt(1));
        $this->assertSame(['sheet1-row-150'], $reader->rowAt(151));

        $reader->onSheet('Other');
        $this->assertSame(['s2'], $reader->rowAt(1));
        $this->assertSame(['sheet2-row-250'], $reader->rowAt(251));
    }

    private function writeIndexed(string $path, int $syncEvery, int $dataRows): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->withRandomAccessIndex(every: $syncEvery);
        $writer->setBufferFlushInterval($syncEvery);
        $writer->startFile(['id', 'value']);
        for ($i = 1; $i <= $dataRows; $i++) {
            $writer->writeRow([$i, "val-{$i}"]);
        }
        $writer->finishFile();
    }

    private function writePlain(string $path, int $dataRows): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($path));
        $writer->startFile(['id', 'value']);
        for ($i = 1; $i <= $dataRows; $i++) {
            $writer->writeRow([$i, "val-{$i}"]);
        }
        $writer->finishFile();
    }
}

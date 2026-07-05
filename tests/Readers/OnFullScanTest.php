<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * E2 — onFullScan(): an observability hook fired when a query CANNOT use
 * the index and degrades to a full row scan (an unindexed column, an
 * un-sorted groupBy). It lets callers log/alert when a query isn't
 * getting the pushdown they expect — the difference between a
 * millisecond sidecar answer and reading the whole sheet.
 */
class OnFullScanTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-fullscan-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /** col1 indexed (stats), col2 NOT indexed. */
    private function writeFixture(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100)->withColumnStats([1]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['indexed', 'plain']);
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([$i, $i * 3]);
        }
        $writer->finishFile();
    }

    public function test_hook_fires_when_querying_an_unindexed_column(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $events = [];
        $reader->onFullScan(function (array $ctx) use (&$events) {
            $events[] = $ctx;
        });

        iterator_to_array($reader->rowsWhere(2, '>=', 100)); // col2 has no zone maps

        $this->assertNotEmpty($events, 'querying an unindexed column must fire onFullScan');
        $this->assertSame('rowsWhere', $events[0]['query']);
        $this->assertSame(2, $events[0]['column']);
        $reader->close();
    }

    public function test_hook_does_not_fire_when_pushdown_applies(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $fired = false;
        $reader->onFullScan(function () use (&$fired) {
            $fired = true;
        });

        iterator_to_array($reader->rowsWhere(1, '>=', 400)); // col1 IS indexed
        $this->assertFalse($fired, 'an indexed query must not fire onFullScan');
        $reader->close();
    }

    public function test_hook_fires_for_group_stats_without_pushdown_basis(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $events = [];
        $reader->onFullScan(function (array $ctx) use (&$events) {
            $events[] = $ctx;
        });

        // groupBy col2 has no stats -> no pushdown -> full scan.
        $reader->groupStats(2, 1);
        $this->assertNotEmpty($events);
        $this->assertSame('groupStats', $events[0]['query']);
        $reader->close();
    }

    public function test_hook_fires_for_rows_where_all_without_indexed_predicate(): void
    {
        $this->writeFixture();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $events = [];
        $reader->onFullScan(function (array $ctx) use (&$events) {
            $events[] = $ctx;
        });

        // Neither predicate touches an indexed column.
        iterator_to_array($reader->rowsWhereAll([[2, '>=', 100], [2, '<=', 400]]));
        $this->assertNotEmpty($events);
        $this->assertSame('rowsWhereAll', $events[0]['query']);
        $reader->close();
    }
}

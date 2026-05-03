<?php

namespace Kolay\XlsxStream\Tests;

use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

class V21FeaturesTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/v21_'.uniqid().'.xlsx';
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }

    // ── Generator / iterable writeRows ──────────────────────────────

    public function test_write_rows_accepts_generator()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['ID', 'Name']);

        $gen = (function () {
            for ($i = 1; $i <= 100; $i++) {
                yield [$i, "User $i"];
            }
        })();

        $writer->writeRows($gen);
        $stats = $writer->finishFile();

        $this->assertEquals(100, $stats['rows']);
    }

    public function test_write_rows_accepts_array_still()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['ID']);

        $writer->writeRows([[1], [2], [3]]);
        $stats = $writer->finishFile();

        $this->assertEquals(3, $stats['rows']);
    }

    public function test_write_rows_accepts_iterator_aggregate()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['ID']);

        $rows = new \ArrayIterator([[1], [2], [3], [4], [5]]);
        $writer->writeRows($rows);
        $stats = $writer->finishFile();

        $this->assertEquals(5, $stats['rows']);
    }

    // ── onProgress callback ─────────────────────────────────────────

    public function test_on_progress_fires_at_correct_interval()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $events = [];
        $writer->onProgress(function (int $rows, int $bytes) use (&$events) {
            $events[] = ['rows' => $rows, 'bytes' => $bytes];
        });
        $writer->setProgressInterval(100);

        $writer->startFile(['ID']);
        for ($i = 1; $i <= 1000; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        // 100, 200, 300, ..., 1000 → exactly 10 events
        $this->assertCount(10, $events);
        $this->assertEquals(100, $events[0]['rows']);
        $this->assertEquals(1000, $events[9]['rows']);
        $this->assertGreaterThan(0, $events[9]['bytes']);
    }

    public function test_on_progress_default_interval_is_10000()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $events = 0;
        $writer->onProgress(function () use (&$events) {
            $events++;
        });

        $writer->startFile(['ID']);
        for ($i = 1; $i <= 5000; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        // 5000 rows < 10000 default interval → no events
        $this->assertEquals(0, $events);
    }

    public function test_on_progress_with_no_callback_has_no_overhead()
    {
        // Smoke test — just confirm no callback path doesn't error
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['ID']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i]);
        }
        $stats = $writer->finishFile();

        $this->assertEquals(100, $stats['rows']);
    }

    public function test_on_progress_invalid_interval_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxStreamException::class);
        $writer->setProgressInterval(0);
    }

    // ── forDisk Laravel Storage integration ─────────────────────────

    public function test_for_disk_local_writes_to_resolved_path()
    {
        $tmpRoot = sys_get_temp_dir().'/xlsx_disk_'.uniqid();
        config([
            'filesystems.disks.testing' => [
                'driver' => 'local',
                'root' => $tmpRoot,
            ],
        ]);

        $writer = SinkableXlsxWriter::forDisk('testing', 'exports/test.xlsx');
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $this->assertFileExists($tmpRoot.'/exports/test.xlsx');

        // Cleanup
        unlink($tmpRoot.'/exports/test.xlsx');
        rmdir($tmpRoot.'/exports');
        rmdir($tmpRoot);
    }

    public function test_for_disk_unknown_disk_throws()
    {
        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxStreamException::class);
        $this->expectExceptionMessage('not configured');

        SinkableXlsxWriter::forDisk('nonexistent', 'file.xlsx');
    }

    public function test_for_disk_unsupported_driver_throws()
    {
        config([
            'filesystems.disks.weird' => [
                'driver' => 'ftp',
                'host' => 'example.com',
            ],
        ]);

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxStreamException::class);
        $this->expectExceptionMessage('not supported');

        SinkableXlsxWriter::forDisk('weird', 'file.xlsx');
    }
}

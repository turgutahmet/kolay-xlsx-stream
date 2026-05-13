<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Pins the bounded-RAM contract that is the reader's headline claim.
 * Reads 100K rows and asserts that PHP's reported peak-vs-baseline
 * delta stays under 4 MB. Catches regressions where a future change
 * accidentally accumulates row data, sst entries, or buffers.
 */
class MemoryFootprintTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-mem-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_streaming_read_uses_bounded_memory_delta(): void
    {
        if (! function_exists('memory_reset_peak_usage')) {
            $this->markTestSkipped('memory_reset_peak_usage() requires PHP 8.3+');
        }

        $w = new SinkableXlsxWriter(new FileSink($this->testFile));
        $w->startFile(['id', 'value']);
        for ($i = 1; $i <= 100_000; $i++) {
            $w->writeRow([$i, "row-{$i}"]);
        }
        $w->finishFile();

        gc_collect_cycles();
        memory_reset_peak_usage();
        $baseline = memory_get_usage(true);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        foreach ($reader->rows() as $_) {
            // discard
        }

        $delta = memory_get_peak_usage(true) - $baseline;

        $this->assertLessThan(
            4 * 1024 * 1024,
            $delta,
            "reader RAM delta {$delta} bytes exceeded 4 MB envelope"
        );
    }
}

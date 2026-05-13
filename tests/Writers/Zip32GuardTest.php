<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Pins the writer's ZIP32 limit guards. Each test injects an
 * implausibly large value into the writer's internal counters via
 * reflection, then drives the next write — the guard must abort with
 * a clear exception instead of silently truncating size fields and
 * shipping a corrupt archive.
 *
 * Real workloads never hit these limits (4 GB / 65,535 entries), but
 * accidental loops or upstream injection bugs can. Loud rejection
 * preserves the writer's correctness contract.
 */
class Zip32GuardTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-zip32-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_writer_rejects_archive_offset_exceeding_4gb(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id']);
        $writer->writeRow([1]);

        // Force the cumulative offset just past the ZIP32 ceiling. The
        // next central-directory write must trip the guard before any
        // size field gets truncated to 32 bits.
        $reflection = new \ReflectionClass($writer);
        $offsetProp = $reflection->getProperty('currentOffset');
        $offsetProp->setValue($writer, 0xFFFFFFFF + 1);

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/ZIP32 limit exceeded.*4 GB/');

        $writer->finishFile();
    }

    public function test_writer_rejects_entry_count_exceeding_65535(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id']);
        $writer->writeRow([1]);

        // Pretend the central directory already holds 70k entries.
        // writeCentralDirectory's count check should fire before any
        // EOCD bytes are emitted.
        $reflection = new \ReflectionClass($writer);
        $cdProp = $reflection->getProperty('centralDirectory');
        $cdProp->setValue($writer, array_fill(0, 70_000, [
            'filename' => 'fake.xml',
            'crc32' => 0,
            'compressed_size' => 1,
            'uncompressed_size' => 1,
            'offset' => 0,
            'compression' => 8,
            'flags' => 0,
            'timestamp' => time(),
        ]));

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessageMatches('/exceeding the 65535 ZIP32 limit/');

        $writer->finishFile();
    }
}

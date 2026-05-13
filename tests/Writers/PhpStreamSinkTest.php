<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\PhpStreamSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * PhpStreamSink — write-direct-to-stream sink covering the HTTP
 * response-streaming use case.
 *
 * Pinned behaviours:
 *   - Produces a valid ZIP (PK magic + workbook parts) when fed an
 *     XLSX byte stream
 *   - output() factory captures via ob_start
 *   - Invalid resource fails fast
 *   - Caller-owned streams stay open after sink close
 */
class PhpStreamSinkTest extends TestCase
{
    public function test_writes_valid_xlsx_to_php_memory(): void
    {
        $h = fopen('php://memory', 'w+b');
        $sink = new PhpStreamSink($h);
        $writer = new SinkableXlsxWriter($sink);
        $writer->startFile(['id', 'name']);
        $writer->writeRow([1, 'Alice']);
        $writer->writeRow([2, 'Bob']);
        $writer->finishFile();

        rewind($h);
        $bytes = stream_get_contents($h);
        fclose($h);

        $this->assertNotEmpty($bytes);
        $this->assertSame('PK', substr($bytes, 0, 2), 'output must start with ZIP local file header magic');
        // EOCD signature must appear somewhere in the trailing 64KB
        $this->assertStringContainsString("PK\x05\x06", $bytes);
    }

    public function test_output_factory_streams_to_php_output(): void
    {
        // Capture php://output through ob_start so the test can inspect bytes.
        ob_start();

        $sink = PhpStreamSink::output();
        $writer = new SinkableXlsxWriter($sink);
        $writer->startFile(['x']);
        $writer->writeRow(['y']);
        $writer->finishFile();

        $captured = ob_get_clean();

        $this->assertNotEmpty($captured);
        $this->assertSame('PK', substr($captured, 0, 2));
        $this->assertGreaterThan(0, $sink->getBytesWritten());
    }

    public function test_temp_factory_returns_writable_sink(): void
    {
        $sink = PhpStreamSink::temp();
        $sink->write('hello');
        $this->assertSame(5, $sink->getBytesWritten());

        $stream = $sink->getStream();
        rewind($stream);
        $this->assertSame('hello', stream_get_contents($stream));
    }

    public function test_memory_factory_returns_writable_sink(): void
    {
        $sink = PhpStreamSink::memory();
        $sink->write('abc');
        $sink->write('def');
        $this->assertSame(6, $sink->getBytesWritten());
    }

    public function test_constructor_rejects_non_resource(): void
    {
        $this->expectException(XlsxStreamException::class);
        new PhpStreamSink('not a stream'); // @phpstan-ignore-line
    }

    public function test_write_after_close_throws(): void
    {
        $sink = PhpStreamSink::memory();
        $sink->close();

        $this->expectException(XlsxStreamException::class);
        $sink->write('x');
    }

    public function test_close_is_idempotent(): void
    {
        $sink = PhpStreamSink::memory();
        $sink->close();
        $sink->close(); // must not throw
        $this->assertTrue(true);
    }

    public function test_caller_owned_stream_stays_open_after_close(): void
    {
        $h = fopen('php://memory', 'w+b');
        $sink = new PhpStreamSink($h, ownsStream: false);
        $sink->write('payload');
        $sink->close();

        // Caller still owns the resource — should be readable.
        $this->assertTrue(is_resource($h));
        rewind($h);
        $this->assertSame('payload', stream_get_contents($h));
        fclose($h);
    }

    public function test_owned_stream_is_closed_on_close(): void
    {
        $sink = PhpStreamSink::memory();
        $stream = $sink->getStream();
        $sink->close();

        $this->assertFalse(is_resource($stream));
    }

    public function test_abort_drops_partial_data_on_owned_stream(): void
    {
        $sink = PhpStreamSink::memory();
        $sink->write('partial');
        $sink->abort();

        // Owned stream is closed without flushing — same shape as FileSink::abort
        $this->assertFalse(is_resource($sink->getStream()));
    }
}

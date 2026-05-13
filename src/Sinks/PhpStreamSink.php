<?php

namespace Kolay\XlsxStream\Sinks;

use Kolay\XlsxStream\Contracts\Sink;
use Kolay\XlsxStream\Exceptions\XlsxStreamException;

/**
 * Sink that writes directly to a PHP stream resource — `php://output`,
 * `php://memory`, `php://temp`, or any caller-owned stream.
 *
 * Primary use case: stream a workbook into an HTTP response without ever
 * materialising it on disk. Pairs naturally with Laravel's
 * `Response::stream()`:
 *
 *     return response()->stream(function () {
 *         $writer = new SinkableXlsxWriter(PhpStreamSink::output());
 *         $writer->startFile(['id', 'name']);
 *         User::query()->lazy()->each(fn ($u) => $writer->writeRow([$u->id, $u->name]));
 *         $writer->finishFile();
 *     }, 200, [
 *         'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
 *         'Content-Disposition' => 'attachment; filename="users.xlsx"',
 *     ]);
 *
 * The sink owns the stream when constructed via `output()` / `temp()` /
 * `memory()` — those factories `fopen()` the stream and `fclose()` it on
 * close. When constructed from a caller-supplied resource, the caller
 * retains ownership and is responsible for closing.
 */
class PhpStreamSink implements Sink
{
    /** @var resource */
    private $stream;

    private bool $ownsStream;

    private int $bytesWritten = 0;

    private bool $closed = false;

    /**
     * @param  resource  $stream
     */
    public function __construct($stream, bool $ownsStream = false)
    {
        if (! is_resource($stream)) {
            throw new XlsxStreamException('PhpStreamSink requires a valid stream resource');
        }
        $this->stream = $stream;
        $this->ownsStream = $ownsStream;
    }

    /**
     * Open `php://output` and stream into the active HTTP response.
     */
    public static function output(): self
    {
        $h = fopen('php://output', 'wb');
        if ($h === false) {
            throw new XlsxStreamException('Could not open php://output');
        }

        return new self($h, ownsStream: true);
    }

    /**
     * Open a `php://temp` buffer (in-memory until 2 MB, then a tmp file).
     * Useful for capturing a workbook for later inspection.
     */
    public static function temp(): self
    {
        $h = fopen('php://temp', 'w+b');
        if ($h === false) {
            throw new XlsxStreamException('Could not open php://temp');
        }

        return new self($h, ownsStream: true);
    }

    /**
     * Open a `php://memory` buffer. Smaller than `temp()` — never spills
     * to disk — so suitable only when the caller knows the workbook is
     * small (tests, fixtures).
     */
    public static function memory(): self
    {
        $h = fopen('php://memory', 'w+b');
        if ($h === false) {
            throw new XlsxStreamException('Could not open php://memory');
        }

        return new self($h, ownsStream: true);
    }

    public function write(string $data): void
    {
        if ($this->closed) {
            throw new XlsxStreamException('Cannot write to closed PhpStreamSink');
        }

        $written = fwrite($this->stream, $data);
        if ($written === false || $written !== strlen($data)) {
            throw new XlsxStreamException('PhpStreamSink: short write to stream');
        }
        $this->bytesWritten += $written;
    }

    public function close(): void
    {
        if ($this->closed) {
            return;
        }

        if (is_resource($this->stream)) {
            fflush($this->stream);
            if ($this->ownsStream) {
                fclose($this->stream);
            }
        }
        $this->closed = true;
    }

    public function abort(): void
    {
        if ($this->closed) {
            return;
        }

        // Owned streams are closed without flush so partial data is dropped;
        // caller-owned streams are left untouched (the caller decides).
        if ($this->ownsStream && is_resource($this->stream)) {
            fclose($this->stream);
        }
        $this->closed = true;
    }

    public function getBytesWritten(): int
    {
        return $this->bytesWritten;
    }

    /**
     * @return resource
     */
    public function getStream()
    {
        return $this->stream;
    }

    public function __destruct()
    {
        if (! $this->closed && $this->ownsStream && is_resource($this->stream)) {
            try {
                fclose($this->stream);
            } catch (\Throwable) {
                // Destructors must not propagate errors.
            }
        }
    }
}

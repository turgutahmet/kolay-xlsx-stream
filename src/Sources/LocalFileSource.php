<?php

namespace Kolay\XlsxStream\Sources;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * Source backed by a local filesystem path.
 *
 * Holds one persistent handle for random-access range() reads and opens
 * a fresh handle per streamFrom() call so the parser's long-running
 * sequential read does not interfere with central-directory lookups.
 */
class LocalFileSource implements Source
{
    /** @var resource */
    private $handle;

    private string $path;
    private int $size;
    private bool $closed = false;

    public function __construct(string $path)
    {
        $h = @fopen($path, 'rb');
        if ($h === false) {
            throw XlsxReadException::sourceUnreadable("cannot open file: {$path}");
        }

        $this->handle = $h;
        $this->path = $path;
        $this->size = (int) filesize($path);
    }

    public function size(): int
    {
        return $this->size;
    }

    public function range(int $offset, int $length): string
    {
        $this->guardOpen();

        if (fseek($this->handle, $offset) !== 0) {
            throw XlsxReadException::sourceUnreadable("fseek failed at offset {$offset}");
        }

        $buf = '';
        $remaining = $length;

        while ($remaining > 0) {
            $chunk = fread($this->handle, min(65536, $remaining));
            if ($chunk === false || $chunk === '') {
                break;
            }
            $buf .= $chunk;
            $remaining -= strlen($chunk);
        }

        return $buf;
    }

    public function streamFrom(int $offset)
    {
        $this->guardOpen();

        $h = @fopen($this->path, 'rb');
        if ($h === false) {
            throw XlsxReadException::sourceUnreadable("cannot reopen file: {$this->path}");
        }

        if ($offset > 0 && fseek($h, $offset) !== 0) {
            fclose($h);
            throw XlsxReadException::sourceUnreadable("fseek failed at offset {$offset}");
        }

        return $h;
    }

    public function close(): void
    {
        if ($this->closed) {
            return;
        }
        if (is_resource($this->handle)) {
            fclose($this->handle);
        }
        $this->closed = true;
    }

    public function __destruct()
    {
        $this->close();
    }

    private function guardOpen(): void
    {
        if ($this->closed) {
            throw XlsxReadException::sourceUnreadable('source has been closed');
        }
    }
}

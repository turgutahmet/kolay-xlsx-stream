<?php

namespace Kolay\XlsxStream\Sources;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Contracts\SupportsSuffixRange;
use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * Source backed by a local filesystem path.
 *
 * Holds one persistent handle for random-access range() reads and opens
 * a fresh handle per streamFrom() call so the parser's long-running
 * sequential read does not interfere with central-directory lookups.
 */
class LocalFileSource implements Source, SupportsSuffixRange
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

    /**
     * Suffix read — trivial locally: the size is already known from the
     * constructor's stat, so this is just a clamped range() plus the
     * cached size. Exists so ZipDirectory can use one code path for
     * every SupportsSuffixRange source.
     *
     * @return array{data: string, size: int}
     */
    public function tail(int $length): array
    {
        $len = max(0, min($length, $this->size));

        return [
            'data' => $this->range($this->size - $len, $len),
            'size' => $this->size,
        ];
    }

    public function streamFrom(int $offset, ?int $length = null)
    {
        $this->guardOpen();

        // $length is ignored: local fread is lazy and the sheet reader
        // already caps how many compressed bytes it pulls (bounded scans
        // stop at a sync boundary via inflatedChunks' compLength), so a
        // plain seeked handle over-reads nothing. Bounding only pays on
        // remote sources where the range is an HTTP fetch size.
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

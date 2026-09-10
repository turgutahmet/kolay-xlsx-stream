<?php

namespace Kolay\XlsxStream\Sources;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Contracts\SupportsBoundedStream;
use Kolay\XlsxStream\Contracts\SupportsSuffixRange;
use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * Source backed by an in-memory byte string.
 *
 * The mirror of LocalFileSource for bytes that never reach the disk — a
 * template rendered inside the same process, a payload pulled from cache
 * or a queue message. Every read is a substr, so `range()` and `tail()`
 * cost nothing and the ZIP central-directory bootstrap needs no I/O at all.
 *
 * Sized for layouts and small payloads: `streamFrom()` materialises the
 * requested slice into a fresh `php://memory` handle (the Source contract
 * hands ownership of each stream to the caller, so handles cannot be
 * shared). Point a multi-hundred-megabyte input at LocalFileSource or
 * S3RangeSource instead — those stream lazily; this one is already
 * entirely resident by construction.
 */
class StringSource implements Source, SupportsBoundedStream, SupportsSuffixRange
{
    private string $data;
    private int $size;
    private bool $closed = false;

    public function __construct(string $data)
    {
        $this->data = $data;
        $this->size = \strlen($data);
    }

    public function size(): int
    {
        return $this->size;
    }

    public function range(int $offset, int $length): string
    {
        $this->guardOpen();

        if ($length <= 0 || $offset >= $this->size) {
            return '';
        }

        return substr($this->data, max(0, $offset), $length);
    }

    /**
     * Suffix read — the size is known by construction, so this is a clamped
     * substr plus that size, giving ZipDirectory its one-call bootstrap.
     *
     * @return array{data: string, size: int}
     */
    public function tail(int $length): array
    {
        $this->guardOpen();

        $len = max(0, min($length, $this->size));

        return [
            'data' => $len === 0 ? '' : substr($this->data, $this->size - $len, $len),
            'size' => $this->size,
        ];
    }

    public function streamFrom(int $offset)
    {
        return $this->openStream($offset, null);
    }

    public function streamFromRange(int $offset, int $length)
    {
        return $this->openStream($offset, $length);
    }

    public function close(): void
    {
        if ($this->closed) {
            return;
        }
        // Drop the payload so a long-lived writer holding a closed template
        // stops pinning its bytes; $size stays queryable.
        $this->data = '';
        $this->closed = true;
    }

    /**
     * Open a fresh forward-only handle over one slice. A new handle per
     * call is required by the contract (the caller fcloses it), so the
     * slice is copied rather than shared.
     *
     * @return resource
     */
    private function openStream(int $offset, ?int $length)
    {
        $this->guardOpen();

        $slice = $length === null
            ? $this->range($offset, max(0, $this->size - max(0, $offset)))
            : $this->range($offset, $length);

        $handle = fopen('php://memory', 'r+b');
        if ($handle === false) {
            throw XlsxReadException::sourceUnreadable('cannot open an in-memory stream');
        }
        if ($slice !== '') {
            fwrite($handle, $slice);
            rewind($handle);
        }

        return $handle;
    }

    private function guardOpen(): void
    {
        if ($this->closed) {
            throw XlsxReadException::sourceUnreadable('source has been closed');
        }
    }
}

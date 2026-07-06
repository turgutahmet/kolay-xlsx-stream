<?php

namespace Kolay\XlsxStream\Contracts;

/**
 * Optional capability: a Source that can open a stream bounded to a byte
 * length — cheaply serving just [offset, offset+length-1] rather than
 * reading to the end of the object.
 *
 * The reader uses this for pruned scans that stop at a ZLIB_FULL_FLUSH
 * sync boundary: on a remote source (S3 range GET) it turns an open-ended
 * [offset, EOF] fetch into an exact ranged read, so a query over an early
 * block doesn't pull the whole file off the wire. It is purely an I/O
 * optimization — a Source that does NOT implement this still works
 * correctly (the reader falls back to streamFrom() and caps the read
 * itself), so this stays out of the base Source contract to keep that
 * contract stable for third-party implementations.
 */
interface SupportsBoundedStream
{
    /**
     * Like Source::streamFrom($offset), but the returned stream serves at
     * most $length bytes from $offset. The bound is an optimization, not a
     * correctness contract on the exact byte count — an implementation may
     * over-read (e.g. to EOF) when bounding is not cheap.
     *
     * @return resource
     */
    public function streamFromRange(int $offset, int $length);
}

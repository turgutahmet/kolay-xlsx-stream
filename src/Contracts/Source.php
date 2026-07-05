<?php

namespace Kolay\XlsxStream\Contracts;

/**
 * Source interface for streaming input abstraction.
 *
 * Symmetric counterpart to Sink. Allows the reader to consume from
 * different origins (local file, S3 range GET, HTTP range, …) without
 * the parser knowing which.
 *
 * Implementations MUST support concurrent random-access reads via range();
 * streamFrom() is for the long-running sequential sheet inflate path and
 * may share or open a separate underlying handle.
 */
interface Source
{
    /**
     * Total size of the source in bytes. May trigger a HEAD or stat() on
     * first call; result MUST be cached for subsequent calls.
     */
    public function size(): int;

    /**
     * Synchronously fetch a contiguous byte range. Used for small reads
     * (EOCD tail, central directory, local file headers).
     */
    public function range(int $offset, int $length): string;

    /**
     * Open a forward-only stream resource starting at the given offset.
     * Caller is responsible for fclose(). The stream MUST support fread()
     * and feof(); seeking is not required.
     *
     * When $length is given the stream is bounded to at most that many
     * bytes from $offset (a ranged [offset, offset+length-1] fetch on
     * remote sources; a read cap on local ones). $length === null keeps
     * the historical behaviour of reading to the end of the object. The
     * bound is an I/O optimization for pruned scans that stop at a
     * ZLIB_FULL_FLUSH sync boundary — never a correctness contract on
     * exact byte counts, so implementations may over-read to EOF if
     * bounding is not cheap.
     *
     * @return resource
     */
    public function streamFrom(int $offset, ?int $length = null);

    /**
     * Release any underlying handles. Idempotent — safe to call multiple times.
     */
    public function close(): void;
}

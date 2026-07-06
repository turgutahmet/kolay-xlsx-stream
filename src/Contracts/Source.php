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
     * Open a forward-only stream resource starting at the given offset,
     * reading to the end of the object. Caller is responsible for
     * fclose(). The stream MUST support fread() and feof(); seeking is not
     * required.
     *
     * A source that can serve a BOUNDED range cheaply (stopping after a
     * given byte length, e.g. an HTTP Range GET) additionally implements
     * SupportsBoundedStream — the reader uses that for pruned scans that
     * end at a sync boundary. Sources that only implement this method keep
     * working; they just fetch to EOF and let the reader cap the read.
     *
     * @return resource
     */
    public function streamFrom(int $offset);

    /**
     * Release any underlying handles. Idempotent — safe to call multiple times.
     */
    public function close(): void;
}

<?php

namespace Kolay\XlsxStream\Contracts;

/**
 * Optional Source capability: fetch the trailing N bytes of the source
 * and learn its total size in a single operation.
 *
 * The ZIP central-directory bootstrap needs exactly this shape — "give
 * me the last 64 KB" — and HTTP range semantics support it natively
 * (`Range: bytes=-N`), which collapses the HEAD(size) + ranged-tail
 * request pair into one round-trip on S3.
 *
 * Kept as a separate interface rather than a new method on Source
 * because Source is a published contract: adding an abstract method
 * would break every external implementor on a minor release. Sources
 * that don't implement it transparently keep the two-step flow.
 */
interface SupportsSuffixRange
{
    /**
     * Return the last min($length, size) bytes of the source together
     * with the total source size in bytes. Implementations MUST handle
     * $length exceeding the source size by returning the whole content,
     * and SHOULD cache the discovered size so a later size() call does
     * not pay another round-trip.
     *
     * @return array{data: string, size: int}
     */
    public function tail(int $length): array;
}

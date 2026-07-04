<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * In-memory shared-strings lookup. Backed by a list<string> indexed
 * positionally.
 *
 * As of v3.2 the reader no longer constructs this class —
 * SharedStringsParser builds a PackedSharedStrings instead, whose flat
 * payload+offset buffers avoid the ~57 bytes/entry PHP-array overhead
 * measured on large tables. This implementation is kept for callers
 * (and tests) that hold a pre-resolved list<string> and want the
 * simplest possible SharedStrings to hand to a StreamingSheetReader.
 */
class InMemorySharedStrings implements SharedStrings
{
    /** @var list<string> */
    private array $strings;

    /**
     * @param  list<string>  $strings
     */
    public function __construct(array $strings)
    {
        $this->strings = $strings;
    }

    public function get(int $index): string
    {
        if (! isset($this->strings[$index])) {
            throw XlsxReadException::corruptCentralDirectory(
                "shared-string index {$index} out of range (table has ".count($this->strings).' entries)'
            );
        }

        return $this->strings[$index];
    }

    public function count(): int
    {
        return count($this->strings);
    }
}

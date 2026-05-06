<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * In-memory shared-strings lookup. Backed by a list<string> indexed
 * positionally. Construction is via SharedStringsParser::parseInMemory()
 * for production use; direct instantiation is for tests and
 * pre-resolved tables.
 *
 * Used when the archive's xl/sharedStrings.xml stays under the size
 * threshold the reader applies (default 20 MB compressed). Beyond that
 * threshold the reader refuses to load the table; an on-disk index
 * variant is tracked as a future addition.
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

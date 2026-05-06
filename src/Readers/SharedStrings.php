<?php

namespace Kolay\XlsxStream\Readers;

/**
 * Lookup contract for the shared-strings table referenced by t="s" cells.
 *
 * Implementations live in this namespace; the reader picks one based on
 * the size of xl/sharedStrings.xml in the archive. Files written by
 * SinkableXlsxWriter never use t="s" so the contract is never invoked
 * for self-written archives — it exists for compatibility with XLSX
 * files produced by external writers (PhpSpreadsheet, openpyxl, Apache
 * POI, etc.) which deduplicate strings into the shared table.
 */
interface SharedStrings
{
    /**
     * Resolve a 0-based shared-string index to its decoded UTF-8 value.
     */
    public function get(int $index): string;

    /**
     * Number of strings in the table.
     */
    public function count(): int;
}

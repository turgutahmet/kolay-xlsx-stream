<?php

namespace Kolay\XlsxStream\Readers;

/**
 * Lookup contract for the shared-strings table referenced by t="s" cells.
 *
 * Implementations live in this namespace; the reader builds a
 * PackedSharedStrings (flat payload + offset index), while
 * InMemorySharedStrings remains for pre-resolved list<string> tables.
 * Files written by SinkableXlsxWriter never use t="s" so the contract
 * is never invoked for self-written archives — it exists for
 * compatibility with XLSX files produced by external writers
 * (PhpSpreadsheet, openpyxl, Apache POI, etc.) which deduplicate
 * strings into the shared table.
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

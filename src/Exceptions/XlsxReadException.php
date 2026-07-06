<?php

namespace Kolay\XlsxStream\Exceptions;

/**
 * Reader-specific errors. Extends XlsxStreamException so callers that already
 * catch the base class continue to work unchanged.
 */
class XlsxReadException extends XlsxStreamException
{
    public static function sourceUnreadable(string $reason): self
    {
        return new self("Cannot read source: {$reason}");
    }

    public static function eocdNotFound(): self
    {
        return new self('ZIP End of Central Directory record not found — file is not a valid XLSX/ZIP archive.');
    }

    public static function corruptCentralDirectory(string $reason): self
    {
        return new self("ZIP Central Directory is corrupt: {$reason}");
    }

    public static function entryNotFound(string $name): self
    {
        return new self("ZIP entry not found in archive: {$name}");
    }

    public static function badLocalFileHeader(string $name): self
    {
        return new self("Local file header signature mismatch for entry: {$name}");
    }

    public static function inflateFailed(string $reason): self
    {
        return new self("DEFLATE inflate failed: {$reason}");
    }

    public static function zip64NotSupported(): self
    {
        return new self(
            'ZIP64 archives (>4 GB or >65535 entries) are not yet supported by the reader. '.
            'Tracked for a future v3.x release.'
        );
    }

    /**
     * A header name matched more than one column — silently picking one
     * would be the silent-wrong-answer failure class, so the caller must
     * disambiguate (rename the header or address the column by number).
     *
     * @param  list<int>  $positions  1-based column positions carrying the name
     */
    public static function ambiguousColumnName(string $name, array $positions): self
    {
        return new self(sprintf(
            "Column name '%s' is ambiguous — it appears at positions %s. ".
            'Address the column by number, or give the headers distinct names.',
            $name,
            implode(' and ', $positions)
        ));
    }

    /**
     * A header name resolved to nothing. The message carries the sheet's
     * actual headers so a typo is a one-glance fix, not a debugging trip.
     *
     * @param  list<string>  $headers  the sheet's addressable header names
     */
    public static function unknownColumnName(string $name, array $headers): self
    {
        $shown = array_slice($headers, 0, 20);
        $list = $shown === [] ? '(sheet has no addressable headers)' : "'".implode("', '", $shown)."'";
        if (count($headers) > 20) {
            $list .= sprintf(' … +%d more', count($headers) - 20);
        }

        return new self(sprintf(
            "Unknown column name '%s'. Available headers: %s.",
            $name,
            $list
        ));
    }
}

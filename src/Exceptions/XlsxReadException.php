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
}

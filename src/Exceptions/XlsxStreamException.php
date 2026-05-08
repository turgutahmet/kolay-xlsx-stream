<?php

namespace Kolay\XlsxStream\Exceptions;

/**
 * Base exception for all XLSX Stream related errors
 */
class XlsxStreamException extends \Exception
{
    /**
     * Create exception for sink write failure
     */
    public static function sinkWriteFailed(string $reason): self
    {
        return new self("Failed to write to sink: {$reason}");
    }

    /**
     * Create exception for sink close failure
     */
    public static function sinkCloseFailed(string $reason): self
    {
        return new self("Failed to close sink: {$reason}");
    }

    /**
     * Create exception for invalid compression level
     */
    public static function invalidCompressionLevel(int $level): self
    {
        return new self("Invalid compression level: {$level}. Must be between 1 and 9.");
    }

    /**
     * Create exception for invalid buffer size
     */
    public static function invalidBufferSize(int $size): self
    {
        return new self("Invalid buffer size: {$size}. Must be at least 1.");
    }

    /**
     * Create exception for writer already closed
     */
    public static function writerAlreadyClosed(): self
    {
        return new self('Cannot perform operation on closed writer.');
    }

    /**
     * Create exception for headers not set
     */
    public static function headersNotSet(): self
    {
        return new self('Headers must be set before writing rows. Call startFile() first.');
    }

    /**
     * Create exception for startFile called more than once
     */
    public static function alreadyStarted(): self
    {
        return new self('Writer has already been started. startFile() can only be called once.');
    }

    /**
     * Create exception for column count exceeding Excel's limit
     */
    public static function tooManyColumns(int $given, int $max): self
    {
        return new self("Column count {$given} exceeds Excel's maximum of {$max} columns per sheet.");
    }

    /**
     * Create exception for finalize-with-no-data
     */
    public static function emptyWorkbook(): self
    {
        return new self(
            'Cannot finalize an empty workbook. Write at least one row via writeRow() '.
            'or call newSheet() to create a sheet before finishFile().'
        );
    }

    /**
     * Create exception for setColumnFormat targeting a column past the header count.
     */
    public static function columnIndexOutOfRange(int $given, int $max): self
    {
        return new self(
            "Column index {$given} is out of range — the current header has {$max} columns. ".
            'Call setColumnFormat() with an index between 1 and the header count.'
        );
    }

    /**
     * Raised when an export approaches a 32-bit ZIP container limit.
     * Loud rejection beats silently truncating size fields and shipping
     * a corrupt archive — split the export across multiple files until
     * ZIP64 writer support lands.
     */
    public static function zip32LimitExceeded(string $detail): self
    {
        return new self(
            "ZIP32 limit exceeded: {$detail}. ".
            'Split the export across multiple files or sheets as a workaround. '.
            'ZIP64 writer support is tracked for a future release.'
        );
    }
}

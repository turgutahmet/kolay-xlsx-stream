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
        return new self("Cannot perform operation on closed writer.");
    }

    /**
     * Create exception for headers not set
     */
    public static function headersNotSet(): self
    {
        return new self("Headers must be set before writing rows. Call startFile() first.");
    }
}
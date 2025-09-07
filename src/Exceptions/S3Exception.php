<?php

namespace Kolay\XlsxStream\Exceptions;

/**
 * S3-specific exceptions
 */
class S3Exception extends XlsxStreamException
{
    /**
     * Create exception for multipart upload initialization failure
     */
    public static function multipartInitFailed(string $bucket, string $key, string $reason): self
    {
        return new self("Failed to initialize multipart upload for s3://{$bucket}/{$key}: {$reason}");
    }

    /**
     * Create exception for part upload failure
     */
    public static function partUploadFailed(int $partNumber, string $reason): self
    {
        return new self("Failed to upload part {$partNumber}: {$reason}");
    }

    /**
     * Create exception for multipart completion failure
     */
    public static function multipartCompleteFailed(string $reason): self
    {
        return new self("Failed to complete multipart upload: {$reason}");
    }

    /**
     * Create exception for multipart abort failure
     */
    public static function multipartAbortFailed(string $reason): self
    {
        return new self("Failed to abort multipart upload: {$reason}");
    }

    /**
     * Create exception for invalid part size
     */
    public static function invalidPartSize(int $size): self
    {
        return new self("Invalid S3 part size: {$size} bytes. Minimum is 5MB (5242880 bytes).");
    }
}
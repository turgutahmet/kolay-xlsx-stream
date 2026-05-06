<?php

namespace Kolay\XlsxStream\Sources;

use Aws\S3\S3Client;
use GuzzleHttp\Psr7\StreamWrapper;
use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * Source backed by an S3 object via HTTP Range GET.
 *
 * range() reads return materialized bytes; streamFrom() returns a php
 * stream wrapped over the AWS SDK's lazy PSR-7 body, so the underlying
 * HTTP body is not buffered into memory or php://temp before consumption.
 *
 * Lazy-stream behaviour requires '@http' => ['stream' => true] on the
 * getObject call — without it the SDK spills bodies >2 MB to php://temp,
 * which would defeat the bounded-memory goal of the reader.
 */
class S3RangeSource implements Source
{
    private S3Client $s3;
    private string $bucket;
    private string $key;
    private ?int $size = null;

    public function __construct(S3Client $s3, string $bucket, string $key)
    {
        $this->s3 = $s3;
        $this->bucket = $bucket;
        $this->key = $key;
    }

    public function size(): int
    {
        if ($this->size !== null) {
            return $this->size;
        }

        try {
            $r = $this->s3->headObject([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
            ]);
        } catch (\Throwable $e) {
            throw XlsxReadException::sourceUnreadable(
                "HEAD s3://{$this->bucket}/{$this->key}: ".$e->getMessage()
            );
        }

        return $this->size = (int) $r['ContentLength'];
    }

    public function range(int $offset, int $length): string
    {
        try {
            $r = $this->s3->getObject([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
                'Range' => sprintf('bytes=%d-%d', $offset, $offset + $length - 1),
            ]);
        } catch (\Throwable $e) {
            throw XlsxReadException::sourceUnreadable(
                "GET range s3://{$this->bucket}/{$this->key}: ".$e->getMessage()
            );
        }

        return (string) $r['Body'];
    }

    public function streamFrom(int $offset)
    {
        $size = $this->size();

        try {
            $r = $this->s3->getObject([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
                'Range' => sprintf('bytes=%d-%d', $offset, $size - 1),
                '@http' => ['stream' => true],
            ]);
        } catch (\Throwable $e) {
            throw XlsxReadException::sourceUnreadable(
                "GET stream s3://{$this->bucket}/{$this->key}: ".$e->getMessage()
            );
        }

        return StreamWrapper::getResource($r['Body']);
    }

    public function close(): void
    {
        // S3 source is stateless — nothing to release.
    }
}

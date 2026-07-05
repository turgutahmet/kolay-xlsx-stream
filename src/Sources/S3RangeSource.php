<?php

namespace Kolay\XlsxStream\Sources;

use Aws\S3\S3Client;
use GuzzleHttp\Psr7\StreamWrapper;
use Kolay\XlsxStream\Contracts\ProvidesCostHints;
use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Contracts\SupportsSuffixRange;
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
class S3RangeSource implements ProvidesCostHints, Source, SupportsSuffixRange
{
    private S3Client $s3;
    private string $bucket;
    private string $key;
    private ?int $size = null;
    private int $rttUs;
    private int $bandwidthBps;

    /**
     * $rttUs / $bandwidthBps are gap-bridging planning hints (see
     * ProvidesCostHints), not benchmark claims — override them for your
     * deployment. Defaults are conservative intra-region estimates
     * (~30 ms round-trip, ~50 MB/s per connection); they only affect how
     * aggressively pruned scans coalesce runs, never correctness.
     */
    public function __construct(
        S3Client $s3,
        string $bucket,
        string $key,
        int $rttUs = 30_000,
        int $bandwidthBps = 52_428_800
    ) {
        $this->s3 = $s3;
        $this->bucket = $bucket;
        $this->key = $key;
        $this->rttUs = $rttUs;
        $this->bandwidthBps = $bandwidthBps;
    }

    public function costHints(): array
    {
        return ['rtt_us' => $this->rttUs, 'bandwidth_bps' => $this->bandwidthBps];
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
        // A zero/negative length would render as `bytes=X-(X-1)` — an
        // unsatisfiable range S3 rejects with 416. An empty read has an
        // empty result; skip the request.
        if ($length <= 0) {
            return '';
        }

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

    /**
     * Suffix-range read: `Range: bytes=-N` returns the last N bytes in
     * one GET, and the 206 response's `Content-Range: bytes X-Y/total`
     * header carries the object size — so this replaces the HEAD +
     * ranged-tail pair ZipDirectory used to need with a single request.
     *
     * @return array{data: string, size: int}
     */
    public function tail(int $length): array
    {
        // `bytes=-0` is an unsatisfiable range (416); an empty suffix
        // has an empty result, so only the size needs resolving.
        if ($length < 1) {
            return ['data' => '', 'size' => $this->size()];
        }

        try {
            $r = $this->s3->getObject([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
                'Range' => sprintf('bytes=-%d', $length),
            ]);
        } catch (\Throwable $e) {
            throw XlsxReadException::sourceUnreadable(
                "GET suffix range s3://{$this->bucket}/{$this->key}: ".$e->getMessage()
            );
        }

        $data = (string) $r['Body'];

        // AWS answers a suffix range with 206 + Content-Range (also when
        // the suffix covers the whole object). Some S3-compatible stores
        // reply 200 without Content-Range in the covers-everything case —
        // the body then IS the whole object, so its length is the size.
        $size = strlen($data);
        if (isset($r['ContentRange']) && preg_match('~/(\d+)$~', (string) $r['ContentRange'], $m)) {
            $size = (int) $m[1];
        }

        // Cache so a later size() call (e.g. streamFrom) skips the HEAD.
        $this->size = $size;

        return ['data' => $data, 'size' => $size];
    }

    public function streamFrom(int $offset, ?int $length = null)
    {
        $size = $this->size();

        // Bounded scans stop at a sync boundary: fetch exactly
        // [offset, offset+length-1] instead of streaming to EOF, so a
        // pruned query over an early block reads only that block's bytes
        // off the wire. Clamp to the last byte so a length past EOF (or
        // null) degrades to the historical open-ended range.
        $end = $length !== null ? min($offset + $length - 1, $size - 1) : $size - 1;

        try {
            $r = $this->s3->getObject([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
                'Range' => sprintf('bytes=%d-%d', $offset, $end),
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

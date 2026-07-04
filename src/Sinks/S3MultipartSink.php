<?php

namespace Kolay\XlsxStream\Sinks;

use Aws\Exception\AwsException;
use Aws\S3\S3Client;
use GuzzleHttp\Promise\CancellationException;
use GuzzleHttp\Promise\PromiseInterface;
use GuzzleHttp\Promise\Utils;
use Kolay\XlsxStream\Contracts\Sink;
use Kolay\XlsxStream\Exceptions\S3Exception;

/**
 * S3 Multipart Upload Sink
 *
 * Streams data directly to S3 using multipart upload
 * - Zero disk usage
 * - Bounded memory: ~partSize x (concurrency + 1)
 * - Automatic part management
 * - Parallel part uploads via a bounded async window
 * - Error recovery with abort
 */
class S3MultipartSink implements Sink
{
    private S3Client $s3;

    private string $bucket;

    private string $key;

    private string $uploadId;

    /** @var array<int, array{PartNumber: int, ETag: string}> keyed by part number */
    private array $parts = [];

    /**
     * In-flight part uploads, keyed by part number in dispatch order
     * (insertion order == FIFO order for window waits).
     *
     * @var array<int, PromiseInterface>
     */
    private array $inFlight = [];

    /**
     * First upload failure (first-error-wins). Checked at the top of
     * write()/close(); once set the sink aborts and throws.
     */
    private ?\Throwable $failed = null;

    private string $buffer = '';

    private int $partSize;

    private int $concurrency;

    private int $partNumber = 1;

    private int $bytesWritten = 0;

    private array $putObjectParams;

    private bool $closed = false;

    // S3 minimum part size is 5MB (except last part)
    public const MIN_PART_SIZE = 5242880; // 5MB

    public const DEFAULT_PART_SIZE = 8388608; // 8MB

    public const DEFAULT_CONCURRENCY = 4;

    /**
     * @param S3Client $s3
     * @param string $bucket
     * @param string $key S3 object key (path)
     * @param int $partSize Size of each part (min 5MB)
     * @param array $putObjectParams Additional S3 parameters (ContentType, ACL, etc)
     * @param int|null $concurrency Max part uploads in flight at once (default 4).
     *        Parts are dispatched asynchronously; when the window is full the
     *        writer waits for the oldest part before dispatching the next one.
     *        Memory ceiling is roughly partSize x (concurrency + 1): up to
     *        `concurrency` part bodies held by in-flight requests plus the
     *        accumulating buffer. Use 1 for strictly sequential uploads
     *        (same S3 request sequence as pre-3.2 versions).
     */
    public function __construct(
        S3Client $s3,
        string $bucket,
        string $key,
        int $partSize = self::DEFAULT_PART_SIZE,
        array $putObjectParams = [],
        ?int $concurrency = self::DEFAULT_CONCURRENCY
    ) {
        $this->s3 = $s3;
        $this->bucket = $bucket;
        $this->key = ltrim($key, '/');

        // Silently adjust part size if too small
        $this->partSize = max(self::MIN_PART_SIZE, $partSize);
        $this->concurrency = max(1, $concurrency ?? self::DEFAULT_CONCURRENCY);

        // Default parameters
        $this->putObjectParams = array_merge([
            'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'ACL' => 'private',
            'ContentDisposition' => 'attachment; filename="' . basename($key) . '"',
            'CacheControl' => 'no-cache',
        ], $putObjectParams);

        $this->initializeMultipartUpload();
    }

    /**
     * Initialize multipart upload
     */
    private function initializeMultipartUpload(): void
    {
        try {
            $result = $this->s3->createMultipartUpload(array_merge([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
            ], $this->putObjectParams));

            $this->uploadId = $result['UploadId'];
        } catch (AwsException $e) {
            throw S3Exception::multipartInitFailed(
                $this->bucket,
                $this->key,
                $e->getMessage()
            );
        }
    }

    /**
     * Write data to S3
     */
    public function write(string $data): void
    {
        if ($this->closed) {
            throw new \RuntimeException('Cannot write to closed sink');
        }

        $this->throwIfFailed();

        $this->buffer .= $data;
        $this->bytesWritten += strlen($data);

        // Dispatch parts when buffer reaches part size
        while (strlen($this->buffer) >= $this->partSize) {
            $chunk = substr($this->buffer, 0, $this->partSize);
            // Shrink the buffer before dispatching so we never hold
            // buffer + chunk + in-flight copies longer than necessary.
            $this->buffer = substr($this->buffer, $this->partSize);
            $this->dispatchPart($chunk);
        }
    }

    /**
     * Dispatch a part upload asynchronously, respecting the concurrency window.
     *
     * When the window is full, waits for the OLDEST in-flight part (FIFO).
     * While waiting, curl_multi advances ALL in-flight transfers, so newer
     * parts keep uploading concurrently; completion order is nondeterministic.
     */
    private function dispatchPart(string $chunk): void
    {
        while (count($this->inFlight) >= $this->concurrency) {
            $this->awaitOldest();
            $this->throwIfFailed();
        }

        $partNumber = $this->partNumber;
        $this->partNumber++;

        $this->inFlight[$partNumber] = $this->sendPartAsync($partNumber, $chunk);
    }

    /**
     * Send one part asynchronously.
     *
     * Transient errors on the in-flight request are handled by the AWS SDK's
     * retry middleware. If the promise still settles rejected, we re-dispatch
     * the part ONCE synchronously before recording the failure (first error
     * wins). Handlers never rethrow: failures surface exclusively through
     * $this->failed so FIFO waits stay simple.
     */
    private function sendPartAsync(int $partNumber, string $chunk): PromiseInterface
    {
        $params = [
            'Bucket' => $this->bucket,
            'Key' => $this->key,
            'UploadId' => $this->uploadId,
            'PartNumber' => $partNumber,
            'Body' => $chunk,
        ];

        return $this->s3->uploadPartAsync($params)->then(
            function ($result) use ($partNumber) {
                $this->recordPart($partNumber, (string) $result['ETag']);
            },
            function ($reason) use ($partNumber, $params) {
                if ($this->failed !== null || $reason instanceof CancellationException) {
                    // The sink is already failing (or this part was cancelled
                    // during teardown) — don't waste a re-dispatch.
                    unset($this->inFlight[$partNumber]);

                    return;
                }

                // One manual re-dispatch (synchronous) after SDK retries failed.
                try {
                    $result = $this->s3->uploadPart($params);
                    $this->recordPart($partNumber, (string) $result['ETag']);
                } catch (\Throwable $retryError) {
                    unset($this->inFlight[$partNumber]);

                    if ($this->failed === null) {
                        $originalMessage = $reason instanceof \Throwable
                            ? $reason->getMessage()
                            : (string) $reason;

                        $this->failed = S3Exception::partUploadFailed(
                            $partNumber,
                            "After SDK retries and one re-dispatch: {$retryError->getMessage()}"
                            . " (original error: {$originalMessage})"
                        );
                    }
                }
            }
        );
    }

    /**
     * Record a successfully uploaded part and release it from the window.
     */
    private function recordPart(int $partNumber, string $etag): void
    {
        $this->parts[$partNumber] = [
            'PartNumber' => $partNumber,
            'ETag' => $etag,
        ];

        unset($this->inFlight[$partNumber]);
    }

    /**
     * Wait for the oldest in-flight part to settle.
     *
     * wait(false) never throws: success/failure is recorded by the promise
     * handlers ($this->parts / $this->failed).
     */
    private function awaitOldest(): void
    {
        $oldest = array_key_first($this->inFlight);

        if ($oldest === null) {
            return;
        }

        $promise = $this->inFlight[$oldest];
        $promise->wait(false);

        // The fulfillment handler normally removes the entry; make sure a
        // settled promise never lingers in the window.
        unset($this->inFlight[$oldest]);
    }

    /**
     * Wait for every in-flight part to settle (never throws).
     */
    private function settleInFlight(): void
    {
        if (empty($this->inFlight)) {
            return;
        }

        $pending = $this->inFlight;
        $this->inFlight = [];

        try {
            Utils::settle($pending)->wait();
        } catch (\Throwable) {
            // Individual outcomes are recorded by the part handlers.
        }
    }

    /**
     * If a part upload failed, tear everything down and throw.
     *
     * Cancels the remaining in-flight parts, settles them, then aborts the
     * multipart upload (abort keeps its orphan-logging semantics).
     */
    private function throwIfFailed(): void
    {
        if ($this->failed === null) {
            return;
        }

        $failure = $this->failed;

        foreach ($this->inFlight as $promise) {
            try {
                $promise->cancel();
            } catch (\Throwable) {
                // Cancellation is best-effort.
            }
        }

        $this->settleInFlight();

        try {
            $this->abort();
        } catch (\Throwable) {
            // abort() already logs orphan upload details
        }

        throw $failure;
    }

    /**
     * Close and complete the multipart upload
     */
    public function close(): void
    {
        if ($this->closed) {
            return;
        }

        $this->throwIfFailed();

        try {
            // Dispatch remaining buffer as final part (may be < partSize)
            if ($this->buffer !== '') {
                $finalChunk = $this->buffer;
                $this->buffer = '';
                $this->dispatchPart($finalChunk);
            }

            // Settle all in-flight parts; completion order is nondeterministic.
            $this->settleInFlight();
            $this->throwIfFailed();

            // Complete multipart upload
            if (empty($this->parts)) {
                // Edge case: no data written, abort instead
                $this->abort();

                return;
            }

            // Every dispatched part must have reported an ETag.
            $expectedParts = $this->partNumber - 1;
            for ($part = 1; $part <= $expectedParts; $part++) {
                if (empty($this->parts[$part]['ETag'])) {
                    throw S3Exception::multipartCompleteFailed(
                        "Part {$part} has no ETag after settling all uploads"
                    );
                }
            }

            // S3 requires parts in ascending PartNumber order.
            ksort($this->parts);

            $this->s3->completeMultipartUpload([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
                'UploadId' => $this->uploadId,
                'MultipartUpload' => ['Parts' => array_values($this->parts)],
            ]);

            $this->closed = true;
        } catch (AwsException $e) {
            try {
                $this->abort();
            } catch (\Throwable) {
                // abort() already logs orphan upload details
            }

            throw S3Exception::multipartCompleteFailed($e->getMessage());
        } catch (S3Exception $e) {
            // throwIfFailed() already aborted; abort() is idempotent otherwise.
            try {
                $this->abort();
            } catch (\Throwable) {
                // abort() already logs orphan upload details
            }

            throw $e;
        }
    }

    /**
     * Abort the multipart upload.
     *
     * Safe to call with parts still in flight: pending promises are settled
     * first so S3 is not receiving parts after the abort.
     *
     * If S3 rejects the abort the upload becomes orphaned and S3 will keep
     * billing for its parts until a lifecycle rule cleans it up — so we log
     * a warning identifying the bucket/key/uploadId. Configure a bucket
     * lifecycle rule on `AbortIncompleteMultipartUpload` to bound exposure.
     */
    public function abort(): void
    {
        if ($this->closed || empty($this->uploadId)) {
            return;
        }

        $this->settleInFlight();

        try {
            $this->s3->abortMultipartUpload([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
                'UploadId' => $this->uploadId,
            ]);

            $this->closed = true;
        } catch (AwsException $e) {
            $this->closed = true;
            error_log(sprintf(
                '[kolay/xlsx-stream] Failed to abort multipart upload — orphan upload may incur S3 charges. bucket=%s key=%s uploadId=%s error=%s',
                $this->bucket,
                $this->key,
                $this->uploadId,
                $e->getMessage()
            ));
        }
    }

    /**
     * Get total bytes written
     */
    public function getBytesWritten(): int
    {
        return $this->bytesWritten;
    }

    /**
     * Destructor - ensure cleanup.
     *
     * Catches \Throwable (not just \Exception) because Errors thrown from
     * destructors crash the PHP process.
     */
    public function __destruct()
    {
        if (!$this->closed && !empty($this->uploadId)) {
            try {
                $this->abort();
            } catch (\Throwable) {
                // Last-resort: cannot recover from destructor.
            }
        }
    }
}

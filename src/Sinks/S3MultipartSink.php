<?php

namespace Kolay\XlsxStream\Sinks;

use Kolay\XlsxStream\Contracts\Sink;
use Kolay\XlsxStream\Exceptions\S3Exception;
use Aws\S3\S3Client;
use Aws\Exception\AwsException;

/**
 * S3 Multipart Upload Sink
 * 
 * Streams data directly to S3 using multipart upload
 * - Zero disk usage
 * - Constant memory (buffer size)
 * - Automatic part management
 * - Error recovery with abort
 */
class S3MultipartSink implements Sink
{
    private S3Client $s3;
    private string $bucket;
    private string $key;
    private string $uploadId;
    private array $parts = [];
    private string $buffer = '';
    private int $partSize;
    private int $partNumber = 1;
    private int $bytesWritten = 0;
    private array $putObjectParams;
    private bool $closed = false;
    
    // S3 minimum part size is 5MB (except last part)
    const MIN_PART_SIZE = 5242880; // 5MB
    const DEFAULT_PART_SIZE = 8388608; // 8MB
    
    /**
     * @param S3Client $s3
     * @param string $bucket
     * @param string $key S3 object key (path)
     * @param int $partSize Size of each part (min 5MB)
     * @param array $putObjectParams Additional S3 parameters (ContentType, ACL, etc)
     */
    public function __construct(
        S3Client $s3,
        string $bucket,
        string $key,
        int $partSize = self::DEFAULT_PART_SIZE,
        array $putObjectParams = []
    ) {
        $this->s3 = $s3;
        $this->bucket = $bucket;
        $this->key = ltrim($key, '/');
        
        // Silently adjust part size if too small
        $this->partSize = max(self::MIN_PART_SIZE, $partSize);
        
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
        
        $this->buffer .= $data;
        $this->bytesWritten += strlen($data);
        
        // Upload parts when buffer reaches part size
        while (strlen($this->buffer) >= $this->partSize) {
            $chunk = substr($this->buffer, 0, $this->partSize);
            $this->uploadPart($chunk);
            $this->buffer = substr($this->buffer, $this->partSize);
        }
    }
    
    /**
     * Upload a part to S3
     */
    private function uploadPart(string $chunk): void
    {
        $retries = 3;
        $lastException = null;
        
        for ($attempt = 1; $attempt <= $retries; $attempt++) {
            try {
                $result = $this->s3->uploadPart([
                    'Bucket' => $this->bucket,
                    'Key' => $this->key,
                    'UploadId' => $this->uploadId,
                    'PartNumber' => $this->partNumber,
                    'Body' => $chunk,
                ]);
                
                $this->parts[] = [
                    'PartNumber' => $this->partNumber,
                    'ETag' => $result['ETag'],
                ];
                
                $this->partNumber++;
                return;
                
            } catch (AwsException $e) {
                $lastException = $e;
                
                if ($attempt < $retries) {
                    // Exponential backoff
                    usleep((int)(100000 * pow(2, $attempt - 1))); // 100ms, 200ms, 400ms
                }
            }
        }
        
        throw S3Exception::partUploadFailed(
            $this->partNumber,
            "After {$retries} attempts: " . $lastException->getMessage()
        );
    }
    
    /**
     * Close and complete the multipart upload
     */
    public function close(): void
    {
        if ($this->closed) {
            return;
        }
        
        try {
            // Upload remaining buffer as final part
            if ($this->buffer !== '') {
                $this->uploadPart($this->buffer);
                $this->buffer = '';
            }
            
            // Complete multipart upload
            if (empty($this->parts)) {
                // Edge case: no data written, abort instead
                $this->abort();
                return;
            }
            
            $this->s3->completeMultipartUpload([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
                'UploadId' => $this->uploadId,
                'MultipartUpload' => ['Parts' => $this->parts],
            ]);
            
            $this->closed = true;
            
        } catch (AwsException $e) {
            // Try to abort on failure
            try {
                $this->abort();
            } catch (\Exception $abortException) {
                // Silently fail
            }
            
            throw S3Exception::multipartCompleteFailed($e->getMessage());
        }
    }
    
    /**
     * Abort the multipart upload
     */
    public function abort(): void
    {
        if ($this->closed || empty($this->uploadId)) {
            return;
        }
        
        try {
            $this->s3->abortMultipartUpload([
                'Bucket' => $this->bucket,
                'Key' => $this->key,
                'UploadId' => $this->uploadId,
            ]);
            
            $this->closed = true;
            
        } catch (AwsException $e) {
            // Silently fail abort
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
     * Destructor - ensure cleanup
     */
    public function __destruct()
    {
        if (!$this->closed && !empty($this->uploadId)) {
            try {
                $this->abort();
            } catch (\Exception $e) {
                // Silently fail in destructor
            }
        }
    }
}
<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Tests\TestCase;

use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Aws\S3\S3Client;

/**
 * Real S3 Integration Test
 *
 * This test actually uploads to AWS S3
 * Requires valid AWS credentials in .env
 */
class RealS3Test extends TestCase
{
    private ?S3Client $s3Client = null;

    protected function setUp(): void
    {
        parent::setUp();

        // Skip if AWS credentials are not set
        if (!$this->hasAwsCredentials()) {
            $this->markTestSkipped('AWS credentials not configured. Set AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_DEFAULT_REGION, AWS_BUCKET in .env');
        }

        $this->s3Client = new S3Client([
            'region' => getenv('AWS_DEFAULT_REGION') ?: 'us-east-1',
            'version' => 'latest',
            'credentials' => [
                'key' => getenv('AWS_ACCESS_KEY_ID'),
                'secret' => getenv('AWS_SECRET_ACCESS_KEY'),
            ],
        ]);
    }

    private function hasAwsCredentials(): bool
    {
        return !empty(getenv('AWS_ACCESS_KEY_ID'))
            && !empty(getenv('AWS_SECRET_ACCESS_KEY'))
            && !empty(getenv('AWS_BUCKET'));
    }

    public function test_real_s3_upload_small_file()
    {
        $bucket = getenv('AWS_BUCKET');
        $key = 'tests/test-' . uniqid() . '.xlsx';

        // Create S3 sink
        $sink = new S3MultipartSink(
            $this->s3Client,
            $bucket,
            $key,
            5 * 1024 * 1024 // 5MB parts
        );

        $writer = new SinkableXlsxWriter($sink);
        $writer->startFile(['ID', 'Name', 'Email', 'Status']);

        // Write 100 rows
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([
                $i,
                "User {$i}",
                "user{$i}@example.com",
                $i % 2 === 0 ? 'Active' : 'Inactive'
            ]);
        }

        $stats = $writer->finishFile();

        // Assertions
        $this->assertEquals(100, $stats['rows']);
        $this->assertEquals(1, $stats['sheets']);
        $this->assertGreaterThan(0, $stats['bytes']);

        // Verify file exists in S3
        $this->assertTrue(
            $this->s3Client->doesObjectExist($bucket, $key),
            "File should exist in S3 at {$bucket}/{$key}"
        );

        // Get file size
        $result = $this->s3Client->headObject([
            'Bucket' => $bucket,
            'Key' => $key,
        ]);

        $this->assertGreaterThan(0, $result['ContentLength']);

        // STDERR bypasses PHPUnit's output capture (beStrictAboutOutputDuringTests)
        fwrite(STDERR, "\n✓ Successfully uploaded {$stats['rows']} rows to S3");
        fwrite(STDERR, "\n✓ File size: " . number_format($result['ContentLength']) . " bytes");
        fwrite(STDERR, "\n✓ Location: s3://{$bucket}/{$key}\n");
    }

    public function test_real_s3_upload_larger_file()
    {
        $bucket = getenv('AWS_BUCKET');
        $key = 'tests/test-large-' . uniqid() . '.xlsx';

        $sink = new S3MultipartSink(
            $this->s3Client,
            $bucket,
            $key,
            5 * 1024 * 1024 // 5MB parts
        );

        $writer = new SinkableXlsxWriter($sink);
        $writer->setCompressionLevel(1) // Fast compression
               ->setBufferFlushInterval(1000);

        $writer->startFile(['ID', 'Name', 'Email', 'Age', 'City', 'Status', 'Created']);

        // Write 10,000 rows to trigger multipart upload
        for ($i = 1; $i <= 10000; $i++) {
            $writer->writeRow([
                $i,
                "User Name {$i}",
                "user{$i}@example.com",
                rand(18, 80),
                "City " . ($i % 100),
                $i % 3 === 0 ? 'Active' : 'Inactive',
                date('Y-m-d H:i:s')
            ]);
        }

        $stats = $writer->finishFile();

        // Assertions
        $this->assertEquals(10000, $stats['rows']);
        $this->assertEquals(1, $stats['sheets']);
        $this->assertGreaterThan(0, $stats['bytes']);

        // Verify file exists
        $this->assertTrue(
            $this->s3Client->doesObjectExist($bucket, $key),
            "Large file should exist in S3"
        );

        $result = $this->s3Client->headObject([
            'Bucket' => $bucket,
            'Key' => $key,
        ]);

        // STDERR bypasses PHPUnit's output capture (beStrictAboutOutputDuringTests)
        fwrite(STDERR, "\n✓ Successfully uploaded {$stats['rows']} rows to S3");
        fwrite(STDERR, "\n✓ File size: " . number_format($result['ContentLength']) . " bytes");
        fwrite(STDERR, "\n✓ Location: s3://{$bucket}/{$key}\n");
    }

    /**
     * Bounded-memory regression guard for the default (synchronous) S3
     * sink. Writing 1M rows produces a multi-part (>8 MB) upload, so the
     * part-boundary GC path is exercised. Peak memory must stay flat at
     * ~the part-buffer working set regardless of row count — before the
     * fix the AWS SDK's async promise/command graph held every dispatched
     * part's body and memory climbed toward the whole file (≈54 MB at 1M,
     * unbounded beyond). The default concurrency=1 + per-part
     * gc_collect_cycles() keeps it O(1).
     */
    public function test_s3_write_memory_stays_bounded_regardless_of_row_count()
    {
        $bucket = getenv('AWS_BUCKET');
        $key = 'tests/mem-' . uniqid() . '.xlsx';

        $sink = new S3MultipartSink($this->s3Client, $bucket, $key); // default concurrency=1
        $writer = new SinkableXlsxWriter($sink);
        $writer->setCompressionLevel(1)->setBufferFlushInterval(10000);

        gc_collect_cycles();
        memory_reset_peak_usage();
        $before = memory_get_usage(true);

        $writer->startFile(['ID', 'Name', 'Email', 'Department', 'Salary', 'Date', 'Status', 'Notes']);
        for ($i = 1; $i <= 1_000_000; $i++) {
            $writer->writeRow([
                $i, 'Employee' . ($i % 100), 'emp' . ($i % 100) . '@company.com',
                'Dept' . ($i % 10), 50000 + ($i % 50) * 1000, '2026-05-09',
                'Status' . ($i % 5), 'Standard notes for employee record',
            ]);
        }
        $writer->finishFile();

        $peakDelta = (memory_get_peak_usage(true) - $before) / 1024 / 1024;

        // Cleanup before asserting so a failure still removes the object.
        $this->s3Client->deleteObject(['Bucket' => $bucket, 'Key' => $key]);

        fwrite(STDERR, "\n✓ 1M-row S3 write peak memory delta: " . round($peakDelta, 1) . " MB\n");

        // Fixed path holds ~30 MB; the pre-fix ratchet exceeded ~50 MB at 1M
        // and grew with the file. 45 MB cleanly separates the two.
        $this->assertLessThan(
            45,
            $peakDelta,
            "S3 write peak memory ({$peakDelta} MB) grew with row count — the multipart sink is leaking part bodies again."
        );
    }
}

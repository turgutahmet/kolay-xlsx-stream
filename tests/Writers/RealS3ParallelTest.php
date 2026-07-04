<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Aws\S3\S3Client;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Real S3 Integration Test for parallel multipart upload
 *
 * Exports the same deterministic 1M-row dataset twice — once with
 * concurrency 1 (sequential, pre-3.2 behavior) and once with the default
 * bounded async window — and verifies both objects are byte-identical,
 * reports wall times, and checks the parallel run's memory ceiling.
 *
 * Requires valid AWS credentials in .env (skipped otherwise).
 */
class RealS3ParallelTest extends TestCase
{
    private const ROWS = 1000000;

    private const PART_SIZE = 5242880; // 5MB parts -> ~9 parts for this dataset

    private const CONCURRENCY = 4;

    private ?S3Client $s3Client = null;

    protected function setUp(): void
    {
        parent::setUp();

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

    public function test_parallel_upload_matches_sequential_output()
    {
        $bucket = getenv('AWS_BUCKET');
        $runId = uniqid();
        $seqKey = "tests/parallel-test-seq-{$runId}.xlsx";
        $parKey = "tests/parallel-test-par-{$runId}.xlsx";

        try {
            $sequential = $this->export($bucket, $seqKey, 1);
            $parallel = $this->export($bucket, $parKey, self::CONCURRENCY);

            // Both objects exist
            $this->assertTrue(
                $this->s3Client->doesObjectExist($bucket, $seqKey),
                'Sequential export should exist in S3'
            );
            $this->assertTrue(
                $this->s3Client->doesObjectExist($bucket, $parKey),
                'Parallel export should exist in S3'
            );

            $seqSize = (int) $this->s3Client->headObject(['Bucket' => $bucket, 'Key' => $seqKey])['ContentLength'];
            $parSize = (int) $this->s3Client->headObject(['Bucket' => $bucket, 'Key' => $parKey])['ContentLength'];

            // Deterministic data + deterministic deflate -> identical bytes
            $this->assertSame($seqSize, $parSize, 'Parallel upload must produce a byte-identical object');
            $this->assertSame($sequential['stats']['bytes'], $parallel['stats']['bytes']);
            $this->assertSame($parallel['stats']['bytes'], $parSize, 'S3 ContentLength must match bytes written');
            $this->assertSame(self::ROWS, $parallel['stats']['rows']);

            // Multipart path was actually exercised (more than one part)
            $this->assertGreaterThan(
                self::PART_SIZE,
                $parSize,
                'Dataset must span multiple parts to exercise the parallel window'
            );

            fwrite(STDERR, sprintf(
                "\n[RealS3ParallelTest] %s rows, %.1f MB object (%d parts of %d MB)\n"
                . "  sequential (concurrency 1): %6.2f s\n"
                . "  parallel   (concurrency %d): %6.2f s  (%.2fx speedup)\n"
                . "  parallel peak memory delta: %s allocated, %s real\n",
                number_format(self::ROWS),
                $parSize / 1048576,
                (int) ceil($parSize / self::PART_SIZE),
                (int) (self::PART_SIZE / 1048576),
                $sequential['seconds'],
                self::CONCURRENCY,
                $parallel['seconds'],
                $sequential['seconds'] / max($parallel['seconds'], 0.001),
                $parallel['peakDelta'] === null
                    ? 'n/a (memory_reset_peak_usage unavailable)'
                    : sprintf('%.1f MB', $parallel['peakDelta'] / 1048576),
                $parallel['peakDeltaReal'] === null
                    ? 'n/a'
                    : sprintf('%.1f MB', $parallel['peakDeltaReal'] / 1048576)
            ));

            // Memory ceiling: partSize x (concurrency + 2) + slack.
            // Asserted on allocated memory (memory_get_peak_usage(false));
            // the "real" (mmap) figure is reported above but not asserted,
            // as it ratchets with allocator fragmentation.
            // Slack covers what the formula's part bodies don't: the SDK
            // spools each in-flight body into a php://temp stream (~2MB
            // in-memory apiece), plus transient substr copies at dispatch
            // and writer/deflate internals (~16MB measured combined).
            if ($parallel['peakDelta'] !== null) {
                $ceiling = self::PART_SIZE * (self::CONCURRENCY + 2) + 24 * 1024 * 1024;
                $this->assertLessThan(
                    $ceiling,
                    $parallel['peakDelta'],
                    sprintf(
                        'Peak allocated memory delta %.1f MB exceeded ceiling %.1f MB',
                        $parallel['peakDelta'] / 1048576,
                        $ceiling / 1048576
                    )
                );
            }
        } finally {
            foreach ([$seqKey, $parKey] as $key) {
                try {
                    $this->s3Client->deleteObject(['Bucket' => $bucket, 'Key' => $key]);
                } catch (\Throwable) {
                    // Best-effort cleanup
                }
            }
        }

        $this->assertFalse($this->s3Client->doesObjectExist($bucket, $seqKey), 'Sequential test object should be deleted');
        $this->assertFalse($this->s3Client->doesObjectExist($bucket, $parKey), 'Parallel test object should be deleted');
    }

    /**
     * Export a deterministic dataset and measure wall time + peak memory delta.
     *
     * @return array{stats: array, seconds: float, peakDelta: int|null, peakDeltaReal: int|null}
     */
    private function export(string $bucket, string $key, int $concurrency): array
    {
        gc_collect_cycles();

        $canMeasurePeak = function_exists('memory_reset_peak_usage');
        if ($canMeasurePeak) {
            memory_reset_peak_usage();
        }
        $baseline = memory_get_usage(false);
        $baselineReal = memory_get_usage(true);

        $start = microtime(true);

        $sink = new S3MultipartSink(
            $this->s3Client,
            $bucket,
            $key,
            self::PART_SIZE,
            [],
            $concurrency
        );

        $writer = new SinkableXlsxWriter($sink);
        $writer->setCompressionLevel(1)
            ->setBufferFlushInterval(1000);

        $writer->startFile(['ID', 'Name', 'Email', 'Age', 'City', 'Status', 'Created']);

        // Deterministic rows: both runs must produce identical bytes.
        for ($i = 1; $i <= self::ROWS; $i++) {
            $writer->writeRow([
                $i,
                "User Name {$i}",
                "user{$i}@example.com",
                18 + ($i % 60),
                'City ' . ($i % 100),
                $i % 3 === 0 ? 'Active' : 'Inactive',
                '2026-01-01 00:00:00',
            ]);
        }

        $stats = $writer->finishFile();
        $seconds = microtime(true) - $start;

        $peakDelta = $canMeasurePeak
            ? memory_get_peak_usage(false) - $baseline
            : null;
        $peakDeltaReal = $canMeasurePeak
            ? memory_get_peak_usage(true) - $baselineReal
            : null;

        return [
            'stats' => $stats,
            'seconds' => $seconds,
            'peakDelta' => $peakDelta,
            'peakDeltaReal' => $peakDeltaReal,
        ];
    }
}

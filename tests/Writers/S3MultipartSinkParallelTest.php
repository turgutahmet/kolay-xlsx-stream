<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Aws\S3\S3Client;
use GuzzleHttp\Promise\FulfilledPromise;
use GuzzleHttp\Promise\Promise;
use GuzzleHttp\Promise\RejectedPromise;
use Kolay\XlsxStream\Exceptions\S3Exception;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Kolay\XlsxStream\Tests\TestCase;
use Mockery;

/**
 * Unit tests for the parallel (bounded async window) S3 multipart sink.
 *
 * Uses a mocked S3Client with hand-built Guzzle promises so completion
 * order, failures and window occupancy can be controlled deterministically.
 */
class S3MultipartSinkParallelTest extends TestCase
{
    private const PART = S3MultipartSink::MIN_PART_SIZE; // 5MB

    private string $previousMemoryLimit = '';

    protected function setUp(): void
    {
        parent::setUp();

        // S3's minimum part size is 5MB and Mockery records every invocation
        // argument (i.e. every part body), so these tests need headroom
        // beyond the default 128M CLI limit.
        $this->previousMemoryLimit = (string) ini_get('memory_limit');
        ini_set('memory_limit', '512M');
    }

    protected function tearDown(): void
    {
        Mockery::close();

        // Promise <-> handler closures form cycles that keep 5MB part bodies
        // alive until the cycle collector runs; collect before shrinking the
        // limit back down (and tolerate a restore that would not fit).
        gc_collect_cycles();

        try {
            ini_set('memory_limit', $this->previousMemoryLimit);
        } catch (\Throwable) {
            // Current usage still above the previous limit; keep the raised one.
        }

        parent::tearDown();
    }

    private function mockS3(): S3Client|Mockery\MockInterface
    {
        $s3 = Mockery::mock(S3Client::class);
        $s3->shouldReceive('createMultipartUpload')
            ->once()
            ->andReturn(['UploadId' => 'test-upload-id']);

        return $s3;
    }

    public function test_parts_completed_out_of_order_are_sorted_before_complete()
    {
        $s3 = $this->mockS3();

        /** @var array<int, Promise> $promises */
        $promises = [];
        $s3->shouldReceive('uploadPartAsync')
            ->times(3)
            ->andReturnUsing(function (array $args) use (&$promises) {
                $promise = new Promise();
                $promises[$args['PartNumber']] = $promise;

                return $promise;
            });

        $completedParts = null;
        $s3->shouldReceive('completeMultipartUpload')
            ->once()
            ->andReturnUsing(function (array $args) use (&$completedParts) {
                $completedParts = $args['MultipartUpload']['Parts'];

                return [];
            });

        // Window larger than part count: all three stay in flight together.
        $sink = new S3MultipartSink($s3, 'bucket', 'key.xlsx', self::PART, [], 8);
        $sink->write(str_repeat('a', self::PART * 3));

        // Simulate nondeterministic network: parts complete in REVERSE order.
        $promises[3]->resolve(['ETag' => 'etag-3']);
        $promises[2]->resolve(['ETag' => 'etag-2']);
        $promises[1]->resolve(['ETag' => 'etag-1']);

        $sink->close();

        $this->assertSame([
            ['PartNumber' => 1, 'ETag' => 'etag-1'],
            ['PartNumber' => 2, 'ETag' => 'etag-2'],
            ['PartNumber' => 3, 'ETag' => 'etag-3'],
        ], $completedParts, 'Parts must reach completeMultipartUpload in ascending PartNumber order');
    }

    public function test_concurrency_window_is_never_exceeded()
    {
        $s3 = $this->mockS3();

        $live = 0;
        $maxLive = 0;
        $s3->shouldReceive('uploadPartAsync')
            ->times(6)
            ->andReturnUsing(function (array $args) use (&$live, &$maxLive) {
                $live++;
                $maxLive = max($maxLive, $live);

                $promise = null;
                $promise = new Promise(function () use (&$promise, &$live, $args) {
                    $live--;
                    $promise->resolve(['ETag' => 'etag-' . $args['PartNumber']]);
                });

                return $promise;
            });

        $completedParts = null;
        $s3->shouldReceive('completeMultipartUpload')
            ->once()
            ->andReturnUsing(function (array $args) use (&$completedParts) {
                $completedParts = $args['MultipartUpload']['Parts'];

                return [];
            });

        $sink = new S3MultipartSink($s3, 'bucket', 'key.xlsx', self::PART, [], 2);

        for ($i = 0; $i < 6; $i++) {
            $sink->write(str_repeat('b', self::PART));
        }
        $sink->close();

        $this->assertSame(2, $maxLive, 'No more than $concurrency parts may be in flight at once');
        $this->assertCount(6, $completedParts);
        $this->assertSame(range(1, 6), array_column($completedParts, 'PartNumber'));
    }

    public function test_rejected_part_is_redispatched_once_then_succeeds()
    {
        $s3 = $this->mockS3();

        $s3->shouldReceive('uploadPartAsync')
            ->times(3)
            ->andReturnUsing(function (array $args) {
                if ($args['PartNumber'] === 2) {
                    return new RejectedPromise(new \RuntimeException('transient network error'));
                }

                return new FulfilledPromise(['ETag' => 'etag-' . $args['PartNumber']]);
            });

        // The single synchronous re-dispatch for part 2 succeeds.
        $s3->shouldReceive('uploadPart')
            ->once()
            ->with(Mockery::on(fn (array $args) => $args['PartNumber'] === 2
                && strlen($args['Body']) === self::PART))
            ->andReturn(['ETag' => 'etag-2-retried']);

        $completedParts = null;
        $s3->shouldReceive('completeMultipartUpload')
            ->once()
            ->andReturnUsing(function (array $args) use (&$completedParts) {
                $completedParts = $args['MultipartUpload']['Parts'];

                return [];
            });

        $sink = new S3MultipartSink($s3, 'bucket', 'key.xlsx', self::PART, [], 8);
        $sink->write(str_repeat('c', self::PART * 3));
        $sink->close();

        $this->assertSame([
            ['PartNumber' => 1, 'ETag' => 'etag-1'],
            ['PartNumber' => 2, 'ETag' => 'etag-2-retried'],
            ['PartNumber' => 3, 'ETag' => 'etag-3'],
        ], $completedParts);
    }

    public function test_first_error_aborts_upload_and_write_throws()
    {
        $s3 = $this->mockS3();

        // Synchronous default (concurrency 1): uploadPart is called directly
        // and its throw is recorded as the failure (the SDK's own retry
        // middleware already handled transience — no manual re-dispatch here).
        $s3->shouldReceive('uploadPart')
            ->once()
            ->andThrow(new \RuntimeException('still failing'));

        $s3->shouldReceive('abortMultipartUpload')
            ->once()
            ->with(Mockery::on(fn (array $args) => $args['UploadId'] === 'test-upload-id'))
            ->andReturn([]);

        $sink = new S3MultipartSink($s3, 'bucket', 'key.xlsx', self::PART, [], 1);
        $sink->write(str_repeat('d', self::PART)); // uploads part 1 -> throws -> failure recorded

        try {
            // The next write() sees the recorded failure, aborts and throws.
            $sink->write(str_repeat('d', self::PART));
            $this->fail('Expected S3Exception was not thrown');
        } catch (S3Exception $e) {
            $this->assertStringContainsString('part 1', $e->getMessage());
            $this->assertStringContainsString('still failing', $e->getMessage());
        }

        // The sink aborted itself: further writes must be refused.
        $this->expectException(\RuntimeException::class);
        $this->expectExceptionMessage('Cannot write to closed sink');
        $sink->write('more data');
    }

    public function test_failure_cancels_remaining_in_flight_without_redispatching_them()
    {
        $s3 = $this->mockS3();

        $s3->shouldReceive('uploadPartAsync')
            ->times(3)
            ->andReturnUsing(function (array $args) {
                if ($args['PartNumber'] === 1) {
                    return new RejectedPromise(new \RuntimeException('boom'));
                }

                // Parts 2 and 3 stay pending; they can only be cancelled.
                return new Promise();
            });

        // Exactly ONE synchronous re-dispatch (for part 1); the cancelled
        // parts 2 and 3 must NOT be re-dispatched.
        $s3->shouldReceive('uploadPart')
            ->once()
            ->with(Mockery::on(fn (array $args) => $args['PartNumber'] === 1))
            ->andThrow(new \RuntimeException('still failing'));

        $s3->shouldReceive('abortMultipartUpload')->once()->andReturn([]);

        $sink = new S3MultipartSink($s3, 'bucket', 'key.xlsx', self::PART, [], 3);
        $sink->write(str_repeat('e', self::PART * 3)); // parts 1-3 in flight

        $this->expectException(S3Exception::class);
        $this->expectExceptionMessageMatches('/part 1/');

        // Fourth part: window full -> waits part 1 -> failure surfaces.
        $sink->write(str_repeat('e', self::PART));
    }

    public function test_final_partial_part_and_complete_payload()
    {
        $s3 = $this->mockS3();

        $bodySizes = [];
        $s3->shouldReceive('uploadPartAsync')
            ->times(2)
            ->andReturnUsing(function (array $args) use (&$bodySizes) {
                $bodySizes[$args['PartNumber']] = strlen($args['Body']);

                return new FulfilledPromise(['ETag' => 'etag-' . $args['PartNumber']]);
            });

        $completedArgs = null;
        $s3->shouldReceive('completeMultipartUpload')
            ->once()
            ->andReturnUsing(function (array $args) use (&$completedArgs) {
                $completedArgs = $args;

                return [];
            });

        $sink = new S3MultipartSink($s3, 'bucket', 'key.xlsx', self::PART, [], 4);
        $sink->write(str_repeat('f', self::PART + 100));
        $sink->close();

        $this->assertSame([1 => self::PART, 2 => 100], $bodySizes);
        $this->assertSame(self::PART + 100, $sink->getBytesWritten());
        $this->assertSame('bucket', $completedArgs['Bucket']);
        $this->assertSame('key.xlsx', $completedArgs['Key']);
        $this->assertSame('test-upload-id', $completedArgs['UploadId']);
        $this->assertSame([
            ['PartNumber' => 1, 'ETag' => 'etag-1'],
            ['PartNumber' => 2, 'ETag' => 'etag-2'],
        ], $completedArgs['MultipartUpload']['Parts']);
    }

    public function test_concurrency_one_keeps_sequential_request_sequence()
    {
        $s3 = $this->mockS3();

        $events = [];
        // concurrency 1 = synchronous uploadPart: each part fully uploads
        // (call returns) before the next is sent — inherently sequential.
        $s3->shouldReceive('uploadPart')
            ->times(3)
            ->andReturnUsing(function (array $args) use (&$events) {
                $n = $args['PartNumber'];
                $events[] = "part-{$n}";

                return ['ETag' => 'etag-' . $n];
            });

        $s3->shouldReceive('completeMultipartUpload')
            ->once()
            ->andReturnUsing(function () use (&$events) {
                $events[] = 'complete-upload';

                return [];
            });

        $sink = new S3MultipartSink($s3, 'bucket', 'key.xlsx', self::PART, [], 1);
        $sink->write(str_repeat('g', self::PART * 3));
        $sink->close();

        // Parts upload in strict order, each finishing before the next.
        $this->assertSame([
            'part-1',
            'part-2',
            'part-3',
            'complete-upload',
        ], $events);
    }

    public function test_abort_settles_pending_parts_before_aborting()
    {
        $s3 = $this->mockS3();

        $events = [];
        $s3->shouldReceive('uploadPartAsync')
            ->once()
            ->andReturnUsing(function (array $args) use (&$events) {
                $promise = null;
                $promise = new Promise(function () use (&$promise, &$events) {
                    $events[] = 'part-settled';
                    $promise->resolve(['ETag' => 'etag-1']);
                });

                return $promise;
            });

        $s3->shouldReceive('abortMultipartUpload')
            ->once()
            ->andReturnUsing(function () use (&$events) {
                $events[] = 'abort';

                return [];
            });

        $sink = new S3MultipartSink($s3, 'bucket', 'key.xlsx', self::PART, [], 4);
        $sink->write(str_repeat('h', self::PART)); // part 1 in flight

        $sink->abort();

        $this->assertSame(['part-settled', 'abort'], $events, 'abort() must settle in-flight parts before aborting the upload');
    }
}

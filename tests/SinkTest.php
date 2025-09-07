<?php

namespace Kolay\XlsxStream\Tests;

use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Aws\S3\S3Client;
use Mockery;

class SinkTest extends TestCase
{
    protected function tearDown(): void
    {
        Mockery::close();
        parent::tearDown();
    }
    
    public function test_file_sink_creates_directory_if_not_exists()
    {
        $tempDir = sys_get_temp_dir() . '/xlsx_test_' . uniqid();
        $filePath = $tempDir . '/subdir/test.xlsx';
        
        $this->assertDirectoryDoesNotExist($tempDir);
        
        $sink = new FileSink($filePath);
        $sink->write('test data');
        $sink->close();
        
        $this->assertFileExists($filePath);
        $this->assertEquals('test data', file_get_contents($filePath));
        
        // Cleanup
        unlink($filePath);
        rmdir($tempDir . '/subdir');
        rmdir($tempDir);
    }
    
    public function test_file_sink_abort_deletes_file()
    {
        $tempFile = sys_get_temp_dir() . '/test_abort_' . uniqid() . '.xlsx';
        
        $sink = new FileSink($tempFile);
        $sink->write('some data');
        
        $this->assertFileExists($tempFile);
        
        $sink->abort();
        
        $this->assertFileDoesNotExist($tempFile);
    }
    
    public function test_file_sink_tracks_bytes_written()
    {
        $tempFile = sys_get_temp_dir() . '/test_bytes_' . uniqid() . '.xlsx';
        
        $sink = new FileSink($tempFile);
        
        $this->assertEquals(0, $sink->getBytesWritten());
        
        $sink->write('12345');
        $this->assertEquals(5, $sink->getBytesWritten());
        
        $sink->write('67890');
        $this->assertEquals(10, $sink->getBytesWritten());
        
        $sink->close();
        
        // Cleanup
        unlink($tempFile);
    }
    
    public function test_file_sink_throws_on_write_after_close()
    {
        $tempFile = sys_get_temp_dir() . '/test_closed_' . uniqid() . '.xlsx';
        
        $sink = new FileSink($tempFile);
        $sink->write('data');
        $sink->close();
        
        $this->expectException(\RuntimeException::class);
        $this->expectExceptionMessage('Cannot write to closed sink');
        
        $sink->write('more data');
    }
    
    public function test_s3_multipart_sink_validates_part_size()
    {
        $s3Client = Mockery::mock(S3Client::class);
        
        $s3Client->shouldReceive('createMultipartUpload')
            ->once()
            ->andReturn(['UploadId' => 'test-upload-id']);
        
        // Part size less than 5MB should be adjusted to 5MB
        $sink = new S3MultipartSink(
            $s3Client,
            'test-bucket',
            'test-key.xlsx',
            1024 * 1024 // 1MB - should be adjusted to 5MB
        );
        
        // Write less than 5MB - should not trigger upload yet
        $sink->write(str_repeat('a', 1024 * 1024)); // 1MB
        
        // Verify no upload was triggered (buffering)
        $s3Client->shouldNotHaveReceived('uploadPart');
        
        // Add assertion
        $this->assertTrue(true, 'Part size validation works correctly');
    }
    
    public function test_s3_multipart_sink_uploads_parts()
    {
        $s3Client = Mockery::mock(S3Client::class);
        
        $s3Client->shouldReceive('createMultipartUpload')
            ->once()
            ->andReturn(['UploadId' => 'test-upload-id']);
        
        $s3Client->shouldReceive('uploadPart')
            ->once()
            ->with(Mockery::on(function ($args) {
                return $args['PartNumber'] === 1 
                    && strlen($args['Body']) === 5 * 1024 * 1024;
            }))
            ->andReturn(['ETag' => 'etag-1']);
        
        $s3Client->shouldReceive('completeMultipartUpload')
            ->once()
            ->andReturn([]);
        
        $sink = new S3MultipartSink(
            $s3Client,
            'test-bucket',
            'test-key.xlsx',
            5 * 1024 * 1024 // 5MB parts
        );
        
        // Write exactly 5MB to trigger part upload
        $sink->write(str_repeat('a', 5 * 1024 * 1024));
        
        $sink->close();
        
        // Add assertion
        $this->assertEquals(5 * 1024 * 1024, $sink->getBytesWritten());
    }
    
    public function test_s3_multipart_sink_aborts_on_error()
    {
        $s3Client = Mockery::mock(S3Client::class);
        
        $s3Client->shouldReceive('createMultipartUpload')
            ->once()
            ->andReturn(['UploadId' => 'test-upload-id']);
        
        $s3Client->shouldReceive('abortMultipartUpload')
            ->once()
            ->with(Mockery::on(function ($args) {
                return $args['UploadId'] === 'test-upload-id';
            }));
        
        $sink = new S3MultipartSink(
            $s3Client,
            'test-bucket',
            'test-key.xlsx'
        );
        
        $sink->abort();
        
        // Add assertion
        $this->assertTrue(true, 'S3 multipart upload aborted successfully');
    }
}
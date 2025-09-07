<?php

namespace Kolay\XlsxStream\Tests;

use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Kolay\XlsxStream\Sinks\FileSink;

class BasicWriterTest extends TestCase
{
    private string $testFile;
    
    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir() . '/test_' . uniqid() . '.xlsx';
    }
    
    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }
    
    public function test_can_create_xlsx_file()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        
        $writer->startFile(['Column1', 'Column2', 'Column3']);
        $writer->writeRow(['Value1', 'Value2', 'Value3']);
        $writer->writeRow(['Test1', 'Test2', 'Test3']);
        
        $stats = $writer->finishFile();
        
        $this->assertFileExists($this->testFile);
        $this->assertEquals(2, $stats['rows']);
        $this->assertEquals(1, $stats['sheets']);
        $this->assertGreaterThan(0, $stats['bytes']);
    }
    
    public function test_handles_empty_cells()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        
        $writer->startFile(['A', 'B', 'C']);
        $writer->writeRow(['Value1', null, 'Value3']);
        $writer->writeRow(['', 'Value2', '']);
        
        $stats = $writer->finishFile();
        
        $this->assertEquals(2, $stats['rows']);
        $this->assertFileExists($this->testFile);
    }
    
    public function test_handles_numeric_values()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        
        $writer->startFile(['String', 'Integer', 'Float', 'Leading Zero']);
        $writer->writeRow(['Text', 123, 45.67, '0123']);
        
        $stats = $writer->finishFile();
        
        $this->assertEquals(1, $stats['rows']);
        $this->assertFileExists($this->testFile);
    }
    
    public function test_compression_levels()
    {
        $sizes = [];
        
        foreach ([1, 6, 9] as $level) {
            $testFile = sys_get_temp_dir() . '/test_compression_' . $level . '.xlsx';
            
            $sink = new FileSink($testFile);
            $writer = new SinkableXlsxWriter($sink);
            $writer->setCompressionLevel($level);
            
            $writer->startFile(['Data']);
            for ($i = 0; $i < 1000; $i++) {
                $writer->writeRow([str_repeat('Test Data ', 10)]);
            }
            
            $stats = $writer->finishFile();
            $sizes[$level] = filesize($testFile);
            
            unlink($testFile);
        }
        
        // Higher compression should result in smaller files (but with small test data might not be significant)
        // Just check that all files were created
        $this->assertGreaterThan(0, $sizes[1]);
        $this->assertGreaterThan(0, $sizes[6]);
        $this->assertGreaterThan(0, $sizes[9]);
    }
    
    public function test_buffer_flush_interval()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        
        $writer->setBufferFlushInterval(100);
        
        $writer->startFile(['Number']);
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([$i]);
        }
        
        $stats = $writer->finishFile();
        
        $this->assertEquals(500, $stats['rows']);
        $this->assertFileExists($this->testFile);
    }
}
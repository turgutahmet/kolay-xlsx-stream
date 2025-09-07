<?php

namespace Kolay\XlsxStream\Tests;

use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Kolay\XlsxStream\Sinks\FileSink;

class LargeFileTest extends TestCase
{
    private string $testFile;
    
    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir() . '/large_test_' . uniqid() . '.xlsx';
    }
    
    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        parent::tearDown();
    }
    
    public function test_handles_multi_sheet_automatically()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        
        $writer->setCompressionLevel(1)
               ->setBufferFlushInterval(10000);
        
        $writer->startFile(['ID', 'Data']);
        
        // Write 1,048,576 rows to trigger second sheet
        // Excel limit is 1,048,576 but we use 1,048,575 for safety
        $rowsToWrite = 1048576; 
        
        for ($i = 1; $i <= $rowsToWrite; $i++) {
            $writer->writeRow([$i, "Row $i data"]);
        }
        
        $stats = $writer->finishFile();
        
        $this->assertEquals($rowsToWrite, $stats['rows']);
        $this->assertEquals(2, $stats['sheets']); // Should create 2 sheets
        $this->assertFileExists($this->testFile);
        
        // Verify sheet distribution - headers are counted
        $this->assertEquals(1048575, $stats['sheet_details'][0]['rows']); // First sheet (max)
        $this->assertEquals(3, $stats['sheet_details'][1]['rows']); // Second sheet (1 data row + 2 headers)
    }
    
    public function test_memory_usage_stays_constant()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        
        $writer->setCompressionLevel(1)
               ->setBufferFlushInterval(1000);
        
        $writer->startFile(['ID', 'Name', 'Email', 'Description']);
        
        $initialMemory = memory_get_usage(true);
        $maxMemoryDiff = 0;
        
        // Write 100k rows and monitor memory
        for ($i = 1; $i <= 100000; $i++) {
            $writer->writeRow([
                $i,
                "User Name $i",
                "user$i@example.com",
                str_repeat("Description text ", 10)
            ]);
            
            if ($i % 10000 === 0) {
                $currentMemory = memory_get_usage(true);
                $memoryDiff = $currentMemory - $initialMemory;
                $maxMemoryDiff = max($maxMemoryDiff, $memoryDiff);
            }
        }
        
        $stats = $writer->finishFile();
        
        // Memory increase should be minimal (less than 50MB)
        $maxMemoryMB = $maxMemoryDiff / 1024 / 1024;
        $this->assertLessThan(50, $maxMemoryMB, "Memory usage exceeded 50MB: {$maxMemoryMB}MB");
        
        $this->assertEquals(100000, $stats['rows']);
        $this->assertFileExists($this->testFile);
    }
    
    public function test_handles_various_data_types()
    {
        $sink = new FileSink($this->testFile);
        $writer = new SinkableXlsxWriter($sink);
        
        $writer->startFile([
            'String',
            'Integer', 
            'Float',
            'Boolean',
            'Null',
            'Leading Zero',
            'Special Chars',
            'Whitespace'
        ]);
        
        $testData = [
            ['Normal text', 123, 45.67, true, null, '00123', 'Test & <special>', '  spaces  '],
            ['', -456, -89.12, false, null, '00000', 'Quote"Test\'', "\ttab\t"],
            ['Very long text ' . str_repeat('x', 1000), 0, 0.0, 1, '', '01', '<>&"\'', "\n\r"],
        ];
        
        foreach ($testData as $row) {
            $writer->writeRow($row);
        }
        
        $stats = $writer->finishFile();
        
        $this->assertEquals(3, $stats['rows']);
        $this->assertFileExists($this->testFile);
        $this->assertGreaterThan(0, filesize($this->testFile));
    }
    
    public function test_performance_with_different_settings()
    {
        $results = [];
        
        // Test different compression levels
        foreach ([1, 6, 9] as $compressionLevel) {
            $testFile = sys_get_temp_dir() . "/perf_test_comp_{$compressionLevel}.xlsx";
            
            $sink = new FileSink($testFile);
            $writer = new SinkableXlsxWriter($sink);
            
            $writer->setCompressionLevel($compressionLevel)
                   ->setBufferFlushInterval(5000);
            
            $writer->startFile(['Data']);
            
            $startTime = microtime(true);
            
            for ($i = 1; $i <= 10000; $i++) {
                $writer->writeRow([str_repeat('Test data ', 50)]);
            }
            
            $stats = $writer->finishFile();
            $duration = microtime(true) - $startTime;
            
            $results[$compressionLevel] = [
                'duration' => $duration,
                'size' => filesize($testFile)
            ];
            
            unlink($testFile);
        }
        
        // Just verify all tests completed
        $this->assertArrayHasKey(1, $results);
        $this->assertArrayHasKey(6, $results);
        $this->assertArrayHasKey(9, $results);
        
        // Verify files were created
        $this->assertGreaterThan(0, $results[1]['size']);
        $this->assertGreaterThan(0, $results[9]['size']);
    }
}
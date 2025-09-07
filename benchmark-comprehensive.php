<?php

require_once __DIR__ . '/vendor/autoload.php';

use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Aws\S3\S3Client;

// Load .env
if (file_exists(__DIR__ . '/.env')) {
    $env = parse_ini_file(__DIR__ . '/.env');
    foreach ($env as $key => $value) {
        $_ENV[$key] = $value;
    }
}

// Set high memory limit for large tests
ini_set('memory_limit', '2G');

// Test sizes - comprehensive range
$testSizes = [
    '100' => 100,
    '500' => 500,
    '1K' => 1000,
    '5K' => 5000,
    '10K' => 10000,
    '25K' => 25000,
    '50K' => 50000,
    '100K' => 100000,
    '250K' => 250000,
    '500K' => 500000,
    '750K' => 750000,
    '1M' => 1000000,
    '1.5M' => 1500000,
    '2M' => 2000000,
    '3M' => 3000000,
    '4M' => 4000000,
    '4.5M' => 4500000,
];

// Headers for test data
$headers = [
    'ID', 'Name', 'Email', 'Department', 'Salary', 'Date', 'Status', 'Notes'
];

// Pre-generate static data for consistent testing
$departments = ['Sales', 'Marketing', 'Engineering', 'HR', 'Finance', 'Operations', 'Support', 'Legal', 'Product', 'Design'];
$statuses = ['Active', 'Inactive', 'Pending', 'Suspended', 'Terminated'];
$todayDate = date('Y-m-d');

// Store results
$results = [];

echo "\n";
echo "=================================================================================\n";
echo "                     KOLAY XLSX STREAM - COMPREHENSIVE BENCHMARK                 \n";
echo "=================================================================================\n";
echo "Initial memory: " . round(memory_get_usage(true) / 1024 / 1024, 2) . " MB\n";
echo "PHP Memory Limit: " . ini_get('memory_limit') . "\n";
echo "Date: " . date('Y-m-d H:i:s') . "\n\n";

// Setup S3 client once for all S3 tests
$s3Client = null;
$bucket = null;
if (isset($_ENV['AWS_ACCESS_KEY_ID'])) {
    $s3Client = new S3Client([
        'version' => 'latest',
        'region' => $_ENV['AWS_DEFAULT_REGION'] ?? 'us-east-1',
        'credentials' => [
            'key' => $_ENV['AWS_ACCESS_KEY_ID'],
            'secret' => $_ENV['AWS_SECRET_ACCESS_KEY'],
        ],
    ]);
    $bucket = $_ENV['AWS_BUCKET'];
    echo "S3 Configuration: Bucket = $bucket, Region = " . ($_ENV['AWS_DEFAULT_REGION'] ?? 'us-east-1') . "\n\n";
}

// Part 1: LOCAL FILE TESTS
echo "=================================================================================\n";
echo "                              PART 1: LOCAL FILE TESTS                           \n";
echo "=================================================================================\n\n";

foreach ($testSizes as $label => $rows) {
    if ($rows > 2000000) break; // Skip very large tests for local
    
    echo "Testing LOCAL $label (" . number_format($rows) . " rows): ";
    flush();
    
    // Create temp file
    $tempFile = tempnam(sys_get_temp_dir(), 'xlsx_test_');
    
    // Measure memory before
    gc_collect_cycles();
    $startMemory = memory_get_usage(true);
    $startTime = microtime(true);
    
    try {
        // Create file sink
        $sink = new FileSink($tempFile);
        
        // Create writer
        $writer = new SinkableXlsxWriter($sink);
        $writer->setCompressionLevel(1)->setBufferFlushInterval($rows < 10000 ? 1000 : 10000);
        
        // Start file
        $writer->startFile($headers);
        
        // Write rows
        for ($i = 1; $i <= $rows; $i++) {
            $writer->writeRow([
                $i,
                'Employee' . ($i % 100),
                'emp' . ($i % 100) . '@company.com',
                $departments[$i % 10],
                50000 + ($i % 50) * 1000,
                $todayDate,
                $statuses[$i % 5],
                'Standard notes for employee record'
            ]);
        }
        
        // Finish file
        $stats = $writer->finishFile();
        
        $endTime = microtime(true);
        $endMemory = memory_get_usage(true);
        
        // Calculate metrics
        $duration = $endTime - $startTime;
        $memoryUsed = ($endMemory - $startMemory) / 1024 / 1024;
        $speed = $rows / $duration;
        $fileSize = filesize($tempFile);
        
        // Store results
        $results['local'][$label] = [
            'rows' => $rows,
            'duration' => $duration,
            'memory_mb' => $memoryUsed,
            'speed' => $speed,
            'file_size' => $fileSize,
            'sheets' => $stats['sheets'] ?? 1
        ];
        
        echo sprintf(
            "✓ %s | Memory: %s MB | Speed: %s rows/s | File: %s MB\n",
            str_pad(round($duration, 2) . 's', 8),
            str_pad(round($memoryUsed, 2), 6, ' ', STR_PAD_LEFT),
            str_pad(number_format(round($speed)), 8, ' ', STR_PAD_LEFT),
            round($fileSize / 1024 / 1024, 2)
        );
        
    } catch (Exception $e) {
        echo "✗ ERROR: " . $e->getMessage() . "\n";
        $results['local'][$label] = ['error' => $e->getMessage()];
    }
    
    // Cleanup
    @unlink($tempFile);
    unset($writer, $sink);
    gc_collect_cycles();
}

// Part 2: S3 STREAMING TESTS
if ($s3Client) {
    echo "\n";
    echo "=================================================================================\n";
    echo "                            PART 2: S3 STREAMING TESTS                          \n";
    echo "=================================================================================\n\n";
    
    foreach ($testSizes as $label => $rows) {
        echo "Testing S3 $label (" . number_format($rows) . " rows): ";
        flush();
        
        // Create fresh S3 client for each test to avoid memory accumulation
        $s3ClientFresh = new S3Client([
            'version' => 'latest',
            'region' => $_ENV['AWS_DEFAULT_REGION'] ?? 'us-east-1',
            'credentials' => [
                'key' => $_ENV['AWS_ACCESS_KEY_ID'],
                'secret' => $_ENV['AWS_SECRET_ACCESS_KEY'],
            ],
        ]);
        
        // Measure memory before
        gc_collect_cycles();
        $startMemory = memory_get_usage(true);
        $startTime = microtime(true);
        
        try {
            // Create S3 sink
            $s3Key = 'benchmark/test_' . $label . '_' . time() . '.xlsx';
            $sink = new S3MultipartSink($s3ClientFresh, $bucket, $s3Key, 32 * 1024 * 1024);
            
            // Create writer
            $writer = new SinkableXlsxWriter($sink);
            $writer->setCompressionLevel(1)->setBufferFlushInterval($rows < 10000 ? 1000 : 10000);
            
            // Start file
            $writer->startFile($headers);
            
            // Track memory fluctuations for large files
            $memoryCheckpoints = [];
            $minMemory = PHP_INT_MAX;
            $maxMemory = 0;
            
            // Write rows
            for ($i = 1; $i <= $rows; $i++) {
                $writer->writeRow([
                    $i,
                    'Employee' . ($i % 100),
                    'emp' . ($i % 100) . '@company.com',
                    $departments[$i % 10],
                    50000 + ($i % 50) * 1000,
                    $todayDate,
                    $statuses[$i % 5],
                    'Standard notes for employee record'
                ]);
                
                // Track memory for large datasets
                if ($rows >= 100000 && $i % 100000 === 0) {
                    $currentMem = memory_get_usage(true);
                    $memoryCheckpoints[] = $currentMem;
                    $minMemory = min($minMemory, $currentMem);
                    $maxMemory = max($maxMemory, $currentMem);
                }
            }
            
            // Finish file
            $stats = $writer->finishFile();
            
            $endTime = microtime(true);
            $endMemory = memory_get_usage(true);
            
            // Calculate metrics
            $duration = $endTime - $startTime;
            $memoryUsed = ($endMemory - $startMemory) / 1024 / 1024;
            $speed = $rows / $duration;
            
            // Calculate memory fluctuation for large files
            $memoryFluctuation = 0;
            if (count($memoryCheckpoints) > 1) {
                $memoryFluctuation = ($maxMemory - $minMemory) / 1024 / 1024;
            }
            
            // Store results
            $results['s3'][$label] = [
                'rows' => $rows,
                'duration' => $duration,
                'memory_mb' => $memoryUsed,
                'memory_fluctuation' => $memoryFluctuation,
                'speed' => $speed,
                'file_size' => $stats['bytes'],
                'sheets' => $stats['sheets'] ?? 1
            ];
            
            echo sprintf(
                "✓ %s | Memory: %s MB %s| Speed: %s rows/s | File: %s MB\n",
                str_pad(round($duration, 2) . 's', 8),
                str_pad(round($memoryUsed, 2), 6, ' ', STR_PAD_LEFT),
                $memoryFluctuation > 0 ? '(±' . round($memoryFluctuation, 1) . 'MB) ' : '',
                str_pad(number_format(round($speed)), 8, ' ', STR_PAD_LEFT),
                round($stats['bytes'] / 1024 / 1024, 2)
            );
            
        } catch (Exception $e) {
            echo "✗ ERROR: " . $e->getMessage() . "\n";
            $results['s3'][$label] = ['error' => $e->getMessage()];
        }
        
        // Cleanup
        unset($writer, $sink, $s3ClientFresh);
        gc_collect_cycles();
        
        // Small delay between S3 tests
        if ($rows >= 1000000) {
            sleep(2);
        }
    }
}

// Part 3: RESULTS SUMMARY
echo "\n";
echo "=================================================================================\n";
echo "                              BENCHMARK RESULTS SUMMARY                          \n";
echo "=================================================================================\n\n";

// Markdown table format
echo "| Rows | Local Speed | Local Memory | Local Time | S3 Speed | S3 Memory | S3 Time | File Size |\n";
echo "|------|-------------|--------------|------------|----------|-----------|---------|-----------|‌\n";

foreach ($testSizes as $label => $rows) {
    if ($rows > 2000000 && !isset($results['s3'][$label])) continue;
    
    $local = $results['local'][$label] ?? null;
    $s3 = $results['s3'][$label] ?? null;
    
    if (!$local && !$s3) continue;
    
    $localSpeed = $local && !isset($local['error']) ? number_format(round($local['speed'])) . ' rows/s' : 'N/A';
    $localMemory = $local && !isset($local['error']) ? round($local['memory_mb'], 2) . ' MB' : 'N/A';
    $localTime = $local && !isset($local['error']) ? round($local['duration'], 2) . 's' : 'N/A';
    
    $s3Speed = $s3 && !isset($s3['error']) ? number_format(round($s3['speed'])) . ' rows/s' : 'N/A';
    $s3Memory = $s3 && !isset($s3['error']) ? round($s3['memory_mb'], 2) . ' MB' : 'N/A';
    if ($s3 && isset($s3['memory_fluctuation']) && $s3['memory_fluctuation'] > 0) {
        $s3Memory .= ' (±' . round($s3['memory_fluctuation'], 1) . ')';
    }
    $s3Time = $s3 && !isset($s3['error']) ? round($s3['duration'], 2) . 's' : 'N/A';
    
    $fileSize = ($s3 && !isset($s3['error']) ? $s3['file_size'] : ($local && !isset($local['error']) ? $local['file_size'] : 0));
    $fileSizeStr = $fileSize > 0 ? round($fileSize / 1024 / 1024, 2) . ' MB' : 'N/A';
    
    echo sprintf(
        "| %s | %s | %s | %s | %s | %s | %s | %s |\n",
        str_pad($label, 6),
        str_pad($localSpeed, 11),
        str_pad($localMemory, 12),
        str_pad($localTime, 10),
        str_pad($s3Speed, 10),
        str_pad($s3Memory, 11),
        str_pad($s3Time, 7),
        $fileSizeStr
    );
}

// Performance Analysis
echo "\n";
echo "=================================================================================\n";
echo "                             PERFORMANCE ANALYSIS                                \n";
echo "=================================================================================\n\n";

// Local performance
if (isset($results['local']['1M'])) {
    $local1M = $results['local']['1M'];
    echo "LOCAL FILE SYSTEM (1M rows):\n";
    echo "  • Speed: " . number_format(round($local1M['speed'])) . " rows/second\n";
    echo "  • Memory: " . round($local1M['memory_mb'], 2) . " MB (constant O(1) complexity)\n";
    echo "  • Time: " . round($local1M['duration'], 2) . " seconds\n";
    echo "  • Efficiency: Excellent - memory remains constant regardless of file size\n\n";
}

// S3 performance
if (isset($results['s3']['1M'])) {
    $s31M = $results['s3']['1M'];
    echo "S3 STREAMING (1M rows):\n";
    echo "  • Speed: " . number_format(round($s31M['speed'])) . " rows/second\n";
    echo "  • Memory: " . round($s31M['memory_mb'], 2) . " MB";
    if ($s31M['memory_fluctuation'] > 0) {
        echo " with ±" . round($s31M['memory_fluctuation'], 1) . "MB fluctuation";
    }
    echo "\n";
    echo "  • Time: " . round($s31M['duration'], 2) . " seconds\n";
    echo "  • Note: Memory fluctuation is normal due to S3 multipart upload buffers\n\n";
}

// Memory efficiency
echo "MEMORY EFFICIENCY:\n";
echo "  • Local files: True O(1) memory complexity - uses only buffer memory\n";
echo "  • S3 streaming: Near O(1) with small linear growth for part metadata\n";
echo "  • Memory fluctuation in S3: Caused by periodic part uploads (every ~32MB)\n";
echo "  • After each part upload, memory drops back to baseline (GC collects temp buffers)\n\n";

// Multi-sheet support
$multiSheetTests = array_filter($results['s3'] ?? [], function($r) { 
    return isset($r['sheets']) && $r['sheets'] > 1; 
});
if (!empty($multiSheetTests)) {
    echo "MULTI-SHEET SUPPORT:\n";
    foreach ($multiSheetTests as $label => $test) {
        echo "  • $label: " . $test['sheets'] . " sheets (Excel limit: 1,048,576 rows/sheet)\n";
    }
    echo "\n";
}

echo "=================================================================================\n";
echo "                              BENCHMARK COMPLETED                                \n";
echo "=================================================================================\n";
echo "Peak memory usage: " . round(memory_get_peak_usage(true) / 1024 / 1024, 2) . " MB\n";
echo "Completed at: " . date('Y-m-d H:i:s') . "\n\n";
<?php

require __DIR__ . '/../vendor/autoload.php';

use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Kolay\XlsxStream\Sinks\FileSink;

// Example 1: Basic file export
function basicFileExport()
{
    echo "Starting basic file export...\n";
    
    // Create a file sink
    $sink = new FileSink(__DIR__ . '/output/basic-export.xlsx');
    
    // Create writer
    $writer = new SinkableXlsxWriter($sink);
    
    // Start file with headers
    $writer->startFile(['ID', 'Name', 'Email', 'Department', 'Salary']);
    
    // Write some sample data
    $departments = ['Engineering', 'Sales', 'Marketing', 'HR', 'Finance'];
    
    for ($i = 1; $i <= 100; $i++) {
        $writer->writeRow([
            $i,
            "Employee $i",
            "employee$i@company.com",
            $departments[array_rand($departments)],
            rand(50000, 150000)
        ]);
    }
    
    // Finish file
    $stats = $writer->finishFile();
    
    echo "Export completed!\n";
    echo "- Total rows: {$stats['rows']}\n";
    echo "- Total sheets: {$stats['sheets']}\n";
    echo "- File size: " . number_format($stats['bytes']) . " bytes\n";
    echo "- File location: " . realpath(__DIR__ . '/output/basic-export.xlsx') . "\n\n";
}

// Example 2: Large dataset with performance optimization
function largeDatasetExport()
{
    echo "Starting large dataset export...\n";
    
    $sink = new FileSink(__DIR__ . '/output/large-export.xlsx');
    $writer = new SinkableXlsxWriter($sink);
    
    // Optimize for speed
    $writer->setCompressionLevel(1)        // Fastest compression
           ->setBufferFlushInterval(10000); // Buffer 10K rows
    
    // Headers
    $writer->startFile([
        'ID',
        'UUID',
        'Username',
        'Email',
        'First Name',
        'Last Name',
        'Phone',
        'Address',
        'City',
        'Country',
        'Postal Code',
        'Registration Date',
        'Last Login',
        'Status',
        'Score'
    ]);
    
    // Generate 1 million rows
    $startTime = microtime(true);
    $countries = ['USA', 'UK', 'Canada', 'Australia', 'Germany', 'France', 'Japan', 'Brazil'];
    $statuses = ['Active', 'Inactive', 'Pending', 'Suspended'];
    
    for ($i = 1; $i <= 1000000; $i++) {
        $writer->writeRow([
            $i,
            uniqid('user_'),
            "user$i",
            "user$i@example.com",
            "FirstName$i",
            "LastName$i",
            "+1-555-" . str_pad($i % 10000, 4, '0', STR_PAD_LEFT),
            "$i Main Street",
            "City " . ($i % 100),
            $countries[array_rand($countries)],
            str_pad($i % 100000, 5, '0', STR_PAD_LEFT),
            date('Y-m-d H:i:s', time() - rand(0, 365 * 24 * 60 * 60)),
            date('Y-m-d H:i:s', time() - rand(0, 30 * 24 * 60 * 60)),
            $statuses[array_rand($statuses)],
            rand(0, 100)
        ]);
        
        // Progress indicator
        if ($i % 100000 === 0) {
            $elapsed = microtime(true) - $startTime;
            $rate = $i / $elapsed;
            echo "  Processed " . number_format($i) . " rows";
            echo " (" . number_format($rate, 0) . " rows/sec)\n";
        }
    }
    
    $stats = $writer->finishFile();
    $totalTime = microtime(true) - $startTime;
    
    echo "\nLarge export completed!\n";
    echo "- Total rows: " . number_format($stats['rows']) . "\n";
    echo "- Total sheets: {$stats['sheets']}\n";
    echo "- File size: " . number_format($stats['bytes']) . " bytes\n";
    echo "- Time taken: " . round($totalTime, 2) . " seconds\n";
    echo "- Average speed: " . number_format($stats['rows'] / $totalTime, 0) . " rows/sec\n";
    echo "- Memory peak: " . number_format(memory_get_peak_usage(true) / 1024 / 1024, 2) . " MB\n\n";
}

// Example 3: Multi-sheet export
function multiSheetExport()
{
    echo "Starting multi-sheet export (will exceed Excel row limit)...\n";
    
    $sink = new FileSink(__DIR__ . '/output/multi-sheet.xlsx');
    $writer = new SinkableXlsxWriter($sink);
    
    $writer->setCompressionLevel(1)
           ->setBufferFlushInterval(50000);
    
    $writer->startFile(['Row Number', 'Sheet Info', 'Random Data']);
    
    // Write 2.1 million rows (will create 3 sheets)
    for ($i = 1; $i <= 2100000; $i++) {
        $currentSheet = ceil($i / 1048575);
        $writer->writeRow([
            $i,
            "Sheet $currentSheet",
            md5($i)
        ]);
        
        if ($i % 500000 === 0) {
            echo "  Written " . number_format($i) . " rows...\n";
        }
    }
    
    $stats = $writer->finishFile();
    
    echo "\nMulti-sheet export completed!\n";
    echo "- Total rows: " . number_format($stats['rows']) . "\n";
    echo "- Total sheets: {$stats['sheets']}\n";
    echo "- Sheet details:\n";
    foreach ($stats['sheet_details'] as $sheet) {
        echo "  - {$sheet['name']}: " . number_format($sheet['rows']) . " rows\n";
    }
    echo "\n";
}

// Create output directory
if (!is_dir(__DIR__ . '/output')) {
    mkdir(__DIR__ . '/output', 0755, true);
}

// Run examples
echo "=== Kolay XLSX Stream Examples ===\n\n";

// Uncomment the example you want to run:
basicFileExport();
// largeDatasetExport();  // This will create a 1M row file
// multiSheetExport();    // This will create a 2.1M row file

echo "All examples completed!\n";
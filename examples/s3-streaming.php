<?php

require __DIR__ . '/../vendor/autoload.php';

use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Aws\S3\S3Client;

/**
 * Example: Direct S3 Streaming with Zero Disk I/O
 * 
 * This example demonstrates how to stream XLSX files directly to S3
 * without using any local disk space.
 */

// Configure your AWS credentials
$awsConfig = [
    'region' => 'us-east-1', // Change to your region
    'version' => 'latest',
    'credentials' => [
        'key' => getenv('AWS_ACCESS_KEY_ID'),
        'secret' => getenv('AWS_SECRET_ACCESS_KEY'),
    ],
];

$bucketName = getenv('AWS_BUCKET') ?: 'your-bucket-name';

function exportToS3()
{
    global $awsConfig, $bucketName;
    
    echo "Starting S3 streaming export...\n";
    echo "Bucket: $bucketName\n\n";
    
    // Create S3 client
    $s3Client = new S3Client($awsConfig);
    
    // Create S3 sink with 32MB parts for optimal performance
    $filename = 'exports/users-' . date('Y-m-d-H-i-s') . '.xlsx';
    $sink = new S3MultipartSink(
        $s3Client,
        $bucketName,
        $filename,
        32 * 1024 * 1024, // 32MB parts
        [
            'ACL' => 'private',
            'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Metadata' => [
                'generated-by' => 'kolay-xlsx-stream',
                'generated-at' => date('c'),
            ]
        ]
    );
    
    // Create writer
    $writer = new SinkableXlsxWriter($sink);
    
    // Optimize for S3 streaming
    $writer->setCompressionLevel(1)        // Fast compression
           ->setBufferFlushInterval(10000); // 10K row buffer
    
    // Start file with headers
    $writer->startFile([
        'User ID',
        'Username',
        'Email',
        'Full Name',
        'Department',
        'Role',
        'Status',
        'Created Date',
        'Last Active'
    ]);
    
    // Simulate database query with chunking
    $totalUsers = 100000;
    $chunkSize = 1000;
    $departments = ['Engineering', 'Product', 'Sales', 'Marketing', 'Support', 'Operations'];
    $roles = ['Admin', 'Manager', 'Employee', 'Contractor'];
    $statuses = ['Active', 'Inactive', 'Pending'];
    
    $startTime = microtime(true);
    
    for ($offset = 0; $offset < $totalUsers; $offset += $chunkSize) {
        // Simulate chunk of users from database
        $users = [];
        for ($i = 1; $i <= $chunkSize && ($offset + $i) <= $totalUsers; $i++) {
            $userId = $offset + $i;
            $users[] = [
                'id' => $userId,
                'username' => "user_$userId",
                'email' => "user$userId@company.com",
                'full_name' => "User Name $userId",
                'department' => $departments[array_rand($departments)],
                'role' => $roles[array_rand($roles)],
                'status' => $statuses[array_rand($statuses)],
                'created_date' => date('Y-m-d', time() - rand(0, 365 * 24 * 60 * 60)),
                'last_active' => date('Y-m-d H:i:s', time() - rand(0, 7 * 24 * 60 * 60))
            ];
        }
        
        // Write users to XLSX
        foreach ($users as $user) {
            $writer->writeRow([
                $user['id'],
                $user['username'],
                $user['email'],
                $user['full_name'],
                $user['department'],
                $user['role'],
                $user['status'],
                $user['created_date'],
                $user['last_active']
            ]);
        }
        
        // Progress update
        if (($offset + $chunkSize) % 10000 === 0) {
            $progress = min($offset + $chunkSize, $totalUsers);
            $elapsed = microtime(true) - $startTime;
            $rate = $progress / $elapsed;
            echo sprintf(
                "Progress: %s / %s rows (%.1f%%) - %.0f rows/sec\n",
                number_format($progress),
                number_format($totalUsers),
                ($progress / $totalUsers) * 100,
                $rate
            );
        }
    }
    
    // Finish file and upload
    echo "\nFinalizing S3 upload...\n";
    $stats = $writer->finishFile();
    
    $totalTime = microtime(true) - $startTime;
    
    echo "\n=== S3 Export Completed ===\n";
    echo "File: s3://$bucketName/$filename\n";
    echo "Total rows: " . number_format($stats['rows']) . "\n";
    echo "Total sheets: {$stats['sheets']}\n";
    echo "File size: " . number_format($stats['bytes']) . " bytes\n";
    echo "Time taken: " . round($totalTime, 2) . " seconds\n";
    echo "Average speed: " . number_format($stats['rows'] / $totalTime, 0) . " rows/sec\n";
    echo "Memory peak: " . number_format(memory_get_peak_usage(true) / 1024 / 1024, 2) . " MB\n";
    
    // Generate presigned URL for download (optional)
    try {
        $cmd = $s3Client->getCommand('GetObject', [
            'Bucket' => $bucketName,
            'Key' => $filename
        ]);
        $request = $s3Client->createPresignedRequest($cmd, '+1 hour');
        $presignedUrl = (string) $request->getUri();
        
        echo "\nDownload URL (valid for 1 hour):\n";
        echo $presignedUrl . "\n";
    } catch (Exception $e) {
        echo "\nCould not generate presigned URL: " . $e->getMessage() . "\n";
    }
}

// Check AWS credentials
if (!getenv('AWS_ACCESS_KEY_ID') || !getenv('AWS_SECRET_ACCESS_KEY')) {
    echo "ERROR: AWS credentials not found!\n";
    echo "Please set the following environment variables:\n";
    echo "  AWS_ACCESS_KEY_ID\n";
    echo "  AWS_SECRET_ACCESS_KEY\n";
    echo "  AWS_BUCKET (optional)\n\n";
    echo "Example:\n";
    echo "  export AWS_ACCESS_KEY_ID=your_key_here\n";
    echo "  export AWS_SECRET_ACCESS_KEY=your_secret_here\n";
    echo "  export AWS_BUCKET=your-bucket-name\n";
    echo "  php " . __FILE__ . "\n";
    exit(1);
}

// Run the export
try {
    exportToS3();
} catch (Exception $e) {
    echo "\nERROR: " . $e->getMessage() . "\n";
    echo "Stack trace:\n" . $e->getTraceAsString() . "\n";
    exit(1);
}
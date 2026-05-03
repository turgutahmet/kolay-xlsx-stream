<?php

require __DIR__.'/../vendor/autoload.php';

use Aws\S3\S3Client;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Example: Direct S3 Streaming with Zero Disk I/O — showcases v2.2 styling.
 *
 * Generates a multi-sheet workbook with header styling, column formats,
 * frozen first row, auto-filter, auto column widths, and progress reporting,
 * then streams it straight to S3 using multipart upload (no temp files).
 *
 * Run with AWS env vars set, then download the file from S3 to inspect:
 *   AWS_ACCESS_KEY_ID=... AWS_SECRET_ACCESS_KEY=... AWS_BUCKET=... \
 *     AWS_DEFAULT_REGION=us-east-2 php examples/s3-streaming.php
 */

// Bootstrap config from env (or .env if vlucas/phpdotenv is in the autoloader)
if (file_exists(__DIR__.'/../.env') && class_exists(\Dotenv\Dotenv::class)) {
    \Dotenv\Dotenv::createUnsafeImmutable(__DIR__.'/..')->safeLoad();
}

$awsConfig = [
    'region' => getenv('AWS_DEFAULT_REGION') ?: 'us-east-1',
    'version' => 'latest',
    'credentials' => [
        'key' => getenv('AWS_ACCESS_KEY_ID'),
        'secret' => getenv('AWS_SECRET_ACCESS_KEY'),
    ],
];

$bucketName = getenv('AWS_BUCKET') ?: 'your-bucket-name';

if (! getenv('AWS_ACCESS_KEY_ID') || ! getenv('AWS_SECRET_ACCESS_KEY') || $bucketName === 'your-bucket-name') {
    echo "ERROR: AWS credentials not found.\n";
    echo "Set AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_BUCKET, AWS_DEFAULT_REGION.\n";
    exit(1);
}

function exportToS3(array $awsConfig, string $bucketName): void
{
    echo "Streaming styled multi-sheet workbook to S3...\n";
    echo "Bucket: {$bucketName}\n\n";

    $s3 = new S3Client($awsConfig);
    $key = 'examples/styled-'.date('Ymd-His').'.xlsx';

    $sink = new S3MultipartSink(
        $s3,
        $bucketName,
        $key,
        32 * 1024 * 1024, // 32 MB parts
        [
            'ACL' => 'private',
            'ContentDisposition' => 'attachment; filename="users-export.xlsx"',
            'Metadata' => [
                'generated-by' => 'kolay-xlsx-stream',
                'generated-at' => date('c'),
            ],
        ]
    );

    $writer = new SinkableXlsxWriter($sink);

    // ─── v2.2 styling ──────────────────────────────────────────────
    $writer
        ->setCompressionLevel(1)
        ->setBufferFlushInterval(10000)
        // header: bold white text on dark blue
        ->setHeaderStyle([
            'bold' => true,
            'fill' => '#4F81BD',
            'color' => '#FFFFFF',
        ])
        // column number formats — Excel will render these natively
        ->setColumnFormat(1, 'integer')        // ID
        ->setColumnFormat(7, 'date')           // Created Date
        ->setColumnFormat(8, 'datetime')       // Last Active
        ->setColumnFormat(9, 'currency_try')   // Salary
        ->setColumnFormat(10, 'percent')       // Performance score
        // explicit + auto column widths
        ->setAutoColumnWidth()
        ->setColumnWidths([4 => 24, 5 => 16, 6 => 14, 7 => 12, 8 => 18])
        // freeze header row + auto-filter dropdowns
        ->freezeFirstRow()
        ->enableAutoFilter();

    // ─── progress callback (v2.1) ──────────────────────────────────
    $startTime = microtime(true);
    $writer->onProgress(function (int $rows, int $bytes) use ($startTime) {
        $elapsed = microtime(true) - $startTime;
        $rate = $rows / max($elapsed, 0.001);
        echo sprintf(
            "  %s rows | %.2f MB streamed | %.0f rows/sec\n",
            number_format($rows),
            $bytes / 1024 / 1024,
            $rate
        );
    })->setProgressInterval(10000);

    // ─── sheet 1: Users (50,000 rows) ──────────────────────────────
    echo "Sheet 1: Users\n";

    $writer->startFile([
        'ID',
        'Username',
        'Email',
        'Full Name',
        'Department',
        'Role',
        'Created Date',
        'Last Active',
        'Salary',
        'Performance',
    ]);

    $departments = ['Engineering', 'Product', 'Sales', 'Marketing', 'Support', 'Operations'];
    $roles = ['Admin', 'Manager', 'Employee', 'Contractor'];

    for ($i = 1; $i <= 50_000; $i++) {
        $writer->writeRow([
            $i,                                                                         // 1: integer
            "user_{$i}",                                                                // 2: string
            "user{$i}@company.com",                                                     // 3: string
            "User Name {$i}",                                                           // 4: string
            $departments[$i % count($departments)],                                     // 5: string
            $roles[$i % count($roles)],                                                 // 6: string
            new DateTime('2026-01-01 +'.($i % 90).' days', new DateTimeZone('UTC')),    // 7: date
            new DateTime('2026-05-01 -'.($i % 7).' days', new DateTimeZone('UTC')),     // 8: datetime
            45000 + ($i % 80) * 1000,                                                   // 9: currency
            ($i % 100) / 100,                                                           // 10: percent
        ]);
    }

    // ─── sheet 2: Summary ──────────────────────────────────────────
    echo "\nSheet 2: Summary\n";

    $writer
        ->setHeaderStyle(['bold' => true, 'fill' => '#9BBB59', 'color' => '#FFFFFF'])
        ->setColumnFormat(2, 'integer')
        ->newSheet('Summary', ['Department', 'Headcount']);

    foreach ($departments as $dept) {
        $writer->writeRow([$dept, rand(500, 15000)]);
    }

    // ─── finalize ──────────────────────────────────────────────────
    $stats = $writer->finishFile();
    $totalTime = microtime(true) - $startTime;

    echo "\n=== S3 Export Completed ===\n";
    echo "File:       s3://{$bucketName}/{$key}\n";
    echo "Sheets:     {$stats['sheets']}\n";
    foreach ($stats['sheet_details'] as $sheet) {
        echo sprintf("  • %-12s %s rows\n", $sheet['name'], number_format($sheet['rows']));
    }
    echo 'Total rows: '.number_format($stats['rows'])."\n";
    echo 'File size:  '.number_format($stats['bytes'])." bytes\n";
    echo 'Time:       '.round($totalTime, 2)." s\n";
    echo 'Speed:      '.number_format($stats['rows'] / $totalTime, 0)." rows/sec\n";
    echo 'Peak mem:   '.number_format(memory_get_peak_usage(true) / 1024 / 1024, 2)." MB\n";

    // Presigned URL for quick browser download
    try {
        $cmd = $s3->getCommand('GetObject', ['Bucket' => $bucketName, 'Key' => $key]);
        $url = (string) $s3->createPresignedRequest($cmd, '+1 hour')->getUri();
        echo "\nDownload (valid 1 hour):\n{$url}\n";
    } catch (\Throwable $e) {
        echo "\nCould not generate presigned URL: ".$e->getMessage()."\n";
    }
}

try {
    exportToS3($awsConfig, $bucketName);
} catch (\Throwable $e) {
    echo "\nERROR: ".$e->getMessage()."\n";
    echo $e->getTraceAsString()."\n";
    exit(1);
}

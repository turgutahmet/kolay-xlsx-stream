<?php

require __DIR__.'/../vendor/autoload.php';

use Aws\S3\S3Client;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Example: Direct S3 Streaming with Zero Disk I/O — full v2.2 feature tour.
 *
 * Generates a 5-sheet workbook covering every styling primitive and edge
 * case the package exposes, then streams it straight to S3 using multipart
 * upload (no temp files). The file is left in place so you can download
 * and inspect it manually.
 *
 *   AWS_ACCESS_KEY_ID=... AWS_SECRET_ACCESS_KEY=... AWS_BUCKET=... \
 *     AWS_DEFAULT_REGION=us-east-2 php examples/s3-streaming.php
 */

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
    echo "Streaming styling-tour workbook to S3...\n";
    echo "Bucket: {$bucketName}\n\n";

    $s3 = new S3Client($awsConfig);
    $key = 'examples/styling-tour-'.date('Ymd-His').'.xlsx';

    $sink = new S3MultipartSink(
        $s3,
        $bucketName,
        $key,
        32 * 1024 * 1024,
        [
            'ACL' => 'private',
            'ContentDisposition' => 'attachment; filename="styling-tour.xlsx"',
            'Metadata' => [
                'generated-by' => 'kolay-xlsx-stream',
                'generated-at' => date('c'),
                'package-version' => '2.2.0',
            ],
        ]
    );

    $writer = new SinkableXlsxWriter($sink);
    $writer
        ->setCompressionLevel(1)
        ->setBufferFlushInterval(10000);

    $startTime = microtime(true);
    $writer->onProgress(function (int $rows, int $bytes) use ($startTime) {
        $rate = $rows / max(microtime(true) - $startTime, 0.001);
        echo sprintf(
            "    %s rows | %.2f MB | %s rows/sec\n",
            number_format($rows),
            $bytes / 1024 / 1024,
            number_format(round($rate))
        );
    })->setProgressInterval(10000);

    // ════════════════════════════════════════════════════════════════
    // Sheet 1 — Sales (large, all major formats together)
    // ════════════════════════════════════════════════════════════════
    echo "[1/5] Sheet 'Sales' — 50K rows, full styling\n";

    $writer
        ->setHeaderStyle([
            'bold' => true,
            'fill' => '#4F81BD',
            'color' => '#FFFFFF',
            'size' => 12,
        ])
        ->setColumnFormat(1, 'integer')
        ->setColumnFormat(5, 'currency_try')   // ₺
        ->setColumnFormat(6, 'currency_usd')   // $
        ->setColumnFormat(7, 'percent')
        ->setColumnFormat(8, 'date')
        ->setColumnFormat(9, 'datetime')
        ->setAutoColumnWidth()                 // format-aware widths now
        ->setColumnWidths([2 => 22, 4 => 18])  // override Customer + Region wider
        ->freezeFirstRow()
        ->enableAutoFilter();

    $writer->startFile([
        'Order ID',
        'Customer',
        'Product Code',
        'Region',
        'Price (TRY)',
        'Price (USD)',
        'Discount',
        'Order Date',
        'Created At',
    ]);

    $regions = ['Europe', 'North America', 'Asia-Pacific', 'LATAM', 'MENA'];
    $products = ['SKU-A001', 'SKU-A002', 'SKU-B100', 'SKU-B101', 'SKU-C500'];

    for ($i = 1; $i <= 50_000; $i++) {
        $writer->writeRow([
            $i,
            "Customer Co. #{$i}",
            $products[$i % 5],
            $regions[$i % 5],
            999.99 + ($i % 1000) * 12.5,             // ₺
            (999.99 + ($i % 1000) * 12.5) / 33.5,    // USD
            ($i % 100) / 100,                        // 0.00 - 0.99
            new DateTime('2026-01-01 +'.($i % 365).' days', new DateTimeZone('UTC')),
            new DateTime('2026-05-03 -'.($i % 30).' hours', new DateTimeZone('UTC')),
        ]);
    }

    // ════════════════════════════════════════════════════════════════
    // Sheet 2 — Edge cases (data-quality stress test)
    // ════════════════════════════════════════════════════════════════
    echo "[2/5] Sheet 'Edge Cases' — special values\n";

    $writer
        ->clearColumnFormats()
        ->setHeaderStyle(['bold' => true, 'fill' => '#C0504D', 'color' => '#FFFFFF'])
        ->setColumnFormat(1, 'integer')
        ->setAutoColumnWidth()
        ->newSheet('Edge Cases', [
            'Row',
            'Empty',
            'Null',
            'Boolean True',
            'Boolean False',
            'Big Int (20 digits)',
            'Phone (+prefix)',
            'Leading Zero',
            'Special XML Chars',
            'Multi-byte UTF-8',
            'Whitespace',
            'Negative',
        ]);

    $writer->writeRow([1, '', null, true, false, '12345678901234567890', '+905551234567', '00123', '<tag>&"\'', 'Türkçe → İŞŞL', '   leading', -42]);
    $writer->writeRow([2, '', null, false, true, '99999999999999999999', '+447911123456', '00007', 'a < b && c > d', '日本語テスト', 'trailing   ', -3.14]);
    $writer->writeRow([3, null, null, true, true, '1234567890', '+1234567890', '0001', '"quoted"', 'العربية', "\ttab\t", 0]);

    // ════════════════════════════════════════════════════════════════
    // Sheet 3 — Manual column widths only (no auto)
    // ════════════════════════════════════════════════════════════════
    echo "[3/5] Sheet 'Manual Widths'\n";

    $writer
        ->clearColumnFormats()
        ->setHeaderStyle(['bold' => true, 'fill' => '#9BBB59', 'color' => '#FFFFFF'])
        ->setAutoColumnWidth(false)
        ->setColumnWidths([1 => 6, 2 => 40, 3 => 12, 4 => 30])
        ->newSheet('Manual Widths', ['#', 'Long Description Column', 'Code', 'Notes']);

    for ($i = 1; $i <= 50; $i++) {
        $writer->writeRow([
            $i,
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur '.$i,
            'CODE-'.str_pad($i, 4, '0', STR_PAD_LEFT),
            'Note '.$i,
        ]);
    }

    // ════════════════════════════════════════════════════════════════
    // Sheet 4 — Number format presets gallery
    // ════════════════════════════════════════════════════════════════
    echo "[4/5] Sheet 'Format Gallery'\n";

    $writer
        ->clearColumnFormats()
        ->setHeaderStyle(['bold' => true, 'fill' => '#8064A2', 'color' => '#FFFFFF'])
        ->setColumnFormat(2, 'integer')
        ->setColumnFormat(3, 'decimal')
        ->setColumnFormat(4, 'percent')
        ->setColumnFormat(5, 'currency_try')
        ->setColumnFormat(6, 'currency_usd')
        ->setColumnFormat(7, 'currency_eur')
        ->setColumnFormat(8, 'currency_gbp')
        ->setColumnFormat(9, 'date')
        ->setColumnFormat(10, 'datetime')
        ->setColumnFormat(11, 'time')
        ->setColumnFormat(12, '0.000000') // raw custom format code
        ->setAutoColumnWidth()
        ->newSheet('Format Gallery', [
            'Label',
            'integer',
            'decimal',
            'percent',
            'currency_try',
            'currency_usd',
            'currency_eur',
            'currency_gbp',
            'date',
            'datetime',
            'time',
            'custom 0.000000',
        ]);

    $samples = [
        ['Small', 1, 1.5, 0.01, 9.99, 9.99, 9.99, 9.99],
        ['Medium', 12345, 1234.56, 0.5, 99999.99, 99999.99, 99999.99, 99999.99],
        ['Large', 1234567890, 9876543.21, 0.999, 1234567.89, 1234567.89, 1234567.89, 1234567.89],
        ['Negative', -42, -3.14, -0.05, -250, -250, -250, -250],
        ['Zero', 0, 0, 0, 0, 0, 0, 0],
    ];
    $time = new DateTime('2026-05-03 14:23:45', new DateTimeZone('UTC'));
    foreach ($samples as $row) {
        $writer->writeRow(array_merge($row, [$time, $time, $time, 0.123456789]));
    }

    // ════════════════════════════════════════════════════════════════
    // Sheet 5 — Frozen pane + many columns wide layout
    // ════════════════════════════════════════════════════════════════
    echo "[5/5] Sheet 'Wide Layout'\n";

    $writer
        ->clearColumnFormats()
        ->setHeaderStyle(['bold' => true, 'fill' => '#1F497D', 'color' => '#FFFFFF'])
        ->freezeRowsAndColumns(rows: 1, columns: 2)
        ->enableAutoFilter()
        ->setAutoColumnWidth()
        ->setColumnFormat(1, 'integer')
        ->setColumnFormat(3, 'currency_try')
        ->setColumnFormat(4, 'currency_try')
        ->setColumnFormat(5, 'currency_try')
        ->setColumnFormat(6, 'currency_try')
        ->setColumnFormat(7, 'currency_try')
        ->setColumnFormat(8, 'percent')
        ->setColumnFormat(9, 'datetime')
        ->newSheet('Wide Layout', [
            'ID', 'Name',
            'Q1 Sales', 'Q2 Sales', 'Q3 Sales', 'Q4 Sales',
            'Total', 'Growth %',
            'Last Update',
        ]);

    for ($i = 1; $i <= 200; $i++) {
        $q1 = 50000 + ($i * 137) % 80000;
        $q2 = 60000 + ($i * 211) % 90000;
        $q3 = 55000 + ($i * 173) % 85000;
        $q4 = 70000 + ($i * 191) % 100000;
        $total = $q1 + $q2 + $q3 + $q4;
        $writer->writeRow([
            $i,
            "Account-".str_pad($i, 4, '0', STR_PAD_LEFT),
            $q1, $q2, $q3, $q4,
            $total,
            ($i % 25 - 12) / 100,
            new DateTime('2026-05-01 +'.($i % 48).' hours', new DateTimeZone('UTC')),
        ]);
    }

    // ════════════════════════════════════════════════════════════════
    // Finalize
    // ════════════════════════════════════════════════════════════════
    $stats = $writer->finishFile();
    $totalTime = microtime(true) - $startTime;

    echo "\n=== S3 Export Completed ===\n";
    echo "File:       s3://{$bucketName}/{$key}\n";
    echo "Sheets:     {$stats['sheets']}\n";
    foreach ($stats['sheet_details'] as $sheet) {
        echo sprintf("  • %-16s %s rows\n", $sheet['name'], number_format($sheet['rows']));
    }
    echo 'Total rows: '.number_format($stats['rows'])."\n";
    echo 'File size:  '.number_format($stats['bytes'])." bytes\n";
    echo 'Time:       '.round($totalTime, 2)." s\n";
    echo 'Speed:      '.number_format($stats['rows'] / $totalTime, 0)." rows/sec\n";
    echo 'Peak mem:   '.number_format(memory_get_peak_usage(true) / 1024 / 1024, 2)." MB\n";

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

<?php
/**
 * Streaming reader benchmark — companion to benchmark-comprehensive.php.
 *
 * Mirrors the write benchmark's workload exactly (8 columns, mixed
 * types, identical row shape) so write- and read-side rows/sec
 * comparisons are apples-to-apples and stable across versions.
 *
 * For each row size:
 *   1. Write a temp XLSX with the same payload the write benchmark uses
 *   2. Read it back via StreamingXlsxReader::fromFile, time the full
 *      sequential scan, record peak RAM
 *   3. Repeat from S3 if AWS credentials are present (cold-cache GET)
 *
 * Run:        php benchmark-read.php
 * Quick run:  php benchmark-read.php 100000
 *             (caps the largest row count tested at the given limit)
 */

require_once __DIR__.'/vendor/autoload.php';

use Aws\S3\S3Client;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

if (file_exists(__DIR__.'/.env')) {
    foreach (parse_ini_file(__DIR__.'/.env') as $key => $value) {
        $_ENV[$key] = $value;
    }
}

ini_set('memory_limit', '2G');

$cap = isset($argv[1]) ? (int) $argv[1] : PHP_INT_MAX;

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

$headers = ['ID', 'Name', 'Email', 'Department', 'Salary', 'Date', 'Status', 'Notes'];
$departments = ['Sales', 'Marketing', 'Engineering', 'HR', 'Finance', 'Operations', 'Support', 'Legal', 'Product', 'Design'];
$statuses = ['Active', 'Inactive', 'Pending', 'Suspended', 'Terminated'];
$today = date('Y-m-d');

$results = ['local' => [], 's3' => []];

echo "\n";
echo "=================================================================================\n";
echo "                  KOLAY XLSX STREAM — READ BENCHMARK (v3.0+)                     \n";
echo "=================================================================================\n";
echo 'PHP: '.PHP_VERSION.'  |  Date: '.date('Y-m-d H:i:s')."\n";
echo "Cap: ".($cap === PHP_INT_MAX ? 'none' : number_format($cap).' rows')."\n\n";

$s3Client = null;
$bucket = null;
if (! empty($_ENV['AWS_ACCESS_KEY_ID']) && ! empty($_ENV['AWS_BUCKET'])) {
    $s3Client = new S3Client([
        'version' => 'latest',
        'region' => $_ENV['AWS_DEFAULT_REGION'] ?? 'us-east-1',
        'credentials' => [
            'key' => $_ENV['AWS_ACCESS_KEY_ID'],
            'secret' => $_ENV['AWS_SECRET_ACCESS_KEY'],
        ],
    ]);
    $bucket = $_ENV['AWS_BUCKET'];
    echo "S3: bucket={$bucket}  region=".($_ENV['AWS_DEFAULT_REGION'] ?? 'us-east-1')."\n\n";
}

// ---------------------------------------------------------------------------
// Part 1 — Local reads
// ---------------------------------------------------------------------------

echo "=================================================================================\n";
echo "                            PART 1: LOCAL READS                                  \n";
echo "=================================================================================\n\n";

foreach ($testSizes as $label => $rows) {
    if ($rows > $cap) {
        break;
    }
    if ($rows > 2000000) {
        // Match benchmark-comprehensive.php: skip the very largest sizes
        // for the local-only path (too slow to be practical here).
        break;
    }

    $tempFile = tempnam(sys_get_temp_dir(), 'kxs_read_');

    echo "Local {$label} (".number_format($rows)." rows): write... ";
    flush();

    writeWorkload($tempFile, $rows, $headers, $departments, $statuses, $today);
    $fileSize = filesize($tempFile);

    echo 'read... ';
    flush();

    gc_collect_cycles();
    memory_reset_peak_usage();
    $rowsRead = 0;
    $start = microtime(true);

    $reader = StreamingXlsxReader::fromFile($tempFile);
    // Workbooks above ~1.05 M rows auto-split across multiple sheets
    // (Excel's per-sheet limit is 1,048,576). Read every sheet so the
    // measured throughput reflects the full dataset, not just sheet 1.
    foreach ($reader->sheets() as $sheet) {
        $reader->onSheet($sheet['name']);
        foreach ($reader->rows() as $_) {
            $rowsRead++;
        }
    }
    $reader->close();

    $duration = microtime(true) - $start;
    $peakBytes = memory_get_peak_usage(true);

    $results['local'][$label] = [
        'rows' => $rows,
        'rows_read' => $rowsRead,
        'duration' => $duration,
        'speed' => $rowsRead / max($duration, 0.0001),
        'peak_mb' => $peakBytes / 1024 / 1024,
        'file_mb' => $fileSize / 1024 / 1024,
    ];

    echo sprintf(
        "%s | %s rows/s | peak %s MB | file %s MB\n",
        str_pad(round($duration, 2).'s', 8),
        str_pad(number_format(round($rowsRead / max($duration, 0.0001))), 9, ' ', STR_PAD_LEFT),
        str_pad(round($peakBytes / 1024 / 1024, 1), 5, ' ', STR_PAD_LEFT),
        round($fileSize / 1024 / 1024, 2)
    );

    @unlink($tempFile);
    unset($reader);
    gc_collect_cycles();
}

// ---------------------------------------------------------------------------
// Part 2 — S3 reads (cold cache)
// ---------------------------------------------------------------------------

if ($s3Client) {
    echo "\n=================================================================================\n";
    echo "                            PART 2: S3 STREAMING READS                           \n";
    echo "=================================================================================\n\n";

    foreach ($testSizes as $label => $rows) {
        if ($rows > $cap) {
            break;
        }

        $key = 'benchmarks/read/v3-'.$label.'-'.uniqid().'.xlsx';

        echo "S3 {$label} (".number_format($rows)." rows): upload... ";
        flush();

        $sink = new S3MultipartSink($s3Client, $bucket, $key);
        writeWorkloadToSink($sink, $rows, $headers, $departments, $statuses, $today);

        // HEAD to get file size (and confirm upload).
        $head = $s3Client->headObject(['Bucket' => $bucket, 'Key' => $key]);
        $fileSize = (int) $head['ContentLength'];

        echo 'read... ';
        flush();

        gc_collect_cycles();
        memory_reset_peak_usage();
        $rowsRead = 0;
        $start = microtime(true);

        $reader = StreamingXlsxReader::fromS3($s3Client, $bucket, $key);
        foreach ($reader->sheets() as $sheet) {
            $reader->onSheet($sheet['name']);
            foreach ($reader->rows() as $_) {
                $rowsRead++;
            }
        }
        $reader->close();

        $duration = microtime(true) - $start;
        $peakBytes = memory_get_peak_usage(true);

        $results['s3'][$label] = [
            'rows' => $rows,
            'rows_read' => $rowsRead,
            'duration' => $duration,
            'speed' => $rowsRead / max($duration, 0.0001),
            'peak_mb' => $peakBytes / 1024 / 1024,
            'file_mb' => $fileSize / 1024 / 1024,
        ];

        echo sprintf(
            "%s | %s rows/s | peak %s MB | file %s MB\n",
            str_pad(round($duration, 2).'s', 8),
            str_pad(number_format(round($rowsRead / max($duration, 0.0001))), 9, ' ', STR_PAD_LEFT),
            str_pad(round($peakBytes / 1024 / 1024, 1), 5, ' ', STR_PAD_LEFT),
            round($fileSize / 1024 / 1024, 2)
        );

        unset($reader);
        gc_collect_cycles();
    }
}

// ---------------------------------------------------------------------------
// Summary tables
// ---------------------------------------------------------------------------

echo "\n=================================================================================\n";
echo "                                   SUMMARY                                       \n";
echo "=================================================================================\n\n";

if ($results['local']) {
    echo "Local sequential reads\n";
    echo "----------------------\n";
    echo "| Rows | Read Speed | Read Time | Peak RAM | File Size |\n";
    echo "|------|------------|-----------|----------|-----------|\n";
    foreach ($results['local'] as $label => $r) {
        echo sprintf(
            "| %s | %s rows/s | %.2fs | %.1f MB | %.2f MB |\n",
            number_format($r['rows']),
            number_format(round($r['speed'])),
            $r['duration'],
            $r['peak_mb'],
            $r['file_mb']
        );
    }
}

if ($results['s3']) {
    echo "\nS3 streaming reads (cold cache)\n";
    echo "-------------------------------\n";
    echo "| Rows | Read Speed | Read Time | Peak RAM | File Size |\n";
    echo "|------|------------|-----------|----------|-----------|\n";
    foreach ($results['s3'] as $label => $r) {
        echo sprintf(
            "| %s | %s rows/s | %.2fs | %.1f MB | %.2f MB |\n",
            number_format($r['rows']),
            number_format(round($r['speed'])),
            $r['duration'],
            $r['peak_mb'],
            $r['file_mb']
        );
    }
}

echo "\n";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function writeWorkload(string $path, int $rows, array $headers, array $departments, array $statuses, string $today): void
{
    $writer = new SinkableXlsxWriter(new FileSink($path));
    $writer->setCompressionLevel(1)->setBufferFlushInterval($rows < 10000 ? 1000 : 10000);
    $writer->startFile($headers);
    for ($i = 1; $i <= $rows; $i++) {
        $writer->writeRow([
            $i,
            'Employee'.($i % 100),
            'emp'.($i % 100).'@company.com',
            $departments[$i % 10],
            50000 + ($i % 50) * 1000,
            $today,
            $statuses[$i % 5],
            'Standard notes for employee record',
        ]);
    }
    $writer->finishFile();
}

function writeWorkloadToSink($sink, int $rows, array $headers, array $departments, array $statuses, string $today): void
{
    $writer = new SinkableXlsxWriter($sink);
    $writer->setCompressionLevel(1)->setBufferFlushInterval($rows < 10000 ? 1000 : 10000);
    $writer->startFile($headers);
    for ($i = 1; $i <= $rows; $i++) {
        $writer->writeRow([
            $i,
            'Employee'.($i % 100),
            'emp'.($i % 100).'@company.com',
            $departments[$i % 10],
            50000 + ($i % 50) * 1000,
            $today,
            $statuses[$i % 5],
            'Standard notes for employee record',
        ]);
    }
    $writer->finishFile();
}

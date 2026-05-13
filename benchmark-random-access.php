<?php
/**
 * Random-access benchmark — writes a born-indexed and a plain XLSX of
 * the same size, then measures rowAt(N) latency at increasing target
 * positions to demonstrate the speedup the xl/_kxs/index.bin sidecar
 * delivers.
 *
 * Run:        php benchmark-random-access.php
 * Quick run:  php benchmark-random-access.php 100000
 *             (caps the total dataset size)
 *
 * Output:
 *   - per-target wall time, plain vs indexed
 *   - speedup factor at each target
 *   - rowCount() comparison (O(N) vs O(1))
 */

require_once __DIR__.'/vendor/autoload.php';

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

if (file_exists(__DIR__.'/.env')) {
    foreach (parse_ini_file(__DIR__.'/.env') as $key => $value) {
        $_ENV[$key] = $value;
    }
}

ini_set('memory_limit', '2G');

$totalRows = isset($argv[1]) ? (int) $argv[1] : 500000;
$syncEvery = 10000;

$headers = ['ID', 'Name', 'Email', 'Department', 'Salary', 'Date', 'Status', 'Notes'];
$departments = ['Sales', 'Marketing', 'Engineering', 'HR', 'Finance', 'Operations', 'Support', 'Legal', 'Product', 'Design'];
$statuses = ['Active', 'Inactive', 'Pending', 'Suspended', 'Terminated'];
$today = date('Y-m-d');

echo "\n";
echo "=================================================================================\n";
echo "          KOLAY XLSX STREAM — RANDOM ACCESS BENCHMARK (v3.0+)                   \n";
echo "=================================================================================\n";
echo "Total rows: ".number_format($totalRows)." (header + ".number_format($totalRows - 1)." data)\n";
echo "Sync period: every ".number_format($syncEvery)." rows\n\n";

$plainPath = tempnam(sys_get_temp_dir(), 'kxs_plain_');
$indexedPath = tempnam(sys_get_temp_dir(), 'kxs_indexed_');

echo "Writing plain XLSX...   ";
flush();
$start = microtime(true);
writeWorkload(new FileSink($plainPath), $totalRows - 1, $headers, $departments, $statuses, $today, indexed: false);
$plainWriteTime = microtime(true) - $start;
$plainSize = filesize($plainPath);
echo round($plainWriteTime, 2)."s, ".round($plainSize / 1024 / 1024, 2)." MB\n";

echo "Writing indexed XLSX... ";
flush();
$start = microtime(true);
writeWorkload(new FileSink($indexedPath), $totalRows - 1, $headers, $departments, $statuses, $today, indexed: true, syncEvery: $syncEvery);
$indexedWriteTime = microtime(true) - $start;
$indexedSize = filesize($indexedPath);
echo round($indexedWriteTime, 2)."s, ".round($indexedSize / 1024 / 1024, 2)." MB\n";

echo sprintf(
    "\nWrite cost of indexing: time %+.2f%%, size %+.3f%%\n\n",
    ($indexedWriteTime - $plainWriteTime) / $plainWriteTime * 100,
    ($indexedSize - $plainSize) / $plainSize * 100
);

// ---------------------------------------------------------------------------
// rowAt(N) latency comparison
// ---------------------------------------------------------------------------

$targets = [
    1,
    2,
    100,
    intdiv($totalRows, 10),
    intdiv($totalRows, 4),
    intdiv($totalRows, 2),
    intdiv($totalRows * 3, 4),
    intdiv($totalRows * 9, 10),
    $totalRows - 100,
    $totalRows,
];
$targets = array_values(array_unique(array_filter($targets, fn ($n) => $n >= 1 && $n <= $totalRows)));
sort($targets);

echo "rowAt(N) latency (local, ms)\n";
echo "-----------------------------\n";
echo "| Target Row | Plain (ms) | Indexed (ms) | Speedup |\n";
echo "|------------|------------|--------------|---------|\n";

$plainReader = StreamingXlsxReader::fromFile($plainPath);
$indexedReader = StreamingXlsxReader::fromFile($indexedPath);

foreach ($targets as $target) {
    // Plain reader is stateful through the source — reopen each call to
    // avoid coupling the timer with prior reads.
    $plainReader = StreamingXlsxReader::fromFile($plainPath);
    $start = microtime(true);
    $rowPlain = $plainReader->rowAt($target);
    $plainMs = (microtime(true) - $start) * 1000;

    $indexedReader = StreamingXlsxReader::fromFile($indexedPath);
    $start = microtime(true);
    $rowIndexed = $indexedReader->rowAt($target);
    $indexedMs = (microtime(true) - $start) * 1000;

    if ($rowPlain !== $rowIndexed) {
        echo "  ! mismatch at row {$target}\n";
        echo "  plain:   ".json_encode($rowPlain)."\n";
        echo "  indexed: ".json_encode($rowIndexed)."\n";
    }

    $speedup = $indexedMs > 0 ? $plainMs / $indexedMs : 0;
    echo sprintf(
        "| %s | %s | %s | %s |\n",
        str_pad(number_format($target), 10, ' ', STR_PAD_LEFT),
        str_pad(round($plainMs, 1), 10, ' ', STR_PAD_LEFT),
        str_pad(round($indexedMs, 1), 12, ' ', STR_PAD_LEFT),
        str_pad(round($speedup, 1).'×', 7, ' ', STR_PAD_LEFT)
    );
}

// ---------------------------------------------------------------------------
// rowCount comparison
// ---------------------------------------------------------------------------

echo "\nrowCount() — O(N) full scan vs O(1) sidecar lookup\n";
echo "---------------------------------------------------\n";

$plainReader = StreamingXlsxReader::fromFile($plainPath);
$start = microtime(true);
$plainCount = $plainReader->rowCount();
$plainCountMs = (microtime(true) - $start) * 1000;

$indexedReader = StreamingXlsxReader::fromFile($indexedPath);
$start = microtime(true);
$indexedCount = $indexedReader->rowCount();
$indexedCountMs = (microtime(true) - $start) * 1000;

echo sprintf(
    "Plain   rowCount(): %s ms (returned %s rows)\n",
    str_pad(round($plainCountMs, 1), 8, ' ', STR_PAD_LEFT),
    number_format($plainCount)
);
echo sprintf(
    "Indexed rowCount(): %s ms (returned %s rows)\n",
    str_pad(round($indexedCountMs, 1), 8, ' ', STR_PAD_LEFT),
    number_format($indexedCount)
);
echo sprintf(
    "Speedup: %.0f×\n",
    $indexedCountMs > 0 ? $plainCountMs / $indexedCountMs : 0
);

@unlink($plainPath);
@unlink($indexedPath);

echo "\n";

function writeWorkload(
    $sink,
    int $dataRows,
    array $headers,
    array $departments,
    array $statuses,
    string $today,
    bool $indexed = false,
    int $syncEvery = 10000,
): void {
    $writer = new SinkableXlsxWriter($sink);
    $writer->setCompressionLevel(1)->setBufferFlushInterval(min(10000, $syncEvery));
    if ($indexed) {
        $writer->withRandomAccessIndex(every: $syncEvery);
    }
    $writer->startFile($headers);
    for ($i = 1; $i <= $dataRows; $i++) {
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

<?php

/**
 * writeRow throughput benchmark.
 *
 * Writes ROWS rows of a realistic mixed-type record to a temp .xlsx and
 * reports wall time + peak memory. Used to measure the cost of the
 * per-row style feature (v3.0.2) against the v3.0.1 baseline.
 *
 * Modes (argv[1]):
 *   unstyled  — writeRow($row)                 (baseline + v3.0.2 fast path)
 *   styled    — writeRow($row, $styleId) where ~1% of rows get a fill+color
 *               (only meaningful once the feature exists; falls back to
 *               unstyled on v3.0.1)
 *
 * Output: a single line of JSON so the runner can aggregate across runs.
 */

require __DIR__ . '/../vendor/autoload.php';

use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

$mode = $argv[1] ?? 'unstyled';
$rows = (int) ($argv[2] ?? 500000);

$tmp = tempnam(sys_get_temp_dir(), 'kxs_bench_') . '.xlsx';

$writer = SinkableXlsxWriter::createForFile($tmp);
$writer->startFile(['ID', 'Name', 'Email', 'Amount', 'CreatedAt', 'Status']);

// Pre-register a row style only when the feature exists AND styled mode is on.
$styleId = null;
$canStyle = method_exists($writer, 'registerRowStyle');
if (($mode === 'styled' || $mode === 'allstyled') && $canStyle) {
    $styleId = $writer->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
}

$base = new DateTimeImmutable('2024-01-01 09:00:00');

// DateTime construction via ->modify('+N minutes') is the expensive part of
// row generation (string parse per call); pre-build a one-day pool so the
// timed loop reflects writer cost, not test-data cost. 1440 entries is far
// beyond deflate's 32KB window, so compression behaviour is unaffected.
$datePool = [];
for ($m = 0; $m < 1440; $m++) {
    $datePool[$m] = $base->modify('+' . $m . ' minutes');
}

// Calibration pass: measure what building the rows costs by itself, in the
// same process right before the timed run, and subtract it. Keeps the
// benchmark honest without pre-materializing 500K rows (which would poison
// the peak-memory number that the constant-RAM claim rests on).
$calStart = hrtime(true);
for ($i = 1; $i <= $rows; $i++) {
    $failed = ($i % 100 === 0);
    $row = [
        $i,
        'Customer ' . $i,
        'user' . $i . '@example.com',
        round((($i * 73) % 1000000) / 100, 2),
        $datePool[$i % 1440],
        $failed ? 'FAILED' : 'OK',
    ];
}
$genSeconds = (hrtime(true) - $calStart) / 1e9;

$start = hrtime(true);

for ($i = 1; $i <= $rows; $i++) {
    $failed = ($i % 100 === 0); // ~1% "failed" rows
    $row = [
        $i,
        'Customer ' . $i,
        'user' . $i . '@example.com',
        round((($i * 73) % 1000000) / 100, 2),
        $datePool[$i % 1440],
        $failed ? 'FAILED' : 'OK',
    ];

    if ($mode === 'allstyled' && $styleId !== null) {
        $writer->writeRow($row, $styleId);           // every row styled (worst case)
    } elseif ($mode === 'styled' && $styleId !== null) {
        $writer->writeRow($row, $failed ? $styleId : null); // ~1% styled (realistic)
    } else {
        $writer->writeRow($row);
    }
}

$writer->finishFile();

$elapsedTotal = (hrtime(true) - $start) / 1e9;
$elapsed = $elapsedTotal - $genSeconds;
$peakMb = memory_get_peak_usage(true) / 1048576;
$sizeMb = filesize($tmp) / 1048576;

@unlink($tmp);

echo json_encode([
    'mode' => $mode,
    'styled_supported' => $canStyle,
    'rows' => $rows,
    'seconds' => round($elapsed, 4),          // writer-only (generation subtracted)
    'total_seconds' => round($elapsedTotal, 4),
    'gen_seconds' => round($genSeconds, 4),
    'rows_per_sec' => (int) round($rows / $elapsed),
    'peak_mb' => round($peakMb, 1),
    'file_mb' => round($sizeMb, 2),
]) . "\n";

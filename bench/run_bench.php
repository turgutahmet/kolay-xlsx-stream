<?php

/**
 * Runs row_style_bench.php N times in a fresh PHP process each time and
 * aggregates the timing. Fresh process per run = no warm-up/GC carryover
 * skewing the numbers.
 *
 * Usage: php bench/run_bench.php <mode> <rows> <runs>
 */

$mode = $argv[1] ?? 'unstyled';
$rows = (int) ($argv[2] ?? 500000);
$runs = (int) ($argv[3] ?? 10);

$bench = __DIR__ . '/row_style_bench.php';
$php = PHP_BINARY;

$seconds = [];
$rps = [];
$peak = null;

fwrite(STDERR, "Running '$mode' x$runs ($rows rows)...\n");

for ($r = 1; $r <= $runs; $r++) {
    $out = shell_exec(sprintf('%s %s %s %d 2>/dev/null', escapeshellarg($php), escapeshellarg($bench), escapeshellarg($mode), $rows));
    $data = json_decode(trim((string) $out), true);
    if (!is_array($data) || !isset($data['seconds'])) {
        fwrite(STDERR, "  run $r FAILED: " . trim((string) $out) . "\n");
        continue;
    }
    $seconds[] = $data['seconds'];
    $rps[] = $data['rows_per_sec'];
    $peak = $data['peak_mb'];
    fwrite(STDERR, sprintf("  run %2d: %.4fs  %d rows/s\n", $r, $data['seconds'], $data['rows_per_sec']));
}

sort($seconds);
$n = count($seconds);
$median = $n % 2 ? $seconds[intdiv($n, 2)] : ($seconds[$n / 2 - 1] + $seconds[$n / 2]) / 2;
$mean = array_sum($seconds) / $n;

echo "\n";
echo json_encode([
    'mode' => $mode,
    'rows' => $rows,
    'runs' => $n,
    'peak_mb' => $peak,
    'min_s' => round(min($seconds), 4),
    'median_s' => round($median, 4),
    'mean_s' => round($mean, 4),
    'max_s' => round(max($seconds), 4),
    'best_rows_per_sec' => max($rps),
], JSON_PRETTY_PRINT) . "\n";

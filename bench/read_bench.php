<?php

/**
 * Full-scan read throughput benchmark (rows/s) over files this
 * package's writer produces natively (inlineStr cells, no shared
 * strings) — the read-path counterpart of row_style_bench.php.
 *
 * Same hygiene as run_bench.php: every measured run happens in a FRESH
 * php process (no warm-up/GC/opcache-state carryover), the parent only
 * orchestrates and aggregates min/median/mean/max. The fixture is
 * built once and cached in the system temp dir keyed by shape+rows, so
 * repeated invocations measure reading, not writing.
 *
 * Modes (argv[1]):
 *   narrow — 6 mixed-type columns (id, name, email, amount, date, status);
 *            the row_style_bench record shape
 *   wide   — 30 columns (10 of each: int, float, short string) per row
 *
 * Usage:  php bench/read_bench.php <narrow|wide> [rows=500000] [runs=5]
 * Child:  (internal) php bench/read_bench.php --child <fixture> — one
 *         timed full scan, one line of JSON on stdout.
 */

require __DIR__.'/../vendor/autoload.php';

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

// ---------------------------------------------------------------- child
if (($argv[1] ?? '') === '--child') {
    $fixture = $argv[2];

    $reader = StreamingXlsxReader::fromFile($fixture);
    $start = hrtime(true);
    $n = 0;
    foreach ($reader->rows() as $row) {
        $n++;
    }
    $seconds = (hrtime(true) - $start) / 1e9;
    $reader->close();

    echo json_encode([
        'rows_read' => $n,
        'seconds' => round($seconds, 4),
        'rows_per_sec' => (int) round($n / $seconds),
        'peak_mb' => round(memory_get_peak_usage(true) / 1048576, 1),
    ])."\n";
    exit(0);
}

// --------------------------------------------------------------- parent
$mode = $argv[1] ?? 'narrow';
$rows = (int) ($argv[2] ?? 500000);
$runs = (int) ($argv[3] ?? 5);

if (! in_array($mode, ['narrow', 'wide'], true)) {
    fwrite(STDERR, "mode must be narrow|wide\n");
    exit(1);
}

$fixture = sys_get_temp_dir()."/kxs_read_bench_{$mode}_{$rows}.xlsx";
if (! file_exists($fixture) || filesize($fixture) < 1000) {
    fwrite(STDERR, "building {$mode} fixture ({$rows} rows)...\n");
    $writer = SinkableXlsxWriter::createForFile($fixture);
    if ($mode === 'narrow') {
        $writer->startFile(['ID', 'Name', 'Email', 'Amount', 'CreatedAt', 'Status']);
        for ($i = 1; $i <= $rows; $i++) {
            $writer->writeRow([
                $i,
                'Customer '.$i,
                'user'.$i.'@example.com',
                round((($i * 73) % 1000000) / 100, 2),
                45292.375 + ($i % 1440) / 1440, // date serial, pre-computed
                ($i % 100 === 0) ? 'FAILED' : 'OK',
            ]);
        }
    } else {
        $header = [];
        for ($c = 1; $c <= 30; $c++) {
            $header[] = 'Col'.$c;
        }
        $writer->startFile($header);
        for ($i = 1; $i <= $rows; $i++) {
            $row = [];
            for ($c = 0; $c < 10; $c++) {
                $row[] = $i + $c;                       // 10 ints
            }
            for ($c = 0; $c < 10; $c++) {
                $row[] = ($i + $c) * 0.25;              // 10 floats
            }
            for ($c = 0; $c < 10; $c++) {
                $row[] = 'v'.(($i + $c) % 1000);        // 10 short strings
            }
            $writer->writeRow($row);
        }
    }
    $writer->finishFile();
}

$php = PHP_BINARY;
$self = __FILE__;
$seconds = [];
$rps = [];
$peak = null;
$rowsRead = null;

fwrite(STDERR, "Reading '{$mode}' x{$runs} ({$rows} rows)...\n");
for ($r = 1; $r <= $runs; $r++) {
    $out = shell_exec(sprintf(
        '%s %s --child %s 2>/dev/null',
        escapeshellarg($php), escapeshellarg($self), escapeshellarg($fixture)
    ));
    $data = json_decode(trim((string) $out), true);
    if (! is_array($data) || ! isset($data['seconds'])) {
        fwrite(STDERR, "  run {$r} FAILED: ".trim((string) $out)."\n");

        continue;
    }
    $seconds[] = $data['seconds'];
    $rps[] = $data['rows_per_sec'];
    $peak = $data['peak_mb'];
    $rowsRead = $data['rows_read'];
    fwrite(STDERR, sprintf("  run %2d: %.4fs  %d rows/s\n", $r, $data['seconds'], $data['rows_per_sec']));
}

sort($seconds);
$n = count($seconds);
$median = $n % 2 ? $seconds[intdiv($n, 2)] : ($seconds[$n / 2 - 1] + $seconds[$n / 2]) / 2;

echo "\n";
echo json_encode([
    'bench' => 'read_full_scan',
    'mode' => $mode,
    'rows' => $rowsRead,
    'file_mb' => round(filesize($fixture) / 1048576, 2),
    'runs' => $n,
    'peak_mb' => $peak,
    'min_s' => round(min($seconds), 4),
    'median_s' => round($median, 4),
    'mean_s' => round(array_sum($seconds) / $n, 4),
    'max_s' => round(max($seconds), 4),
    'best_rows_per_sec' => max($rps),
], JSON_PRETTY_PRINT)."\n";

<?php

/**
 * Query-surface benchmark for born-indexed files: cold open, rowAt,
 * findRow, rowsWhere at several selectivities, rows(skip=N),
 * groupStats, quantile and countDistinct — the operations the KXSI
 * sidecar exists to accelerate.
 *
 * Same hygiene as run_bench.php: each op is measured in a FRESH php
 * process per run (cold open included in-process but timed separately),
 * the parent aggregates the median across runs. The fixture — 1M rows,
 * sync every 10K, stats on the id, amount and group columns, sketches
 * (TDIG/CHLL) on the id, name and amount columns, group column sorted
 * with 20 duplicate groups — is built once and cached in the system
 * temp dir.
 *
 * Usage:  php bench/query_bench.php [op=all] [rows=1000000] [runs=5]
 *         op ∈ all | cold_open | rowAt | findRow | rowsWhere |
 *              rows_skip | groupStats | quantile | countDistinct
 * Child:  (internal) php bench/query_bench.php --child <op> <fixture> <rows>
 */

require __DIR__.'/../vendor/autoload.php';

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

const OPS = ['cold_open', 'rowAt', 'findRow', 'rowsWhere', 'rows_skip', 'groupStats', 'quantile', 'countDistinct'];

// ---------------------------------------------------------------- child
if (($argv[1] ?? '') === '--child') {
    [, , $op, $fixture, $rowsArg] = $argv;
    $rows = (int) $rowsArg;

    // Cold open is part of every child (fresh process = cold CD/index
    // state) and reported on its own so the per-call numbers below
    // reflect warm-reader cost, the way a query workload runs.
    $t0 = hrtime(true);
    $reader = StreamingXlsxReader::fromFile($fixture);
    $reader->rowCount(); // forces sidecar load + staleness validation
    $coldOpenMs = (hrtime(true) - $t0) / 1e6;

    $result = ['op' => $op, 'cold_open_ms' => round($coldOpenMs, 3)];
    mt_srand(42); // fixed targets — comparable across runs and versions

    switch ($op) {
        case 'cold_open':
            break; // already measured

        case 'rowAt': // 200 random point reads, uniform in-block offsets
            $calls = 200;
            $t0 = hrtime(true);
            for ($i = 0; $i < $calls; $i++) {
                $reader->rowAt(mt_rand(2, $rows));
            }
            $result['ms_per_call'] = round((hrtime(true) - $t0) / 1e6 / $calls, 3);
            break;

        case 'findRow': // 200 random id lookups (sorted column -> ~1 block each)
            $calls = 200;
            $t0 = hrtime(true);
            for ($i = 0; $i < $calls; $i++) {
                $reader->findRow(1, mt_rand(1, $rows));
            }
            $result['ms_per_call'] = round((hrtime(true) - $t0) / 1e6 / $calls, 3);
            break;

        case 'rowsWhere': // range scans at three selectivities on the sorted id
            foreach ([100 => '0.01%', 10000 => '1%', 100000 => '10%'] as $span => $label) {
                $lo = intdiv($rows, 2);
                $t0 = hrtime(true);
                $n = 0;
                foreach ($reader->rowsWhere(1, 'between', $lo, $lo + $span - 1) as $row) {
                    $n++;
                }
                $result['selectivity_'.$label] = ['rows' => $n, 'ms' => round((hrtime(true) - $t0) / 1e6, 1)];
            }
            break;

        case 'rows_skip': // deep sequential entry: skip almost everything
            $skip = $rows - 1000;
            $t0 = hrtime(true);
            $n = 0;
            foreach ($reader->rows($skip, 100) as $row) {
                $n++;
            }
            $result['skip'] = $skip;
            $result['rows_yielded'] = $n;
            $result['ms'] = round((hrtime(true) - $t0) / 1e6, 1);
            break;

        case 'groupStats': // sorted-group pushdown over the whole sheet
            $t0 = hrtime(true);
            $groups = $reader->groupStats(4, 3);
            $result['groups'] = count($groups);
            $result['ms'] = round((hrtime(true) - $t0) / 1e6, 1);
            break;

        case 'quantile': // sidecar-only TDIG estimates; first call pays the lazy deserialize
            $t0 = hrtime(true);
            $reader->median(3);
            $result['first_call_ms'] = round((hrtime(true) - $t0) / 1e6, 3);
            $calls = 1000;
            $t0 = hrtime(true);
            for ($i = 0; $i < $calls; $i++) {
                $reader->quantile(3, mt_rand(0, 1000) / 1000);
            }
            $result['ms_per_call'] = round((hrtime(true) - $t0) / 1e6 / $calls, 4);
            break;

        case 'countDistinct': // sidecar-only CHLL estimates; ditto on the first call
            $t0 = hrtime(true);
            $reader->countDistinct(2);
            $result['first_call_ms'] = round((hrtime(true) - $t0) / 1e6, 3);
            $calls = 1000;
            $t0 = hrtime(true);
            for ($i = 0; $i < $calls; $i++) {
                $reader->countDistinct($i % 2 === 0 ? 1 : 2);
            }
            $result['ms_per_call'] = round((hrtime(true) - $t0) / 1e6 / $calls, 4);
            break;

        default:
            fwrite(STDERR, "unknown op {$op}\n");
            exit(1);
    }

    $reader->close();
    $result['peak_mb'] = round(memory_get_peak_usage(true) / 1048576, 1);
    echo json_encode($result)."\n";
    exit(0);
}

// --------------------------------------------------------------- parent
$opArg = $argv[1] ?? 'all';
$rows = (int) ($argv[2] ?? 1000000);
$runs = (int) ($argv[3] ?? 5);
$ops = $opArg === 'all' ? OPS : [$opArg];

foreach ($ops as $op) {
    if (! in_array($op, OPS, true)) {
        fwrite(STDERR, 'op must be all|'.implode('|', OPS)."\n");
        exit(1);
    }
}

// v2 fixture name: v3.2 sketches (TDIG/CHLL) were added to the fixture —
// a cached pre-sketch file would silently answer quantile ops with null.
$fixture = sys_get_temp_dir()."/kxs_query_bench_v2_{$rows}.xlsx";
if (! file_exists($fixture) || filesize($fixture) < 1000) {
    fwrite(STDERR, "building query fixture ({$rows} rows)...\n");
    $writer = SinkableXlsxWriter::createForFile($fixture);
    $writer->withRandomAccessIndex(10000)->withColumnStats([1, 3, 4])->withColumnSketches([1, 2, 3]);
    $writer->startFile(['ID', 'Name', 'Amount', 'Group']);
    $groupSpan = max(1, intdiv($rows, 20)); // 20 sorted duplicate groups
    for ($i = 1; $i <= $rows; $i++) {
        $writer->writeRow([$i, 'name-'.($i % 1000), $i * 0.5, intdiv($i - 1, $groupSpan)]);
    }
    $writer->finishFile();
}

$php = PHP_BINARY;
$self = __FILE__;
$report = [
    'bench' => 'query',
    'rows' => $rows,
    'file_mb' => round(filesize($fixture) / 1048576, 2),
    'runs_each' => $runs,
];

foreach ($ops as $op) {
    fwrite(STDERR, "op '{$op}' x{$runs}...\n");
    $samples = [];
    foreach (range(1, $runs) as $r) {
        $out = shell_exec(sprintf(
            '%s %s --child %s %s %d 2>/dev/null',
            escapeshellarg($php),
            escapeshellarg($self),
            escapeshellarg($op),
            escapeshellarg($fixture),
            $rows
        ));
        $data = json_decode(trim((string) $out), true);
        if (! is_array($data)) {
            fwrite(STDERR, "  run {$r} FAILED: ".trim((string) $out)."\n");

            continue;
        }
        $samples[] = $data;
        fwrite(STDERR, '  run '.$r.': '.json_encode($data)."\n");
    }
    if ($samples === []) {
        continue;
    }

    // Median by the op's primary timing metric (fresh-process medians —
    // the same aggregation run_bench.php applies to writer runs).
    $metric = match ($op) {
        'cold_open' => 'cold_open_ms',
        'rowAt', 'findRow', 'quantile', 'countDistinct' => 'ms_per_call',
        default => 'ms',
    };
    usort($samples, function (array $a, array $b) use ($metric) {
        $av = $a[$metric] ?? $a['selectivity_1%']['ms'];
        $bv = $b[$metric] ?? $b['selectivity_1%']['ms'];

        return $av <=> $bv;
    });
    $median = $samples[intdiv(count($samples), 2)];
    unset($median['op']);
    $report[$op] = $median;
}

echo "\n".json_encode($report, JSON_PRETTY_PRINT)."\n";

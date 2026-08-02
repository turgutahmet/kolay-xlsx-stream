<?php

/**
 * Regenerates the KXSI conformance vectors (see SPEC.md §8).
 *
 * Run manually — CI and the phpunit suite NEVER execute this file:
 *
 *     php tests/SpecVectors/generate.php [vector-name]
 *
 * An optional vector name regenerates ONLY that vector — the tool for
 * adding a new vector without churning the committed bytes (and the
 * platform-stamped .xlsx timestamps) of the existing ones.
 *
 * Each vector is a small .xlsx produced by the real writer, plus:
 *   - <name>.expected.json  decoded-sidecar golden (rows, sync points,
 *     SCRC values, per-column block stats)
 *   - <name>.sidecar.hex    hexdump of the raw xl/_kxs/index.bin payload
 *
 * SpecVectorsTest reads the committed outputs and asserts they still
 *   agree with each other and with the current decoder — it regenerates
 * nothing, so the committed bytes pin the format against accidental
 * drift. Regenerate ONLY on a deliberate, spec-reviewed format change,
 * and commit the diff alongside the SPEC.md change that justifies it.
 */

use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

require __DIR__.'/../../vendor/autoload.php';

// Shortest-round-trip float serialization so goldens re-parse to the
// exact float64 the sidecar carries.
ini_set('serialize_precision', '-1');

/**
 * @param  callable(SinkableXlsxWriter): void  $write
 */
function generateVector(string $name, callable $write): void
{
    $only = $GLOBALS['argv'][1] ?? null;
    if ($only !== null && $only !== $name) {
        return;
    }

    $path = __DIR__."/{$name}.xlsx";
    @unlink($path);

    $writer = new SinkableXlsxWriter(new FileSink($path));
    $write($writer);

    $source = new LocalFileSource($path);
    $cd = ZipDirectory::fromSource($source);
    $payload = $cd->readEntry($source, ReaderIndex::ENTRY_PATH);
    $index = ReaderIndex::decode($payload);

    // Sheet entries in core-body order, straight from the payload.
    $entries = [];
    $sheetCount = unpack('v', substr($payload, 6, 2))[1];
    $body = substr($payload, 16);
    $cursor = 0;
    for ($i = 0; $i < $sheetCount; $i++) {
        $pathLen = unpack('v', substr($body, $cursor, 2))[1];
        $entries[] = substr($body, $cursor + 2, $pathLen);
        $cursor += 2 + $pathLen + 8;
        $syncCount = unpack('V', substr($body, $cursor, 4))[1];
        $cursor += 4 + 24 * $syncCount;
    }

    $sheets = [];
    foreach ($entries as $entry) {
        $columnStats = [];
        foreach ($index->statsColumns($entry) as $col) {
            $columnStats[(string) $col] = $index->columnStats($entry, $col);
        }
        // Derived sketch values (quantiles from TDIG, distinct estimate
        // from CHLL) — pins the ESTIMATORS against the committed bytes,
        // which the raw hexdump alone cannot do.
        $columnSketches = [];
        foreach ($index->digestColumns($entry) as $col) {
            $digest = $index->columnDigest($entry, $col);
            $quantiles = [];
            foreach (['0', '0.25', '0.5', '0.75', '1'] as $q) {
                $quantiles[$q] = $digest->quantile((float) $q);
            }
            $columnSketches[(string) $col] = [
                'quantiles' => $quantiles,
                'numeric_count' => $digest->count(),
                'distinct' => $index->columnHll($entry, $col)?->count(),
            ];
        }
        // String zone maps (STRZ). Added only when present so the pre-STRZ
        // vectors' goldens stay byte-for-byte unchanged.
        $stringZones = [];
        foreach ($index->stringStatsColumns($entry) as $col) {
            $stringZones[(string) $col] = $index->columnStringStats($entry, $col);
        }

        // Range-quantile superblocks (TDGB). Like the whole-column sketch
        // above, the golden pins each superblock's end_row plus the
        // quantiles its committed t-digest reproduces. Added only when
        // present so pre-TDGB goldens stay byte-for-byte unchanged.
        $rangeQuantiles = [];
        foreach ($index->rangeQuantileColumns($entry) as $col) {
            $superblocks = [];
            foreach ($index->rangeQuantileSuperblocks($entry, $col) as $sb) {
                $quantiles = [];
                foreach (['0', '0.5', '1'] as $q) {
                    $quantiles[$q] = $sb['digest']->quantile((float) $q);
                }
                $superblocks[] = [
                    'end_row' => $sb['end_row'],
                    'numeric_count' => $sb['digest']->count(),
                    'quantiles' => $quantiles,
                ];
            }
            $rangeQuantiles[(string) $col] = $superblocks;
        }

        // Frequent-items sketches (TOPK). The golden pins each column's
        // saturated bit and its (value, count) list in the sketch's own
        // deterministic order. Added only when present.
        $topValues = [];
        foreach ($index->topValueColumns($entry) as $col) {
            $sketch = $index->columnTopValues($entry, $col);
            $topValues[(string) $col] = [
                'saturated' => $sketch->saturated(),
                'values' => $sketch->topValues(),
            ];
        }

        // Argmin/argmax row pointers (ARGP), block-aligned 1:1 with STAT.
        // The golden pins each block's {minRow, maxRow} (0 = no numeric
        // value). Added only when present so pre-ARGP goldens stay unchanged.
        $argPointers = [];
        foreach ($index->argPointerColumns($entry) as $col) {
            $argPointers[(string) $col] = $index->argPointers($entry, $col);
        }

        $sheet = [
            'entry' => $entry,
            'total_rows' => $index->totalRows($entry),
            'sheet_crc32' => $index->sheetCrc32($entry),
            'sync_points' => $index->syncPoints($entry),
            'sync_point_crcs' => $index->syncPointCrcs($entry),
            'column_stats' => $columnStats === [] ? new stdClass() : $columnStats,
            'column_sketches' => $columnSketches === [] ? new stdClass() : $columnSketches,
        ];
        if ($stringZones !== []) {
            $sheet['string_zones'] = $stringZones;
        }
        if ($rangeQuantiles !== []) {
            $sheet['range_quantiles'] = $rangeQuantiles;
        }
        if ($topValues !== []) {
            $sheet['top_values'] = $topValues;
        }
        if ($argPointers !== []) {
            $sheet['arg_pointers'] = $argPointers;
        }
        $sheets[] = $sheet;
    }

    $golden = [
        'sync_period' => $index->syncPeriod(),
        'sheets' => $sheets,
    ];

    file_put_contents(
        __DIR__."/{$name}.expected.json",
        json_encode($golden, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES)."\n"
    );
    file_put_contents(
        __DIR__."/{$name}.sidecar.hex",
        rtrim(chunk_split(bin2hex($payload), 32, "\n"))."\n"
    );

    echo "  {$name}: ".strlen($payload)." sidecar bytes, ".count($sheets)." sheet(s)\n";
}

echo "Generating KXSI spec vectors...\n";

// Vector 1 — plain indexed, no stats. 250 data rows, sync every 100.
generateVector('vector-01-plain-indexed', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 100);
    $w->setBufferFlushInterval(100);
    $w->startFile(['id', 'name']);
    for ($i = 1; $i <= 250; $i++) {
        $w->writeRow([$i, "user-{$i}"]);
    }
    $w->finishFile();
});

// Vector 2 — indexed + STAT on one column carrying a mix of numeric and
// non-numeric values (exercises count vs other per block). Values are
// exact in float64 (integers + halves) so the golden round-trips.
generateVector('vector-02-stats', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 100);
    $w->withColumnStats([2]);
    $w->setBufferFlushInterval(100);
    $w->startFile(['id', 'amount', 'label']);
    for ($i = 1; $i <= 300; $i++) {
        $amount = $i % 10 === 0 ? 'n/a' : (($i * 7) % 100) + 0.5;
        $w->writeRow([$i, $amount, "row-{$i}"]);
    }
    $w->finishFile();
});

// Vector 3 — multi-sheet, indexed + STAT: sections repeat per sheet in
// core-body order, each sheet with its own sync/block cadence.
generateVector('vector-03-multisheet', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 50);
    $w->withColumnStats([1]);
    $w->setBufferFlushInterval(50);
    $w->startFile(['id', 'name']);
    for ($i = 1; $i <= 100; $i++) {
        $w->writeRow([$i, "alpha-{$i}"]);
    }
    $w->newSheet('Second', ['id', 'name']);
    for ($i = 1; $i <= 200; $i++) {
        $w->writeRow([$i * 3, "beta-{$i}"]);
    }
    $w->finishFile();
});

// Vector 4 — sortedness flags: col 1 ascending, col 2 descending,
// col 3 numeric but unsorted.
generateVector('vector-04-sorted', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 100);
    $w->withColumnStats([1, 2, 3]);
    $w->setBufferFlushInterval(100);
    $w->startFile(['asc', 'desc', 'shuffled']);
    for ($i = 1; $i <= 300; $i++) {
        $w->writeRow([$i, 1000 - $i, ($i * 37) % 100]);
    }
    $w->finishFile();
});

// Vector 5 — TDIG + CHLL sketches, no STAT (the sections are orthogonal
// opt-ins). Col 2 mixes numeric values (exact in float64: x + 0.25) with
// non-numeric 'n/a' markers — invisible to the t-digest, counted by the
// HLL; col 3 is pure text, covered by the HLL only. Header row excluded
// from both by the format.
generateVector('vector-05-sketches', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 100);
    $w->withColumnSketches([2, 3]);
    $w->setBufferFlushInterval(100);
    $w->startFile(['id', 'score', 'city']);
    $cities = ['Istanbul', 'Ankara', 'Izmir', 'Bursa', 'Antalya',
        'Adana', 'Konya', 'Gaziantep', 'Mersin', 'Kayseri'];
    for ($i = 1; $i <= 300; $i++) {
        $score = $i % 25 === 0 ? 'n/a' : (($i * 7) % 100) + 0.25;
        $w->writeRow([$i, $score, $cities[$i % 10].'-'.($i % 30)]);
    }
    $w->finishFile();
});

// Vector 6 — STRZ string zone maps. col 2 is a sorted invoice number with
// a long shared prefix (exercises deferred separator truncation); col 3 is
// a shuffled tag (unsorted string, weaker pruning); an empty cell every
// 40th row exercises the null/other class. sync every 50 -> multiple blocks.
generateVector('vector-06-string-zones', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 50);
    $w->withStringStats([2, 3]);
    $w->setBufferFlushInterval(50);
    $w->startFile(['id', 'invoice', 'tag']);
    $tags = ['alpha', 'bravo', 'charlie', 'delta', 'echo'];
    for ($i = 1; $i <= 200; $i++) {
        $invoice = sprintf('INV-2024-%08d', $i);
        $tag = $i % 40 === 0 ? '' : $tags[$i % 5].'-'.($i % 7);
        $w->writeRow([$i, $invoice, $tag]);
    }
    $w->finishFile();
});

// Vector 7 — TDGB range-quantile superblocks. The sheet is far under the
// 16384-row superblock width, so col 2 yields a SINGLE superblock whose
// end_row is the last data sheet row; the golden pins that boundary and
// the quantiles its committed t-digest reproduces. Non-numeric cells
// ('n/a' every 10th row) are excluded from the digest, so numeric_count
// pins the STAT-interpretation population rule. Multi-superblock spans
// (>16384 rows) are exercised by RangeQuantileTest rather than committed
// as a heavyweight fixture; the per-superblock frame repeats exactly like
// STAT's per-block frame, which the other vectors already pin.
generateVector('vector-07-range-quantiles', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 100);
    $w->withRangeQuantiles([2]);
    $w->setBufferFlushInterval(100);
    $w->startFile(['id', 'amount']);
    for ($i = 1; $i <= 300; $i++) {
        $amount = $i % 10 === 0 ? 'n/a' : (($i * 7) % 100) + 0.5;
        $w->writeRow([$i, $amount]);
    }
    $w->finishFile();
});

// Vector 8 — TOPK frequent-items sketches, both branches of the exactness
// switch in one file at k=8: col 2 'status' has 4 distinct values (≤ k →
// saturated=false, the golden pins the exact complete distribution), col 3
// 'region' has 20 distinct (> k → saturated=true, top-k with the N/k
// bound). The golden pins each column's saturated bit and (value, count)
// list; the hexdump pins the payload byte layout.
generateVector('vector-08-top-values', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 100);
    $w->withTopValues([2, 3], 8);
    $w->setBufferFlushInterval(100);
    $w->startFile(['id', 'status', 'region']);
    $statuses = ['paid', 'pending', 'refunded', 'failed'];
    for ($i = 1; $i <= 300; $i++) {
        $status = $i % 100 < 70 ? 'paid' : $statuses[$i % 4];
        // Two hot regions + an 18-way cold tail → 20 distinct > k, so the
        // sketch saturates while the heavy hitters survive with real counts.
        $region = $i % 100 < 55 ? 'r'.($i % 2) : 'r'.(2 + $i % 18);
        $w->writeRow([$i, $status, $region]);
    }
    $w->finishFile();
});

// Vector 9 — ARGP argmin/argmax row pointers, block-aligned 1:1 with STAT.
// A non-monotone amount column with a unique global max at data row 137
// (sheet row 138) and a unique global min at data row 200 (sheet row 201),
// so the golden pins the per-block {minRow, maxRow} and the two blocks that
// carry the global extremes. The hexdump pins the two-uint32-per-block layout.
generateVector('vector-09-arg-pointers', function (SinkableXlsxWriter $w): void {
    $w->withRandomAccessIndex(every: 50);
    $w->withArgPointers([2]);
    $w->setBufferFlushInterval(50);
    $w->startFile(['id', 'amount']);
    for ($i = 1; $i <= 250; $i++) {
        $amount = 100.0 + (($i * 31) % 50);
        if ($i === 137) {
            $amount = 9999.0;
        }
        if ($i === 200) {
            $amount = -5.0;
        }
        $w->writeRow([$i, $amount]);
    }
    $w->finishFile();
});

echo "Done.\n";

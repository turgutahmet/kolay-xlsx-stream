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
        $sheets[] = [
            'entry' => $entry,
            'total_rows' => $index->totalRows($entry),
            'sheet_crc32' => $index->sheetCrc32($entry),
            'sync_points' => $index->syncPoints($entry),
            'sync_point_crcs' => $index->syncPointCrcs($entry),
            'column_stats' => $columnStats === [] ? new stdClass() : $columnStats,
            'column_sketches' => $columnSketches === [] ? new stdClass() : $columnSketches,
        ];
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

echo "Done.\n";

<?php

/**
 * External XLSX reader/writer compatibility roundtrip — runs in the
 * tag-only `external-compat.yml` workflow with PhpSpreadsheet +
 * OpenSpout installed as dev deps. NOT a phpunit test: the optional
 * libs are not present in the regular CI matrix, so loading them
 * inside a phpunit class would fail the suite.
 *
 * Two assertions:
 *   1. PhpSpreadsheet can open a file produced by SinkableXlsxWriter
 *      (with and without random-access index).
 *   2. StreamingXlsxReader can read a file produced by PhpSpreadsheet
 *      (which uses xl/sharedStrings.xml — the external-XLSX path).
 *
 * Exits non-zero on any failure; intended to be run from the workflow.
 */

declare(strict_types=1);

require __DIR__.'/../../vendor/autoload.php';

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use OpenSpout\Reader\XLSX\Reader as OpenSpoutReader;
use PhpOffice\PhpSpreadsheet\IOFactory;

$failures = [];
$tmp = sys_get_temp_dir().'/kxs-compat';
if (! is_dir($tmp)) {
    mkdir($tmp, 0755, true);
}

function fail(array &$failures, string $name, string $reason): void
{
    $failures[] = "[FAIL] {$name} — {$reason}";
    fwrite(STDERR, end($failures)."\n");
}

function pass(string $name): void
{
    fwrite(STDOUT, "[OK]   {$name}\n");
}

// ---------------------------------------------------------------
// Test 1: kolay/xlsx-stream → PhpSpreadsheet
// ---------------------------------------------------------------
$kxsOut = $tmp.'/kxs-output.xlsx';
$writer = new SinkableXlsxWriter(new FileSink($kxsOut));
$writer->startFile(['id', 'name', 'amount']);
for ($i = 1; $i <= 250; $i++) {
    $writer->writeRow([$i, "user-{$i}", $i * 9.95]);
}
$writer->finishFile();

try {
    $book = IOFactory::load($kxsOut);
    $sheet = $book->getActiveSheet();
    $rowCount = $sheet->getHighestDataRow();
    if ($rowCount !== 251) {
        fail($failures, 'kxs→PhpSpreadsheet rowcount', "expected 251, got {$rowCount}");
    } else {
        $first = $sheet->getCell('B2')->getValue();
        $last = $sheet->getCell('B251')->getValue();
        if ($first !== 'user-1' || $last !== 'user-250') {
            fail($failures, 'kxs→PhpSpreadsheet content', "got first={$first} last={$last}");
        } else {
            pass('kxs→PhpSpreadsheet (250 rows)');
        }
    }
} catch (Throwable $e) {
    fail($failures, 'kxs→PhpSpreadsheet', $e->getMessage());
}

// ---------------------------------------------------------------
// Test 2: kolay/xlsx-stream with random-access index → PhpSpreadsheet
// (must ignore xl/_kxs/index.bin sidecar gracefully)
// ---------------------------------------------------------------
$kxsIndexed = $tmp.'/kxs-indexed.xlsx';
$writer = new SinkableXlsxWriter(new FileSink($kxsIndexed));
$writer->withRandomAccessIndex(every: 50);
$writer->startFile(['id', 'name']);
for ($i = 1; $i <= 200; $i++) {
    $writer->writeRow([$i, "user-{$i}"]);
}
$writer->finishFile();

try {
    $book = IOFactory::load($kxsIndexed);
    $sheet = $book->getActiveSheet();
    $rowCount = $sheet->getHighestDataRow();
    if ($rowCount !== 201) {
        fail($failures, 'kxs-indexed→PhpSpreadsheet', "expected 201, got {$rowCount}");
    } else {
        pass('kxs-indexed→PhpSpreadsheet (200 rows + sidecar ignored)');
    }
} catch (Throwable $e) {
    fail($failures, 'kxs-indexed→PhpSpreadsheet', $e->getMessage());
}

// ---------------------------------------------------------------
// Test 3: PhpSpreadsheet → kolay/xlsx-stream reader
// (PhpSpreadsheet emits xl/sharedStrings.xml so this exercises the
//  external-XLSX path)
// ---------------------------------------------------------------
$psOut = $tmp.'/ps-output.xlsx';
$psBook = new PhpOffice\PhpSpreadsheet\Spreadsheet();
$psSheet = $psBook->getActiveSheet();
$psSheet->fromArray([['id', 'name', 'amount']], null, 'A1');
$rows = [];
for ($i = 1; $i <= 100; $i++) {
    $rows[] = [$i, "user-{$i}", $i * 9.95];
}
$psSheet->fromArray($rows, null, 'A2');
$writerPs = IOFactory::createWriter($psBook, 'Xlsx');
$writerPs->save($psOut);

try {
    $reader = StreamingXlsxReader::fromFile($psOut);
    $allRows = iterator_to_array($reader->rows(), false);
    if (count($allRows) !== 101) {
        fail($failures, 'PhpSpreadsheet→kxs rowcount', 'expected 101, got '.count($allRows));
    } elseif ($allRows[0] !== ['id', 'name', 'amount']) {
        fail($failures, 'PhpSpreadsheet→kxs header', 'got '.json_encode($allRows[0]));
    } elseif ($allRows[1][1] !== 'user-1' || $allRows[100][1] !== 'user-100') {
        fail($failures, 'PhpSpreadsheet→kxs content', "first.name={$allRows[1][1]} last.name={$allRows[100][1]}");
    } else {
        pass('PhpSpreadsheet→kxs (100 rows via shared-strings)');
    }
} catch (Throwable $e) {
    fail($failures, 'PhpSpreadsheet→kxs', $e->getMessage());
}

// ---------------------------------------------------------------
// Test 4: kolay/xlsx-stream → OpenSpout
// ---------------------------------------------------------------
try {
    $reader = new OpenSpoutReader();
    $reader->open($kxsOut);
    $count = 0;
    foreach ($reader->getSheetIterator() as $sheet) {
        foreach ($sheet->getRowIterator() as $row) {
            $count++;
        }
    }
    $reader->close();
    if ($count !== 251) {
        fail($failures, 'kxs→OpenSpout', "expected 251 rows, got {$count}");
    } else {
        pass('kxs→OpenSpout (250 data + header)');
    }
} catch (Throwable $e) {
    fail($failures, 'kxs→OpenSpout', $e->getMessage());
}

// ---------------------------------------------------------------
// Cleanup + summary
// ---------------------------------------------------------------
foreach (glob($tmp.'/*.xlsx') ?: [] as $f) {
    @unlink($f);
}
@rmdir($tmp);

if (! empty($failures)) {
    fwrite(STDERR, "\n".count($failures)." failure(s):\n");
    foreach ($failures as $f) {
        fwrite(STDERR, $f."\n");
    }
    exit(1);
}

fwrite(STDOUT, "\nAll external compatibility checks passed.\n");
exit(0);

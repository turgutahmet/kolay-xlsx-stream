<?php

/**
 * Shared-strings resolution benchmark (peak RAM + parse time) over
 * external-shaped files whose cells reference xl/sharedStrings.xml —
 * the sst counterpart of read_bench.php.
 *
 * Same hygiene as run_bench.php: every measured run happens in a FRESH
 * php process (no warm-up/GC/opcache-state carryover), the parent only
 * orchestrates and aggregates min/median/mean/max. The fixture is
 * built once and cached in the system temp dir keyed by entries+rows,
 * so repeated invocations measure reading, not fixture synthesis.
 * SinkableXlsxWriter never emits t="s", so the fixture is a hand-built
 * OPC archive mirroring the PhpSpreadsheet/openpyxl layout.
 *
 * What the numbers mean: peak_mb is dominated by the packed table
 * (string payload + 4 bytes/entry offset index) — the sst XML itself
 * streams through the parser and never exists in memory. For scale,
 * the v3.2 packed+streaming model measured 24 MB peak on a 30 MB / 1M
 * entry table vs 84 MB for the previous inflate-then-array model.
 *
 * Usage:  php bench/sst_bench.php [entries=1000000] [rows=100000] [runs=5]
 * Child:  (internal) php bench/sst_bench.php --child <fixture> — one
 *         timed full scan, one line of JSON on stdout.
 */

require __DIR__.'/../vendor/autoload.php';

use Kolay\XlsxStream\Readers\StreamingXlsxReader;

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
$entries = (int) ($argv[1] ?? 1_000_000);
$rows = (int) ($argv[2] ?? 100_000);
$runs = (int) ($argv[3] ?? 5);

$fixture = sys_get_temp_dir()."/kxs_sst_bench_{$entries}_{$rows}.xlsx";
if (! file_exists($fixture) || filesize($fixture) < 1000) {
    fwrite(STDERR, "building sst fixture ({$entries} entries, {$rows} rows)...\n");

    // sst: realistic mixed-length strings, a sprinkle of entities.
    $parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'.$entries.'" uniqueCount="'.$entries.'">'];
    for ($i = 0; $i < $entries; $i++) {
        $s = 'customer-'.$i.($i % 50 === 0 ? ' &amp; sons' : '');
        $parts[] = '<si><t>'.$s.'</t></si>';
    }
    $parts[] = '</sst>';
    $sstXml = implode('', $parts);
    unset($parts);

    // sheet: every row references two sst entries + one numeric cell —
    // the dedup-heavy shape external writers produce.
    $sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'];
    $sheet[] = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>';
    for ($r = 2; $r <= $rows + 1; $r++) {
        $a = ($r * 37) % $entries;
        $b = ($r * 101) % $entries;
        $sheet[] = '<row r="'.$r.'"><c r="A'.$r.'" t="s"><v>'.$a.'</v></c><c r="B'.$r.'" t="s"><v>'.$b.'</v></c><c r="C'.$r.'" t="n"><v>'.$r.'</v></c></row>';
    }
    $sheet[] = '</sheetData></worksheet>';
    $sheetXml = implode('', $sheet);
    unset($sheet);

    $zip = new ZipArchive();
    $zip->open($fixture, ZipArchive::CREATE | ZipArchive::OVERWRITE);
    $zip->addFromString('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>');
    $zip->addFromString('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    $zip->addFromString('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>');
    $zip->addFromString('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>');
    $zip->addFromString('xl/sharedStrings.xml', $sstXml);
    $zip->addFromString('xl/worksheets/sheet1.xml', $sheetXml);
    $zip->close();
    fwrite(STDERR, sprintf("  sst %.1f MB uncompressed, file %.1f MB\n", strlen($sstXml) / 1048576, filesize($fixture) / 1048576));
    unset($sstXml, $sheetXml);
}

$php = PHP_BINARY;
$self = __FILE__;
$seconds = [];
$rps = [];
$peak = null;
$rowsRead = null;

fwrite(STDERR, "Reading sst-heavy file x{$runs} ({$entries} entries, {$rows} rows)...\n");
for ($r = 1; $r <= $runs; $r++) {
    $out = shell_exec(sprintf(
        '%s %s --child %s 2>/dev/null',
        escapeshellarg($php),
        escapeshellarg($self),
        escapeshellarg($fixture)
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
    fwrite(STDERR, sprintf("  run %2d: %.4fs  %d rows/s  peak %.1f MB\n", $r, $data['seconds'], $data['rows_per_sec'], $data['peak_mb']));
}

sort($seconds);
$n = count($seconds);
$median = $n % 2 ? $seconds[intdiv($n, 2)] : ($seconds[$n / 2 - 1] + $seconds[$n / 2]) / 2;

echo "\n";
echo json_encode([
    'bench' => 'sst_resolution_scan',
    'sst_entries' => $entries,
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

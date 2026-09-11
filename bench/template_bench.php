<?php

/**
 * Template-mode write benchmark.
 *
 * Answers the two questions template mode has to answer about cost:
 * what does streaming into someone else's layout cost against writing a
 * plain sheet, and does the shared string table hold memory flat.
 *
 * The template is a leave-report layout — merged title, styled header,
 * per-column number formats, two sample rows for the zebra, widths,
 * heights, frozen pane — produced once by this package's own classic
 * writer with the same styles the streamed rows will wear, so the two
 * modes are writing the same visible sheet and only the mechanism
 * differs. Both modes write identical data.
 *
 * Modes (argv[1]):
 *   classic   — startFile + writeRow with a registered row style
 *   template  — sheet() + writeRow(variant:), rows into the template's cut
 *   template_sst — the same, into a template that carries a shared string
 *                  table, so every text cell is interned
 *
 * Usage:  php bench/template_bench.php <mode> <rows> <cols> [runs]
 *         php bench/template_bench.php compare 100000 20 5
 *
 * A single run prints one JSON line so the compare mode can aggregate the
 * median across fresh processes — no warm-up or GC carryover between runs.
 */

require __DIR__.'/../vendor/autoload.php';

use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Templates\Template;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

$mode = $argv[1] ?? 'compare';
$rows = (int) ($argv[2] ?? 100000);
$cols = (int) ($argv[3] ?? 20);
$runs = (int) ($argv[4] ?? 5);

/**
 * One row of realistic mixed types: a name, a date, a datetime, then
 * alternating floats, ints and short strings out to $cols columns.
 *
 * @return list<mixed>
 */
function benchRow(int $i, int $cols, DateTimeImmutable $base): array
{
    $row = ['Employee '.$i, $base, $base->modify('+'.($i % 720).' hours')];
    for ($c = 3; $c < $cols; $c++) {
        $row[] = match ($c % 3) {
            0 => $i * 1.25 + $c,
            1 => ($i + $c) % 97,
            default => 'note-'.(($i + $c) % 50),
        };
    }

    return array_slice($row, 0, $cols);
}

/** @return list<string> */
function benchHeaders(int $cols): array
{
    $headers = ['Personnel', 'Start', 'Checked in'];
    for ($c = 3; $c < $cols; $c++) {
        $headers[] = 'Column '.($c + 1);
    }

    return array_slice($headers, 0, $cols);
}

/**
 * Build the layout once: header row plus two body rows carrying the two
 * alternating styles, so the cut has a sample per variant.
 */
function buildTemplate(int $cols): string
{
    $path = sys_get_temp_dir().'/kxs_tpl_bench_'.$cols.'.xlsx';
    if (is_file($path)) {
        return $path;
    }

    $writer = SinkableXlsxWriter::createForFile($path);
    $writer->setHeaderStyle(['bold' => true, 'fill' => '#1F4E78', 'color' => '#FFFFFF']);
    $writer->setColumnWidths([1 => 28, 2 => 14, 3 => 20]);
    $writer->freezeFirstRow();
    $writer->setColumnFormat(2, 'date');
    $writer->setColumnFormat(3, 'datetime');
    $plain = $writer->registerRowStyle(['fill' => '#FFFFFF']);
    $zebra = $writer->registerRowStyle(['fill' => '#EEF3F8']);
    $writer->startFile(benchHeaders($cols));
    $base = new DateTimeImmutable('2026-01-01 08:00:00');
    $writer->writeRow(benchRow(1, $cols, $base), $plain);
    $writer->writeRow(benchRow(2, $cols, $base), $zebra);
    $writer->finishFile();

    return $path;
}

/**
 * The same layout, but carrying an xl/sharedStrings.xml part.
 *
 * This package's own writer emits inline strings and no such part, so a
 * template it produced would never exercise the shared-string path. A
 * template from PhpSpreadsheet or Excel always carries one, and interning
 * every text cell is the one place template mode does more work per row
 * than the classic writer — which is exactly the cost this mode measures.
 */
function buildTemplateWithSharedStrings(int $cols): string
{
    $path = sys_get_temp_dir().'/kxs_tpl_bench_sst_'.$cols.'.xlsx';
    if (is_file($path)) {
        return $path;
    }

    copy(buildTemplate($cols), $path);

    $zip = new ZipArchive();
    $zip->open($path);

    $entries = '';
    foreach (benchHeaders($cols) as $header) {
        $entries .= '<si><t>'.htmlspecialchars($header, ENT_XML1).'</t></si>';
    }
    $zip->addFromString(
        'xl/sharedStrings.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        .'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        .'count="'.$cols.'" uniqueCount="'.$cols.'">'.$entries.'</sst>'
    );

    $rels = (string) $zip->getFromName('xl/_rels/workbook.xml.rels');
    $zip->addFromString('xl/_rels/workbook.xml.rels', str_replace(
        '</Relationships>',
        '<Relationship Id="rIdSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        .'Target="sharedStrings.xml"/></Relationships>',
        $rels
    ));

    $types = (string) $zip->getFromName('[Content_Types].xml');
    $zip->addFromString('[Content_Types].xml', str_replace(
        '</Types>',
        '<Override PartName="/xl/sharedStrings.xml" '
        .'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
        $types
    ));

    $zip->close();

    return $path;
}

// ---------------------------------------------------------------- child
if ($mode === 'classic' || $mode === 'template' || $mode === 'template_sst') {
    $out = tempnam(sys_get_temp_dir(), 'kxs_tplbench_').'.xlsx';
    $base = new DateTimeImmutable('2026-01-01 08:00:00');
    $headers = benchHeaders($cols);

    if ($mode === 'classic') {
        $writer = SinkableXlsxWriter::createForFile($out);
        $writer->setHeaderStyle(['bold' => true, 'fill' => '#1F4E78', 'color' => '#FFFFFF']);
        $writer->setColumnWidths([1 => 28, 2 => 14, 3 => 20]);
        $writer->freezeFirstRow();
        $writer->setColumnFormat(2, 'date');
        $writer->setColumnFormat(3, 'datetime');
        $plain = $writer->registerRowStyle(['fill' => '#FFFFFF']);
        $zebra = $writer->registerRowStyle(['fill' => '#EEF3F8']);
        $writer->startFile($headers);

        $start = hrtime(true);
        for ($i = 1; $i <= $rows; $i++) {
            $writer->writeRow(benchRow($i, $cols, $base), $i % 2 === 0 ? $zebra : $plain);
        }
        $writer->finishFile();
        $seconds = (hrtime(true) - $start) / 1e9;
    } else {
        $template = Template::open(
            $mode === 'template_sst' ? buildTemplateWithSharedStrings($cols) : buildTemplate($cols)
        );
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $template);
        $writer->sheet('Report', 2);

        $start = hrtime(true);
        for ($i = 1; $i <= $rows; $i++) {
            $writer->writeRow(benchRow($i, $cols, $base), variant: $i % 2);
        }
        $writer->finishFile();
        $seconds = (hrtime(true) - $start) / 1e9;
        $template->close();
    }

    $bytes = filesize($out);
    @unlink($out);

    echo json_encode([
        'mode' => $mode,
        'rows' => $rows,
        'cols' => $cols,
        'seconds' => round($seconds, 4),
        'rows_per_second' => (int) round($rows / $seconds),
        'peak_mb' => round(memory_get_peak_usage(true) / 1048576, 2),
        'bytes' => $bytes,
    ]), "\n";
    exit(0);
}

// --------------------------------------------------------------- parent
if ($mode !== 'compare') {
    fwrite(STDERR, "Unknown mode '{$mode}'. Use classic, template, template_sst or compare.\n");
    exit(1);
}

$results = [];
foreach (['classic', 'template', 'template_sst'] as $child) {
    $samples = [];
    fwrite(STDERR, "Running '{$child}' x{$runs} ({$rows} rows x {$cols} cols)...\n");
    for ($r = 0; $r < $runs; $r++) {
        $output = shell_exec(sprintf(
            '%s %s %s %d %d 2>/dev/null',
            escapeshellarg(PHP_BINARY),
            escapeshellarg(__FILE__),
            escapeshellarg($child),
            $rows,
            $cols
        ));
        $decoded = json_decode(trim((string) $output), true);
        if (! is_array($decoded) || ! isset($decoded['seconds'])) {
            fwrite(STDERR, "  run {$r} FAILED: ".trim((string) $output)."\n");

            continue;
        }
        $samples[] = $decoded;
    }
    if ($samples === []) {
        fwrite(STDERR, "No successful runs for '{$child}'.\n");
        exit(1);
    }

    usort($samples, fn ($a, $b) => $a['seconds'] <=> $b['seconds']);
    $results[$child] = $samples[intdiv(count($samples), 2)];
}

$classic = $results['classic'];

printf(
    "\n%-10s %10s %14s %10s %12s\n",
    'mode',
    'seconds',
    'rows/s',
    'peak MB',
    'bytes'
);
foreach ($results as $name => $r) {
    printf(
        "%-10s %10.3f %14s %10.2f %12s\n",
        $name,
        $r['seconds'],
        number_format($r['rows_per_second']),
        $r['peak_mb'],
        number_format($r['bytes'])
    );
}
echo "\n";
$summary = ['classic' => $classic];
foreach (['template', 'template_sst'] as $name) {
    $overhead = ($results[$name]['seconds'] / $classic['seconds'] - 1) * 100;
    $summary[$name] = $results[$name] + ['overhead_percent' => round($overhead, 1)];
    printf(
        "%-13s vs classic: %+.1f%% wall time, %+.2f MB peak\n",
        $name,
        $overhead,
        $results[$name]['peak_mb'] - $classic['peak_mb']
    );
}
echo "\n", json_encode($summary), "\n";

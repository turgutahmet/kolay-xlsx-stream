<?php

/**
 * Template-mode parity — runs in the tag-only `external-compat.yml`
 * workflow with PhpSpreadsheet installed as a dev dep. NOT a phpunit
 * test: the optional lib is absent from the regular CI matrix.
 *
 * The claim under test is the one template mode is sold on: take a layout
 * another producer authored, stream rows into it, and the result is the
 * file that producer would have written itself. So the comparison is not
 * against a hand-written expectation but against PhpSpreadsheet's own
 * output for the same data, read back through PhpSpreadsheet, cell by
 * cell, on every property a reader can observe:
 *
 *   value, data type, font, fill, border, alignment, number format,
 *   merges, column widths, row heights, frozen pane, gridlines.
 *
 * The layout deliberately mirrors a real leave-report sheet rather than a
 * minimal fixture: a merged title, a styled header row, per-column number
 * formats including a date and a datetime, alternating row styles, column
 * widths, custom row heights, a frozen pane and gridlines turned off. A
 * synthetic two-column sheet round-trips trivially and would prove nothing
 * about the traps that actually cost the proof-of-concept its first two
 * passes.
 *
 * Exits non-zero on any difference; intended to be run from the workflow.
 */

declare(strict_types=1);

require __DIR__.'/../../vendor/autoload.php';

use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Templates\Template;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date as ExcelDate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;

const DATA_ROWS = 400;
const DATA_START_ROW = 3;
const COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F'];

$failures = [];
$tmp = sys_get_temp_dir().'/kxs-template-compat';
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
// The data both paths must end up carrying
// ---------------------------------------------------------------

/** @return list<array{0: string, 1: float, 2: \DateTimeImmutable, 3: \DateTimeImmutable, 4: float, 5: int}> */
function dataRows(): array
{
    $rows = [];
    for ($i = 1; $i <= DATA_ROWS; $i++) {
        $rows[] = [
            'Employee '.$i,
            $i * 137.25,
            new DateTimeImmutable('2026-01-01 00:00:00 UTC'),
            (new DateTimeImmutable('2026-01-01 08:30:00 UTC'))->modify('+'.$i.' hours'),
            $i * 0.5,
            $i % 31,
        ];
    }

    return $rows;
}

// ---------------------------------------------------------------
// The layout, authored once and shared by both paths
// ---------------------------------------------------------------

/**
 * Build a leave-report layout: merged title, styled header, per-column
 * formats, alternating body styles, widths, heights, freeze, no gridlines.
 * $bodyRows says how many body rows to lay out — two for the template
 * (one per style variant), all of them for the reference workbook.
 */
function buildLayout(int $bodyRows, bool $withData): Spreadsheet
{
    $book = new Spreadsheet();
    $sheet = $book->getActiveSheet();
    $sheet->setTitle('Leaves');

    // Title row: merged across the table, tall, centred, filled.
    $sheet->setCellValue('A1', 'Leave report 2026');
    $sheet->mergeCells('A1:F1');
    $sheet->getRowDimension(1)->setRowHeight(30);
    $sheet->getStyle('A1:F1')->applyFromArray([
        'font' => ['bold' => true, 'size' => 14, 'color' => ['rgb' => 'FFFFFF']],
        'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => '1F4E78']],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
    ]);

    // Header row: bold on a light fill, boxed, centred, custom height.
    foreach (['Personnel', 'Amount', 'Start', 'Checked in', 'Days', 'Balance'] as $i => $label) {
        $sheet->setCellValue(COLUMNS[$i].'2', $label);
    }
    $sheet->getRowDimension(2)->setRowHeight(22);
    $sheet->getStyle('A2:F2')->applyFromArray([
        'font' => ['bold' => true],
        'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'DCE6F1']],
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['rgb' => '9BB7D4']]],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
    ]);

    // Column widths and the frozen pane below the header.
    foreach ([1 => 28.5, 2 => 15, 3 => 14, 4 => 20, 5 => 10, 6 => 12] as $col => $width) {
        $sheet->getColumnDimension(COLUMNS[$col - 1])->setWidth($width);
    }
    $sheet->freezePane('A'.DATA_START_ROW);
    $sheet->setShowGridlines(false);

    $rows = dataRows();
    for ($n = 0; $n < $bodyRows; $n++) {
        $row = DATA_START_ROW + $n;
        $sheet->getRowDimension($row)->setRowHeight(18.75);

        // Alternating body style — the zebra a template ships as its two
        // sample rows, and the thing the streamed variant must reproduce.
        $zebra = $n % 2 === 1;
        $sheet->getStyle('A'.$row.':F'.$row)->applyFromArray([
            'font' => ['size' => 11, 'color' => ['rgb' => $zebra ? '1F3864' : '000000']],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => $zebra ? 'EEF3F8' : 'FFFFFF'],
            ],
            'borders' => ['bottom' => ['borderStyle' => Border::BORDER_HAIR, 'color' => ['rgb' => 'BFBFBF']]],
        ]);
        $sheet->getStyle('B'.$row)->getNumberFormat()->setFormatCode('#,##0.00');
        $sheet->getStyle('C'.$row)->getNumberFormat()->setFormatCode('dd/mm/yyyy');
        $sheet->getStyle('D'.$row)->getNumberFormat()->setFormatCode('dd/mm/yyyy hh:mm');
        $sheet->getStyle('E'.$row)->getNumberFormat()->setFormatCode('0.0');
        $sheet->getStyle('F'.$row)->getNumberFormat()->setFormatCode('0');
        $sheet->getStyle('B'.$row.':F'.$row)->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_RIGHT);

        if (! $withData) {
            continue;
        }

        $values = $rows[$n];
        $sheet->setCellValueExplicit('A'.$row, $values[0], DataType::TYPE_STRING);
        $sheet->setCellValue('B'.$row, $values[1]);
        $sheet->setCellValue('C'.$row, ExcelDate::PHPToExcel($values[2]));
        $sheet->setCellValue('D'.$row, ExcelDate::PHPToExcel($values[3]));
        $sheet->setCellValue('E'.$row, $values[4]);
        $sheet->setCellValue('F'.$row, $values[5]);
    }

    return $book;
}

function save(Spreadsheet $book, string $path): void
{
    (new XlsxWriter($book))->save($path);
    $book->disconnectWorksheets();
}

// ---------------------------------------------------------------
// Path A — PhpSpreadsheet writes the whole thing
// ---------------------------------------------------------------
$referencePath = $tmp.'/reference.xlsx';
memory_reset_peak_usage();
$baseline = memory_get_usage(true);
$referenceStart = microtime(true);
save(buildLayout(DATA_ROWS, withData: true), $referencePath);
$referenceSeconds = microtime(true) - $referenceStart;
$referenceMemory = memory_get_peak_usage(true) - $baseline;

// ---------------------------------------------------------------
// Path B — PhpSpreadsheet writes the layout, xlsx-stream writes the rows
// ---------------------------------------------------------------
$templatePath = $tmp.'/template.xlsx';
save(buildLayout(2, withData: false), $templatePath);

$streamedPath = $tmp.'/streamed.xlsx';
memory_reset_peak_usage();
$baseline = memory_get_usage(true);
$streamStart = microtime(true);
$template = Template::open($templatePath);
$writer = SinkableXlsxWriter::fromTemplate(new FileSink($streamedPath), $template);
$writer->sheet('Leaves', DATA_START_ROW);
$writer->writeRows(dataRows(), fn (array $row, int $i): int => $i % 2);
$writer->finishFile();
$template->close();
$streamSeconds = microtime(true) - $streamStart;
$streamMemory = memory_get_peak_usage(true) - $baseline;

// ---------------------------------------------------------------
// Compare, property by property
// ---------------------------------------------------------------

/** Every observable property of one cell, as a comparable array. */
function cellFacts(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet, string $ref): array
{
    $cell = $sheet->getCell($ref);
    $style = $sheet->getStyle($ref);
    $font = $style->getFont();
    $borders = $style->getBorders();

    $value = $cell->getValue();
    if (is_object($value)) {
        $value = (string) $value; // RichText from an inline string
    }
    // Numeric cells are compared as numbers, not as PHP values. Both files
    // hold the same quantity; PhpSpreadsheet writes a whole float with a
    // trailing ".0" to preserve its own PHP type, which PHP's float-to-
    // string drops, so a strict comparison would report int 549 against
    // float 549.0 and call an identical number a difference. Excel reads
    // both as the same cell. Rounding also absorbs the last bits of a
    // fractional date serial. Text cells stay strings, so a numeric-looking
    // string is still compared strictly.
    if ($cell->getDataType() === DataType::TYPE_NUMERIC && is_numeric($value)) {
        $value = round((float) $value, 9);
    }

    return [
        'value' => $value,
        'type' => $cell->getDataType(),
        'numFmt' => $style->getNumberFormat()->getFormatCode(),
        'bold' => $font->getBold(),
        'size' => $font->getSize(),
        'color' => $font->getColor()->getRGB(),
        'fillType' => $style->getFill()->getFillType(),
        'fillColor' => $style->getFill()->getStartColor()->getRGB(),
        'borderBottom' => $borders->getBottom()->getBorderStyle(),
        'borderTop' => $borders->getTop()->getBorderStyle(),
        'borderLeft' => $borders->getLeft()->getBorderStyle(),
        'borderRight' => $borders->getRight()->getBorderStyle(),
        'horizontal' => $style->getAlignment()->getHorizontal(),
        'vertical' => $style->getAlignment()->getVertical(),
    ];
}

try {
    $reference = IOFactory::load($referencePath)->getActiveSheet();
    $streamed = IOFactory::load($streamedPath)->getActiveSheet();

    $diffs = [];
    $lastRow = DATA_START_ROW + DATA_ROWS - 1;

    // 1-7: per-cell value, type, number format, font, fill, border, alignment.
    for ($row = 1; $row <= $lastRow; $row++) {
        foreach (COLUMNS as $column) {
            $ref = $column.$row;
            $a = cellFacts($reference, $ref);
            $b = cellFacts($streamed, $ref);
            foreach ($a as $property => $expected) {
                if ($b[$property] !== $expected) {
                    $diffs[] = sprintf(
                        '%s %s: expected %s, got %s',
                        $ref,
                        $property,
                        var_export($expected, true),
                        var_export($b[$property], true)
                    );
                }
            }
            if (count($diffs) > 25) {
                break 2;
            }
        }
    }

    // 8: merges.
    $mergeA = array_values($reference->getMergeCells());
    $mergeB = array_values($streamed->getMergeCells());
    sort($mergeA);
    sort($mergeB);
    if ($mergeA !== $mergeB) {
        $diffs[] = 'merges: expected '.implode(',', $mergeA).' got '.implode(',', $mergeB);
    }

    // 9: column widths.
    foreach (COLUMNS as $column) {
        $expected = $reference->getColumnDimension($column)->getWidth();
        $actual = $streamed->getColumnDimension($column)->getWidth();
        if (abs($expected - $actual) > 1e-9) {
            $diffs[] = "width {$column}: expected {$expected}, got {$actual}";
        }
    }

    // 10: row heights, header block and body alike.
    foreach ([1, 2, DATA_START_ROW, DATA_START_ROW + 1, $lastRow] as $row) {
        $expected = $reference->getRowDimension($row)->getRowHeight();
        $actual = $streamed->getRowDimension($row)->getRowHeight();
        if (abs($expected - $actual) > 1e-9) {
            $diffs[] = "height row {$row}: expected {$expected}, got {$actual}";
        }
    }

    // 11: frozen pane.
    if ($reference->getFreezePane() !== $streamed->getFreezePane()) {
        $diffs[] = 'freeze pane: expected '.var_export($reference->getFreezePane(), true)
            .' got '.var_export($streamed->getFreezePane(), true);
    }

    // 12: gridlines.
    if ($reference->getShowGridlines() !== $streamed->getShowGridlines()) {
        $diffs[] = 'gridlines: expected '.var_export($reference->getShowGridlines(), true)
            .' got '.var_export($streamed->getShowGridlines(), true);
    }

    // The row counts must agree too, or a diff-free prefix would pass.
    if ($reference->getHighestDataRow() !== $streamed->getHighestDataRow()) {
        $diffs[] = 'row count: expected '.$reference->getHighestDataRow()
            .' got '.$streamed->getHighestDataRow();
    }

    if ($diffs === []) {
        pass(sprintf(
            'template parity — %d rows x %d columns, 12 properties, 0 differences',
            DATA_ROWS,
            count(COLUMNS)
        ));
    } else {
        fail($failures, 'template parity', count($diffs).' difference(s): '.implode(' | ', array_slice($diffs, 0, 10)));
    }
} catch (Throwable $e) {
    fail($failures, 'template parity', $e->getMessage());
}

// ---------------------------------------------------------------
// Cost, on the same layout and the same data
// ---------------------------------------------------------------
fwrite(STDOUT, sprintf(
    "\n  layout by PhpSpreadsheet, rows by PhpSpreadsheet: %6.3f s  %6.1f MB  %s\n".
    "  layout by PhpSpreadsheet, rows by xlsx-stream:     %6.3f s  %6.1f MB  %s\n".
    "  speed %.1fx, memory %.1fx\n\n",
    $referenceSeconds,
    $referenceMemory / 1048576,
    number_format(filesize($referencePath)).' B',
    $streamSeconds,
    $streamMemory / 1048576,
    number_format(filesize($streamedPath)).' B',
    $referenceSeconds / max($streamSeconds, 1e-9),
    $referenceMemory / max($streamMemory, 1024)
));

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

fwrite(STDOUT, "All template-mode parity checks passed.\n");
exit(0);

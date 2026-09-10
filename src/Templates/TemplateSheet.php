<?php

namespace Kolay\XlsxStream\Templates;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Readers\CellTokenizer;
use Kolay\XlsxStream\Readers\SharedStrings;

/**
 * One template sheet cut at its `<sheetData>` seam.
 *
 * The cut is the whole of template mode's contact with a foreign sheet:
 * everything before `<sheetData>` is `head`, the rows above `dataStartRow`
 * are kept verbatim as `headerRowsXml`, the rows from `dataStartRow` on are
 * consumed as SAMPLES (never emitted), and everything from `</sheetData>` on
 * is `tail`, copied byte-for-byte. The package does not interpret the sheet;
 * it opens one seam and closes it.
 *
 * A sample row is a **style-id oracle**: its cells' `s="…"` attributes say
 * how the producing application encoded that row's look, so the writer can
 * stamp the same ids on streamed rows without knowing anything about fonts,
 * fills or number formats. One sample row per variant (zebra striping, a
 * "missing" highlight) — `variant` selects among them at write time.
 *
 * Parsing is a hand-written forward scan rather than a regex with a lazy
 * quantifier: sheet XML is attacker-reachable input and the package does not
 * put backtracking-prone patterns anywhere it parses documents.
 */
class TemplateSheet
{
    /**
     * A template is a layout, not a report. Past this many sample rows the
     * caller almost certainly passed a full data file by mistake, and each
     * sample would cost a parsed style map.
     */
    public const MAX_SAMPLE_ROWS = 256;

    /** The SpreadsheetML namespace streamed rows must land in to be valid. */
    private const MAIN_NAMESPACE = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    private function __construct(
        private string $head,
        private string $headerRowsXml,
        private string $tail,
        private int $dataStartRow,
        /** @var list<array<int, int>> variant => 0-based column => cellXfs id */
        private array $styleMaps,
        /** @var list<array{ht: ?string, customHeight: bool, s: ?int, customFormat: bool}> */
        private array $rowAttributes,
        private string $prefix,
        private bool $acceptsUnprefixedRows,
    ) {
    }

    /** XML from the document start through the opening `<sheetData>` tag. */
    public function head(): string
    {
        return $this->head;
    }

    /** The rows above `dataStartRow`, verbatim (may be empty). */
    public function headerRowsXml(): string
    {
        return $this->headerRowsXml;
    }

    /** `</sheetData>` through the document end, verbatim. */
    public function tail(): string
    {
        return $this->tail;
    }

    /** First sheet row number the streamed data occupies. */
    public function dataStartRow(): int
    {
        return $this->dataStartRow;
    }

    /** How many style variants the template offers (always ≥ 1). */
    public function variantCount(): int
    {
        return \count($this->styleMaps);
    }

    /**
     * 0-based column index => cellXfs id for one variant. Columns the sample
     * left unstyled are absent, so a caller writing past the sample's width
     * simply gets no `s=` for those cells.
     *
     * @return array<int, int>
     */
    public function styleMap(int $variant): array
    {
        return $this->styleMaps[$this->normaliseVariant($variant)];
    }

    /**
     * Row-level attributes copied from the sample: height (`ht`), whether it
     * is a custom height, the row style id and custom-format flag.
     *
     * @return array{ht: ?string, customHeight: bool, s: ?int, customFormat: bool}
     */
    public function rowAttributes(int $variant): array
    {
        return $this->rowAttributes[$this->normaliseVariant($variant)];
    }

    /**
     * The header rows' cell values, keyed by sheet row number, each row a
     * 0-based column => value array.
     *
     * Only the random-access index needs these: the template's header rows
     * are real rows inside the streamed sheet's first block, so a zone map
     * built from the data alone could prune a block that a full scan would
     * have matched. Tokenizing them with the reader's own tokenizer folds in
     * exactly what a reader will see — a sheet whose cells the tokenizer
     * cannot read yields nothing here, and the un-pruned path reads nothing
     * from it either, so the two paths still agree.
     *
     * @return array<int, array<int, mixed>>
     */
    public function headerRowValues(?SharedStrings $sharedStrings = null): array
    {
        $values = [];
        foreach (self::scanRows($this->headerRowsXml, $this->prefix) as $row) {
            $values[$row['index']] = CellTokenizer::tokenizeRow($row['xml'], $sharedStrings);
        }

        return $values;
    }

    /** Namespace prefix the sheet uses on its elements ('' or e.g. 'x:'). */
    public function elementPrefix(): string
    {
        return $this->prefix;
    }

    /**
     * Whether the writer may stream plain `<row>`/`<c>` elements into this
     * sheet — true exactly when the document binds the DEFAULT namespace to
     * SpreadsheetML, so unprefixed elements land in it.
     *
     * The row builders emit unprefixed tags (the classic hot path must not
     * pay for prefix concatenation), so a sheet that only binds a prefix —
     * `<x:worksheet xmlns:x="…main">` — would take our rows into no
     * namespace at all and Excel would offer to repair the file. Template
     * mode refuses that dialect rather than opening a fifth seam to inject
     * an xmlns at the root. A document that binds BOTH (prefixed elements
     * plus a default xmlns) is perfectly writable and is accepted.
     */
    public function acceptsUnprefixedRows(): bool
    {
        return $this->acceptsUnprefixedRows;
    }

    private function normaliseVariant(int $variant): int
    {
        $count = \count($this->styleMaps);

        return ($variant >= 0 && $variant < $count) ? $variant : 0;
    }

    // ─────────────────────────────────────────────────────────────────────
    // Parsing
    // ─────────────────────────────────────────────────────────────────────

    /**
     * Cut a sheet XML document at its `<sheetData>` seam.
     *
     * @param  string  $entry  zip entry name, for error messages
     */
    public static function parse(string $xml, int $dataStartRow, string $entry): self
    {
        if ($dataStartRow < 1) {
            throw new XlsxStreamException("dataStartRow must be >= 1, got {$dataStartRow}.");
        }

        [$prefix, $openStart, $openEnd, $selfClosing] = self::locateSheetData($xml, $entry);

        if ($selfClosing) {
            // Normalise `<sheetData/>` into the open/close pair the writer
            // needs to stream between.
            $head = substr($xml, 0, $openStart).'<'.$prefix.'sheetData>';
            $rowsRegion = '';
            $tail = '</'.$prefix.'sheetData>'.substr($xml, $openEnd);
        } else {
            $head = substr($xml, 0, $openEnd);
            $closeStart = self::locateSheetDataClose($xml, $prefix, $openEnd, $entry);
            $rowsRegion = substr($xml, $openEnd, $closeStart - $openEnd);
            $tail = substr($xml, $closeStart);
        }

        // <dimension> spans the whole used range, which a streaming write
        // cannot know when head goes out — and it is optional in the schema
        // (our own preamble never emits one). Drop it; Excel recomputes.
        $head = self::stripDimension($head);

        self::assertNoRangeInDataRegion($tail, $dataStartRow);

        $rows = self::scanRows($rowsRegion, $prefix);

        $headerRows = '';
        $samples = [];
        foreach ($rows as $row) {
            if ($row['index'] < $dataStartRow) {
                $headerRows .= $row['xml'];
            } else {
                $samples[] = $row;
            }
        }

        if (\count($samples) > self::MAX_SAMPLE_ROWS) {
            throw XlsxStreamException::templateTooManySampleRows(\count($samples), self::MAX_SAMPLE_ROWS);
        }

        $styleMaps = [];
        $rowAttributes = [];
        foreach ($samples as $sample) {
            $styleMaps[] = self::scanCellStyles($sample['body'], $prefix);
            $rowAttributes[] = self::parseRowAttributes($sample['attrs']);
        }

        if ($styleMaps === []) {
            // No sample rows: one unstyled variant so writeRow() always has a
            // variant to resolve against.
            $styleMaps[] = [];
            $rowAttributes[] = ['ht' => null, 'customHeight' => false, 's' => null, 'customFormat' => false];
        }

        return new self(
            $head,
            $headerRows,
            $tail,
            $dataStartRow,
            $styleMaps,
            $rowAttributes,
            $prefix,
            self::bindsDefaultSpreadsheetNamespace($head)
        );
    }

    /**
     * Find the `<sheetData>` opening tag, tolerating any namespace prefix.
     *
     * @return array{0: string, 1: int, 2: int, 3: bool} prefix, tag start, tag end (exclusive), self-closing
     */
    private static function locateSheetData(string $xml, string $entry): array
    {
        if (! preg_match('/<((?:[A-Za-z_][\w.\-]*:)?)sheetData(\s[^>]*?)?(\/?)>/', $xml, $m, PREG_OFFSET_CAPTURE)) {
            throw XlsxStreamException::templateSheetDataMissing($entry);
        }

        return [
            $m[1][0],
            $m[0][1],
            $m[0][1] + \strlen($m[0][0]),
            $m[3][0] === '/',
        ];
    }

    private static function locateSheetDataClose(string $xml, string $prefix, int $from, string $entry): int
    {
        if (! preg_match('/<\/'.preg_quote($prefix, '/').'sheetData\s*>/', $xml, $m, PREG_OFFSET_CAPTURE, $from)) {
            throw XlsxStreamException::templateSheetDataMissing($entry);
        }

        return $m[0][1];
    }

    private static function stripDimension(string $head): string
    {
        $stripped = preg_replace(
            '/<(?:[A-Za-z_][\w.\-]*:)?dimension\b[^>]*(?:\/>|>.*?<\/(?:[A-Za-z_][\w.\-]*:)?dimension\s*>)/s',
            '',
            $head,
            1
        );

        return $stripped ?? $head;
    }

    /**
     * Reject any tail range that reaches into the streamed rows.
     *
     * The tail is copied verbatim, so a range the template author drew over
     * the sample rows keeps exactly those rows once real data is streamed: a
     * merge that lands mid-data, a filter that stops after four rows, a
     * conditional format that colours only the sample. Rewriting them would
     * mean editing the tail against a row count known only at the end, and
     * for a table or a filter defined-name it would mean editing a different
     * part altogether — the second seam this design refuses to open.
     *
     * The check reads no semantics. It walks the tail's start tags and looks
     * at every `ref` and `sqref` attribute, plus the element form the x14
     * conditional-formatting extension uses, and refuses when any corner
     * falls on dataStartRow or below. A range with no row at all (whole
     * columns, `A:C`) covers every row and is refused too.
     *
     * Only the tail is scanned, and deliberately so. The head carries ranges
     * that describe a viewport rather than a region of data — `<selection
     * sqref="A5"/>` is where the author left the cursor — and refusing a
     * template over the saved cursor position would be a false rejection,
     * the failure that makes this mode unusable on real layouts.
     */
    private static function assertNoRangeInDataRegion(string $tail, int $dataStartRow): void
    {
        $i = 0;
        while (($i = strpos($tail, '<', $i)) !== false) {
            $end = self::tagEnd($tail, $i + 1);
            if ($end === false) {
                break;
            }

            $tag = substr($tail, $i + 1, $end - $i - 1);
            $i = $end + 1;

            if ($tag === '' || $tag[0] === '/' || $tag[0] === '!' || $tag[0] === '?') {
                continue;
            }

            $nameLength = strcspn($tag, " \t\r\n/>");
            $name = substr($tag, 0, $nameLength);

            // A table keeps its range in xl/tables/tableN.xml and its filter
            // in a workbook defined name, both of which are carried across
            // untouched, so the table would cover the sample rows only.
            //
            // The trigger is a <tablePart> child, never the <tableParts>
            // wrapper: PhpSpreadsheet 1.x writes an empty <tableParts
            // count="0"/> on every sheet it produces, so refusing the wrapper
            // would refuse every template that library made — the false
            // rejection this guard exists to avoid, not to cause.
            if (self::localName($name) === 'tablePart') {
                throw XlsxStreamException::templateTablePartsUnsupported();
            }

            if (preg_match_all('/(?<![A-Za-z])(?:sqref|ref)="([^"]*)"/', substr($tag, $nameLength), $matches)) {
                foreach ($matches[1] as $refs) {
                    self::assertRefsAboveDataRegion($name, $refs, $dataStartRow);
                }
            }
        }

        // <xm:sqref>A3:C4</xm:sqref> — the extension list spells its ranges
        // as element text rather than an attribute.
        if (preg_match_all('/<(?:[A-Za-z_][\w.\-]*:)?sqref\b[^>]*>([^<]*)</', $tail, $matches)) {
            foreach ($matches[1] as $refs) {
                self::assertRefsAboveDataRegion('sqref', trim($refs), $dataStartRow);
            }
        }
    }

    /**
     * Refuse a whitespace-separated reference list whose any corner reaches
     * dataStartRow or below. A corner carrying no row number spans the whole
     * column, which reaches the data region by definition.
     */
    private static function assertRefsAboveDataRegion(string $element, string $refs, int $dataStartRow): void
    {
        foreach (preg_split('/\s+/', trim($refs), -1, PREG_SPLIT_NO_EMPTY) ?: [] as $ref) {
            foreach (explode(':', $ref) as $corner) {
                if (! preg_match('/(\d+)\s*$/', $corner, $row)) {
                    throw XlsxStreamException::templateRangeInDataRegion($element, $ref, $dataStartRow);
                }
                if ((int) $row[1] >= $dataStartRow) {
                    throw XlsxStreamException::templateRangeInDataRegion($element, $ref, $dataStartRow);
                }
            }
        }
    }

    /** Element name with any namespace prefix removed. */
    private static function localName(string $name): string
    {
        $colon = strrpos($name, ':');

        return $colon === false ? $name : substr($name, $colon + 1);
    }

    /**
     * Forward scan over `<row>` elements. Rows carry their number in `r`;
     * a producer that omits it (our own compact mode) is numbered
     * positionally from the previous row.
     *
     * @return list<array{xml: string, attrs: string, body: string, index: int}>
     */
    private static function scanRows(string $region, string $prefix): array
    {
        $open = '<'.$prefix.'row';
        $close = '</'.$prefix.'row';
        $openLen = \strlen($open);

        $rows = [];
        $cursor = 0;
        $previousIndex = 0;

        while (true) {
            $start = strpos($region, $open, $cursor);
            if ($start === false) {
                break;
            }
            if (! self::isTagBoundary($region[$start + $openLen] ?? '')) {
                $cursor = $start + $openLen; // e.g. <rowBreaks>
                continue;
            }

            $tagEnd = self::tagEnd($region, $start + $openLen);
            if ($tagEnd === false) {
                break;
            }
            $selfClosing = ($region[$tagEnd - 1] ?? '') === '/';
            $attrs = substr(
                $region,
                $start + $openLen,
                $tagEnd - ($start + $openLen) - ($selfClosing ? 1 : 0)
            );

            if ($selfClosing) {
                $end = $tagEnd + 1;
                $body = '';
            } else {
                $closePos = strpos($region, $close, $tagEnd);
                if ($closePos === false) {
                    break;
                }
                $closeEnd = self::tagEnd($region, $closePos + \strlen($close));
                if ($closeEnd === false) {
                    break;
                }
                $body = substr($region, $tagEnd + 1, $closePos - $tagEnd - 1);
                $end = $closeEnd + 1;
            }

            $index = preg_match('/(?<![A-Za-z])r="(\d+)"/', $attrs, $m)
                ? (int) $m[1]
                : $previousIndex + 1;
            $previousIndex = $index;

            $rows[] = [
                'xml' => substr($region, $start, $end - $start),
                'attrs' => $attrs,
                'body' => $body,
                'index' => $index,
            ];
            $cursor = $end;
        }

        return $rows;
    }

    /**
     * Sample row → 0-based column => cellXfs id. Only styled cells land in
     * the map; an unstyled sample cell means "no `s=` for this column".
     *
     * @return array<int, int>
     */
    private static function scanCellStyles(string $body, string $prefix): array
    {
        $open = '<'.$prefix.'c';
        $close = '</'.$prefix.'c';
        $openLen = \strlen($open);

        $map = [];
        $cursor = 0;
        $previousColumn = -1;

        while (true) {
            $start = strpos($body, $open, $cursor);
            if ($start === false) {
                break;
            }
            if (! self::isTagBoundary($body[$start + $openLen] ?? '')) {
                $cursor = $start + $openLen; // e.g. <col…> is not <c…>
                continue;
            }

            $tagEnd = self::tagEnd($body, $start + $openLen);
            if ($tagEnd === false) {
                break;
            }
            $selfClosing = ($body[$tagEnd - 1] ?? '') === '/';
            $attrs = substr(
                $body,
                $start + $openLen,
                $tagEnd - ($start + $openLen) - ($selfClosing ? 1 : 0)
            );

            $column = preg_match('/(?<![A-Za-z])r="([A-Z]+)\d*"/', $attrs, $m)
                ? CellTokenizer::columnLettersToIndex($m[1])
                : $previousColumn + 1;
            $previousColumn = $column;

            if (preg_match('/(?<![A-Za-z])s="(\d+)"/', $attrs, $m)) {
                $map[$column] = (int) $m[1];
            }

            if ($selfClosing) {
                $cursor = $tagEnd + 1;

                continue;
            }
            $closePos = strpos($body, $close, $tagEnd);
            $closeEnd = $closePos === false ? false : self::tagEnd($body, $closePos + \strlen($close));
            $cursor = $closeEnd === false ? $tagEnd + 1 : $closeEnd + 1;
        }

        ksort($map);

        return $map;
    }

    /**
     * Row attributes worth carrying onto streamed rows. Each lookup is
     * lookbehind-guarded so `customHeight="1"` cannot be read as `ht="1"`
     * and `spans="1:4"` cannot be read as the row style `s="…"`.
     *
     * @return array{ht: ?string, customHeight: bool, s: ?int, customFormat: bool}
     */
    private static function parseRowAttributes(string $attrs): array
    {
        return [
            'ht' => preg_match('/(?<![A-Za-z])ht="([^"]*)"/', $attrs, $m) ? $m[1] : null,
            'customHeight' => (bool) preg_match('/(?<![A-Za-z])customHeight="(?:1|true)"/', $attrs),
            's' => preg_match('/(?<![A-Za-z])s="(\d+)"/', $attrs, $m) ? (int) $m[1] : null,
            'customFormat' => (bool) preg_match('/(?<![A-Za-z])customFormat="(?:1|true)"/', $attrs),
        ];
    }

    /**
     * Does the root element bind the default namespace to SpreadsheetML?
     * Only the root can, in practice — `<sheetData>`'s only ancestor is
     * `<worksheet>` — so one look at the opening tag settles it.
     */
    private static function bindsDefaultSpreadsheetNamespace(string $head): bool
    {
        if (! preg_match('/<(?:[A-Za-z_][\\w.\\-]*:)?worksheet\\b([^>]*)>/', $head, $m)) {
            return false;
        }

        return (bool) preg_match('/\\bxmlns="'.preg_quote(self::MAIN_NAMESPACE, '/').'"/', $m[1]);
    }

    private static function isTagBoundary(string $char): bool
    {
        return $char === ' ' || $char === '>' || $char === '/' || $char === "\t" || $char === "\n" || $char === "\r";
    }

    /**
     * Offset of the '>' closing the tag that starts at $from, respecting
     * quoted attribute values. Same walk CellTokenizer uses.
     */
    private static function tagEnd(string $s, int $from): int|false
    {
        $i = $from;
        $len = \strlen($s);

        while ($i < $len) {
            $i += strcspn($s, '>"\'', $i);
            if ($i >= $len) {
                return false;
            }
            if ($s[$i] === '>') {
                return $i;
            }
            $closing = strpos($s, $s[$i], $i + 1);
            if ($closing === false) {
                return false;
            }
            $i = $closing + 1;
        }

        return false;
    }
}

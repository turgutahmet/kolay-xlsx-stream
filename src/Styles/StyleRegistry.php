<?php

namespace Kolay\XlsxStream\Styles;

/**
 * Builds the dynamic portion of xl/styles.xml.
 *
 * The writer registers logical styles (header, column formats) here and gets
 * back stable cellXfs indexes (style ids) that it can stamp onto cells via
 * `s="N"`. Styles that aren't used aren't emitted, so the styles.xml stays
 * tight regardless of how many presets are exposed.
 *
 * Indexes are append-only and stable for the lifetime of a writer — that lets
 * the row builder cache style ids per column without invalidation.
 */
class StyleRegistry
{
    /** Excel built-in numFmtIds (no <numFmt> entry needed). */
    public const BUILTIN_NUMFMT_GENERAL = 0;
    public const BUILTIN_NUMFMT_INT = 1;        // 0
    public const BUILTIN_NUMFMT_TWO_DECIMAL = 2; // 0.00
    public const BUILTIN_NUMFMT_DATE = 14;      // m/d/yyyy
    public const BUILTIN_NUMFMT_DATETIME = 22;  // m/d/yyyy h:mm

    /** Custom numFmtIds start at 164 per OOXML spec. */
    private const CUSTOM_NUMFMT_START = 164;

    /**
     * Named format presets. Values are Excel format codes.
     *
     * Tip: locale-specific currency presets (`currency_try` etc.) use the
     * literal symbol so the file is self-contained — no need for the user's
     * Excel locale to be set.
     */
    public const PRESETS = [
        'date' => 'yyyy-mm-dd',
        'datetime' => 'yyyy-mm-dd hh:mm:ss',
        'datetime_iso' => 'yyyy-mm-dd"T"hh:mm:ss',
        'time' => 'hh:mm:ss',
        'integer' => '#,##0',
        'decimal' => '#,##0.00',
        'percent' => '0.00%',
        'currency_try' => '#,##0.00\\ ₺',
        'currency_usd' => '$#,##0.00',
        'currency_eur' => '€#,##0.00',
        'currency_gbp' => '£#,##0.00',
    ];

    /** @var array<string, int> formatCode => numFmtId */
    private array $customNumFmts = [];

    /** @var array<int, array{numFmtId:int, fontId:int, fillId:int, applyNumberFormat:int, applyFont:int, applyFill:int}> */
    private array $cellXfs = [
        // index 0 = default, index 1 = preserved historical datetime style
        ['numFmtId' => 0, 'fontId' => 0, 'fillId' => 0, 'applyNumberFormat' => 0, 'applyFont' => 0, 'applyFill' => 0],
        ['numFmtId' => 164, 'fontId' => 0, 'fillId' => 0, 'applyNumberFormat' => 1, 'applyFont' => 0, 'applyFill' => 0],
    ];

    /** @var array<int, array{bold:bool, color:?string, size:int, name:string}> */
    private array $fonts = [
        ['bold' => false, 'color' => null, 'size' => 11, 'name' => 'Calibri'],
    ];

    /** @var array<int, ?string> indexed by fillId; null = no fill (use built-in 0/1). */
    private array $fills = [
        null, // built-in: <patternFill patternType="none"/>
        null, // built-in: <patternFill patternType="gray125"/>
    ];

    /**
     * Template mode: the seed stylesheet, kept VERBATIM. A template carries
     * borders, cellStyleXfs, dxfs, tableStyles and colors this package does
     * not model, so a seeded registry never regenerates the document — it
     * appends into it and leaves every other byte alone.
     */
    private ?string $seedXml = null;

    private bool $seeded = false;

    /** Index of our first appended entry in each seeded table. */
    private int $fontOffset = 0;

    private int $fillOffset = 0;

    private int $xfOffset = 0;

    /** Highest numFmtId already used by the seed; ours continue above it. */
    private int $numFmtFloor = self::CUSTOM_NUMFMT_START - 1;

    /**
     * Seed the registry from a template's styles.xml.
     *
     * The seed is the authority. Registering nothing gives that stylesheet
     * back byte-for-byte — the common case, since most templates already
     * carry every style their sample rows reference. Registering a style
     * APPENDS a font/fill/xf to the end of the matching table and bumps its
     * count, so ids the template already handed out never shift.
     *
     * Entries are not de-duplicated against the seed's own (that would mean
     * parsing fonts and fills we deliberately treat as opaque), so a style
     * identical to one already in the template becomes a second entry —
     * harmless, and the price of not interpreting the file.
     */
    public static function fromStylesXml(string $xml): self
    {
        $registry = new self();
        $registry->seedXml = $xml;
        $registry->seeded = true;
        // Our tables now hold ONLY what we append; the seed owns the rest.
        $registry->customNumFmts = [];
        $registry->fonts = [];
        $registry->fills = [];
        $registry->cellXfs = [];
        $registry->fontOffset = self::seedCount($xml, 'fonts');
        $registry->fillOffset = self::seedCount($xml, 'fills');
        $registry->xfOffset = self::seedCount($xml, 'cellXfs');
        // NB: counted from the elements themselves — see seedCount().
        $registry->numFmtFloor = self::seedNumFmtFloor($xml);

        return $registry;
    }

    /**
     * True when this registry is a template seed nothing has been appended
     * to, so the template's own styles.xml can be moved into the output
     * instead of re-rendered.
     */
    public function isSeedUnmodified(): bool
    {
        return $this->seeded
            && $this->customNumFmts === []
            && $this->fonts === []
            && $this->fills === []
            && $this->cellXfs === [];
    }

    /** Seed table name => the child element whose occurrences define its size. */
    private const SEED_TABLE_CHILD = [
        'numFmts' => 'numFmt',
        'fonts' => 'font',
        'fills' => 'fill',
        'cellXfs' => 'xf',
    ];

    /**
     * Size of a seed table, counted from its CHILD ELEMENTS — never from the
     * `count` attribute.
     *
     * Style ids are positional, and `count` is optional in the schema. A
     * missing or stale attribute would put our first appended xf at index 0,
     * colliding with the template's own xf 0, and the file would open with
     * silently wrong styling rather than an error. Counting is scoped to the
     * block's own body because `<fill>` also lives inside `<dxf>` and `<xf>`
     * also lives inside `<cellStyleXfs>`.
     */
    private static function seedCount(string $xml, string $tag): int
    {
        return self::countChildren(self::blockBody($xml, $tag), self::SEED_TABLE_CHILD[$tag]);
    }

    /** Contents between a block's open and close tag, '' when absent or empty. */
    private static function blockBody(string $xml, string $tag): string
    {
        $open = '/<(?:[A-Za-z_][\w.\-]*:)?'.$tag.'\b[^>]*?(\/?)>/';
        if (! preg_match($open, $xml, $m, PREG_OFFSET_CAPTURE)) {
            return '';
        }
        if ($m[1][0] === '/') {
            return ''; // <fills/> — an empty table
        }

        $start = $m[0][1] + \strlen($m[0][0]);
        $close = '/<\/(?:[A-Za-z_][\w.\-]*:)?'.$tag.'\s*>/';
        if (! preg_match($close, $xml, $c, PREG_OFFSET_CAPTURE, $start)) {
            return '';
        }

        return substr($xml, $start, $c[0][1] - $start);
    }

    private static function countChildren(string $body, string $child): int
    {
        return preg_match_all('/<(?:[A-Za-z_][\w.\-]*:)?'.$child.'(?=[\s\/>])/', $body);
    }

    /**
     * Highest numFmtId the seed declares. Only <numFmt> declarations count —
     * an <xf numFmtId="165"> merely references one.
     */
    private static function seedNumFmtFloor(string $xml): int
    {
        $floor = self::CUSTOM_NUMFMT_START - 1;
        if (preg_match_all('/<numFmt\b[^>]*?\bnumFmtId="(\d+)"/', $xml, $m)) {
            foreach ($m[1] as $id) {
                $floor = max($floor, (int) $id);
            }
        }

        return $floor;
    }

    public function __construct()
    {
        // Reserve numFmtId 164 for the legacy datetime format used at cellXfs[1].
        $this->customNumFmts['yyyy-mm-dd hh:mm:ss'] = 164;
    }

    /**
     * Register a number format and return its cellXfs index (style id).
     *
     * Accepts a preset name (e.g. "date", "currency_try") or a raw Excel
     * format code (e.g. "0.000"). Same code → same style id (idempotent).
     * Pass $raw = true to skip preset resolution AND the typo guard —
     * the string is taken verbatim as an Excel format code.
     */
    public function registerColumnFormat(string $presetOrCode, bool $raw = false): int
    {
        $code = $raw ? $presetOrCode : (self::PRESETS[$presetOrCode] ?? $presetOrCode);

        // A string made purely of lowercase letters/underscores is a
        // preset NAME shape, almost never a raw format code (real codes
        // carry #, 0, %, punctuation or quoting). Letting a typo like
        // "currency" fall through as a literal formatCode produces a
        // file MS Excel refuses to open without repair — fail loudly at
        // write time instead of corrupting the workbook silently.
        //
        // The one legitimate pure-lowercase family: date-token runs like
        // "dddd" (weekday), "mmmm" (month name), "mmss", "yyyymmdd" —
        // strings built ONLY from the d/m/y/h/s token letters. Those are
        // valid codes and pass through; anything else lowercase either
        // names a preset or is a typo. $raw bypasses the whole check.
        if (! $raw
            && $code === $presetOrCode
            && preg_match('/^[a-z_]+$/', $presetOrCode)
            && ! preg_match('/^[dmyhs]+$/', $presetOrCode)
        ) {
            throw new \Kolay\XlsxStream\Exceptions\XlsxStreamException(
                "Unknown format preset '{$presetOrCode}'. Available presets: ".
                implode(', ', array_keys(self::PRESETS)).
                '. To use a raw Excel format code, pass the code itself (e.g. "#,##0.00") '.
                'or call setColumnFormat($col, $code, raw: true).'
            );
        }

        $numFmtId = $this->resolveNumFmtId($code);

        return $this->resolveCellXf([
            'numFmtId' => $numFmtId,
            'fontId' => 0,
            'fillId' => 0,
            'applyNumberFormat' => 1,
            'applyFont' => 0,
            'applyFill' => 0,
        ]);
    }

    /**
     * Register a built-in numFmtId (0-49 reserved range) without
     * synthesising a <numFmt> entry. Excel resolves the format code
     * locale-aware on the reader side — `BUILTIN_NUMFMT_DATE = 14`
     * shows mm-dd-yy in en-US, dd.mm.yyyy in tr-TR, etc.
     */
    public function registerBuiltinNumFmt(int $numFmtId): int
    {
        return $this->resolveCellXf([
            'numFmtId' => $numFmtId,
            'fontId' => 0,
            'fillId' => 0,
            'applyNumberFormat' => 1,
            'applyFont' => 0,
            'applyFill' => 0,
        ]);
    }

    /**
     * Register a header style and return its cellXfs index.
     *
     * Options: bold (bool), color (#RRGGBB), fill (#RRGGBB), size (int),
     * name (string — font family, default "Calibri").
     *
     * Color values are validated as 6-character hex (with or without a
     * leading #). Anything else throws so the caller catches the typo
     * here instead of producing an invalid xl/styles.xml that Excel
     * silently rejects.
     */
    public function registerHeaderStyle(array $options): int
    {
        if (isset($options['color'])) {
            $this->assertHexColor($options['color'], 'color');
        }
        if (isset($options['fill'])) {
            $this->assertHexColor($options['fill'], 'fill');
        }

        $fontId = $this->resolveFont([
            'bold' => (bool) ($options['bold'] ?? false),
            'color' => $options['color'] ?? null,
            'size' => (int) ($options['size'] ?? 11),
            'name' => (string) ($options['name'] ?? 'Calibri'),
        ]);

        $fillId = isset($options['fill']) ? $this->resolveFill($options['fill']) : 0;

        return $this->resolveCellXf([
            'numFmtId' => 0,
            'fontId' => $fontId,
            'fillId' => $fillId,
            'applyNumberFormat' => 0,
            'applyFont' => $fontId > 0 ? 1 : 0,
            'applyFill' => $fillId > 0 ? 1 : 0,
        ]);
    }

    /**
     * Register a per-row style and return its cellXfs index.
     *
     * Structurally identical to a header style — fill (background) + font
     * (color/bold/size/name) — but exposed separately so the intent reads
     * clearly at the call site:
     *
     *   $red = $writer->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
     *   $writer->writeRow($row, $failed ? $red : null);
     *
     * Same options → same style id (dedup via resolveCellXf), so painting a
     * million rows with one logical style still adds a single cellXfs entry.
     */
    public function registerRowStyle(array $options): int
    {
        return $this->registerHeaderStyle($options);
    }

    /**
     * Compose a row style (fill + font) with a column's number format and
     * return the merged cellXfs index.
     *
     * Needed because a styled row must still respect a column's numFmt —
     * a currency column on a highlighted row should stay "1.234,50 ₺", not
     * collapse to a raw number. Takes the numFmt from the column xf and the
     * font/fill from the row xf. Dedup'd, so each (row-style, column) pair
     * resolves to one entry no matter how many rows hit it.
     */
    public function mergeRowStyleWithColumn(int $rowStyleId, int $columnStyleId): int
    {
        $row = $this->cellXfs[$rowStyleId - $this->xfOffset] ?? null;
        $col = $this->cellXfs[$columnStyleId - $this->xfOffset] ?? null;

        // Defensive: an unknown id means the caller passed something we never
        // handed out. Fall back to whichever side we do know rather than
        // emitting a dangling s="N".
        if ($row === null) {
            return $columnStyleId;
        }
        if ($col === null) {
            return $rowStyleId;
        }

        return $this->resolveCellXf([
            'numFmtId' => $col['numFmtId'],
            'fontId' => $row['fontId'],
            'fillId' => $row['fillId'],
            'applyNumberFormat' => $col['applyNumberFormat'],
            'applyFont' => $row['applyFont'],
            'applyFill' => $row['applyFill'],
        ]);
    }

    /**
     * Reject anything that isn't a 6-character hex color, with or without
     * a leading "#". Surfaces typos at registration time instead of
     * producing a styles.xml Excel will refuse to open.
     */
    private function assertHexColor(string $value, string $optionName): void
    {
        if (! preg_match('/^#?[0-9a-fA-F]{6}$/', $value)) {
            throw new \Kolay\XlsxStream\Exceptions\XlsxStreamException(
                "Style option '{$optionName}' must be a 6-character hex color (e.g. '#4F81BD'); got: {$value}"
            );
        }
    }

    /**
     * Render xl/styles.xml.
     */
    public function toXml(): string
    {
        if ($this->seedXml !== null) {
            return $this->seededXml();
        }

        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $xml .= '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

        // <numFmts>
        if (! empty($this->customNumFmts)) {
            $xml .= '<numFmts count="'.count($this->customNumFmts).'">';
            foreach ($this->customNumFmts as $code => $id) {
                $xml .= '<numFmt numFmtId="'.$id.'" formatCode="'.htmlspecialchars($code, ENT_QUOTES | ENT_XML1).'"/>';
            }
            $xml .= '</numFmts>';
        }

        // <fonts>
        $xml .= '<fonts count="'.count($this->fonts).'">';
        foreach ($this->fonts as $font) {
            $xml .= self::renderFont($font);
        }
        $xml .= '</fonts>';

        // <fills> — first 2 are required built-ins
        $xml .= '<fills count="'.count($this->fills).'">';
        $xml .= '<fill><patternFill patternType="none"/></fill>';
        $xml .= '<fill><patternFill patternType="gray125"/></fill>';
        for ($i = 2; $i < count($this->fills); $i++) {
            $xml .= self::renderSolidFill((string) ($this->fills[$i] ?? ''));
        }
        $xml .= '</fills>';

        $xml .= '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>';

        // <cellXfs>
        $xml .= '<cellXfs count="'.count($this->cellXfs).'">';
        foreach ($this->cellXfs as $xf) {
            $xml .= self::renderXf($xf);
        }
        $xml .= '</cellXfs>';

        $xml .= '</styleSheet>';

        return $xml;
    }

    /**
     * Seeded output: the template's stylesheet with our appends spliced in.
     * Nothing registered → the seed comes back untouched, byte-for-byte.
     */
    private function seededXml(): string
    {
        $xml = $this->seedXml ?? '';

        if ($this->customNumFmts === [] && $this->fonts === [] && $this->fills === [] && $this->cellXfs === []) {
            return $xml;
        }

        $xml = $this->spliceNumFmts($xml);
        $xml = self::spliceInto($xml, 'fonts', array_map(self::renderFont(...), $this->fonts));
        $xml = self::spliceInto($xml, 'fills', array_map(
            static fn ($color): string => self::renderSolidFill((string) $color),
            $this->fills
        ));

        return self::spliceInto($xml, 'cellXfs', array_map(self::renderXf(...), $this->cellXfs));
    }

    /**
     * Append fragments to a seed table and bump its `count`. The block is
     * unique in a stylesheet, so the first match is the right one, and the
     * insert is a plain substr splice — never a regex replacement, whose
     * `$` and backslash escapes would mangle a format code.
     *
     * The refreshed `count` is derived from the elements present plus what we
     * add — not from the old attribute, which may be missing or stale — and is
     * created when the block never carried one.
     *
     * @param  list<string>  $fragments
     */
    private static function spliceInto(string $xml, string $tag, array $fragments): string
    {
        if ($fragments === []) {
            return $xml;
        }

        $total = self::countChildren(self::blockBody($xml, $tag), self::SEED_TABLE_CHILD[$tag])
            + count($fragments);
        $xml = self::withBlockCount($xml, $tag, $total);

        $close = '</'.$tag.'>';
        $at = strpos($xml, $close);
        if ($at === false) {
            return $xml;
        }

        return substr($xml, 0, $at).implode('', $fragments).substr($xml, $at);
    }

    /** Set the block's `count`, adding the attribute when it is absent. */
    private static function withBlockCount(string $xml, string $tag, int $total): string
    {
        $replaced = preg_replace_callback(
            '/(<(?:[A-Za-z_][\w.\-]*:)?'.$tag.'\b[^>]*?)\bcount="\d+"/',
            static fn (array $m): string => $m[1].'count="'.$total.'"',
            $xml,
            1,
            $hits
        );
        if ($hits > 0 && $replaced !== null) {
            return $replaced;
        }

        $added = preg_replace_callback(
            '/<(?:[A-Za-z_][\w.\-]*:)?'.$tag.'\b[^>]*?(?=\/?>)/',
            static fn (array $m): string => $m[0].' count="'.$total.'"',
            $xml,
            1
        );

        return $added ?? $xml;
    }

    /**
     * Splice our custom number formats in, creating the <numFmts> block when
     * the template has none — schema order puts it first inside styleSheet.
     */
    private function spliceNumFmts(string $xml): string
    {
        if ($this->customNumFmts === []) {
            return $xml;
        }

        $fragments = [];
        foreach ($this->customNumFmts as $code => $id) {
            $fragments[] = '<numFmt numFmtId="'.$id.'" formatCode="'
                .htmlspecialchars($code, ENT_QUOTES | ENT_XML1).'"/>';
        }

        if (preg_match('/<numFmts\b[^>]*>/', $xml)) {
            return self::spliceInto($xml, 'numFmts', $fragments);
        }

        $block = '<numFmts count="'.count($fragments).'">'.implode('', $fragments).'</numFmts>';
        $created = preg_replace_callback(
            '/<styleSheet\b[^>]*>/',
            static fn (array $m): string => $m[0].$block,
            $xml,
            1
        );

        return $created ?? $xml;
    }

    /** @param array{bold:bool, color:?string, size:int, name:string} $font */
    private static function renderFont(array $font): string
    {
        $xml = '<font>';
        $xml .= '<sz val="'.$font['size'].'"/>';
        if ($font['bold']) {
            $xml .= '<b/>';
        }
        if ($font['color'] !== null) {
            $xml .= '<color rgb="FF'.ltrim($font['color'], '#').'"/>';
        }
        $xml .= '<name val="'.htmlspecialchars($font['name'], ENT_QUOTES | ENT_XML1).'"/>';

        return $xml.'</font>';
    }

    private static function renderSolidFill(string $color): string
    {
        return '<fill><patternFill patternType="solid"><fgColor rgb="FF'
            .ltrim($color, '#').'"/></patternFill></fill>';
    }

    /**
     * borderId="0" is the OOXML convention for "no border" and is what every
     * producer puts at index 0, so an appended xf inherits the seed's empty
     * border rather than one of its decorated ones.
     *
     * @param array{numFmtId:int, fontId:int, fillId:int, applyNumberFormat:int, applyFont:int, applyFill:int} $xf
     */
    private static function renderXf(array $xf): string
    {
        $xml = '<xf';
        $xml .= ' numFmtId="'.$xf['numFmtId'].'"';
        $xml .= ' fontId="'.$xf['fontId'].'"';
        $xml .= ' fillId="'.$xf['fillId'].'"';
        $xml .= ' borderId="0"';
        $xml .= ' xfId="0"';
        if ($xf['applyNumberFormat']) {
            $xml .= ' applyNumberFormat="1"';
        }
        if ($xf['applyFont']) {
            $xml .= ' applyFont="1"';
        }
        if ($xf['applyFill']) {
            $xml .= ' applyFill="1"';
        }

        return $xml.'/>';
    }

    private function resolveNumFmtId(string $code): int
    {
        if (isset($this->customNumFmts[$code])) {
            return $this->customNumFmts[$code];
        }

        $id = $this->seeded
            ? $this->numFmtFloor + 1 + count($this->customNumFmts)
            : self::CUSTOM_NUMFMT_START + count($this->customNumFmts);
        $this->customNumFmts[$code] = $id;

        return $id;
    }

    private function resolveCellXf(array $xf): int
    {
        // Strict comparison is safe: every cellXfs entry is constructed
        // with the same key order in registerHeaderStyle / registerColumnFormat,
        // so === is both faster and semantically more correct than ==.
        foreach ($this->cellXfs as $i => $existing) {
            if ($existing === $xf) {
                return $this->xfOffset + $i;
            }
        }
        $this->cellXfs[] = $xf;

        return $this->xfOffset + count($this->cellXfs) - 1;
    }

    private function resolveFont(array $font): int
    {
        foreach ($this->fonts as $i => $existing) {
            if ($existing === $font) {
                return $this->fontOffset + $i;
            }
        }
        $this->fonts[] = $font;

        return $this->fontOffset + count($this->fonts) - 1;
    }

    private function resolveFill(string $color): int
    {
        $color = ltrim($color, '#');
        foreach ($this->fills as $i => $existing) {
            if ($this->fillOffset === 0 && $i < 2) {
                continue; // skip the two built-ins the classic table starts with
            }
            if (ltrim((string) $existing, '#') === $color) {
                return $this->fillOffset + $i;
            }
        }
        $this->fills[] = $color;

        return $this->fillOffset + count($this->fills) - 1;
    }
}

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
     */
    public function registerColumnFormat(string $presetOrCode): int
    {
        $code = self::PRESETS[$presetOrCode] ?? $presetOrCode;

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
            $xml .= '<font>';
            $xml .= '<sz val="'.$font['size'].'"/>';
            if ($font['bold']) {
                $xml .= '<b/>';
            }
            if ($font['color'] !== null) {
                $xml .= '<color rgb="FF'.ltrim($font['color'], '#').'"/>';
            }
            $xml .= '<name val="'.htmlspecialchars($font['name'], ENT_QUOTES | ENT_XML1).'"/>';
            $xml .= '</font>';
        }
        $xml .= '</fonts>';

        // <fills> — first 2 are required built-ins
        $xml .= '<fills count="'.count($this->fills).'">';
        $xml .= '<fill><patternFill patternType="none"/></fill>';
        $xml .= '<fill><patternFill patternType="gray125"/></fill>';
        for ($i = 2; $i < count($this->fills); $i++) {
            $color = ltrim($this->fills[$i] ?? '', '#');
            $xml .= '<fill><patternFill patternType="solid"><fgColor rgb="FF'.$color.'"/></patternFill></fill>';
        }
        $xml .= '</fills>';

        $xml .= '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>';

        // <cellXfs>
        $xml .= '<cellXfs count="'.count($this->cellXfs).'">';
        foreach ($this->cellXfs as $xf) {
            $xml .= '<xf';
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
            $xml .= '/>';
        }
        $xml .= '</cellXfs>';

        $xml .= '</styleSheet>';

        return $xml;
    }

    private function resolveNumFmtId(string $code): int
    {
        if (isset($this->customNumFmts[$code])) {
            return $this->customNumFmts[$code];
        }

        $id = self::CUSTOM_NUMFMT_START + count($this->customNumFmts);
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
                return $i;
            }
        }
        $this->cellXfs[] = $xf;

        return count($this->cellXfs) - 1;
    }

    private function resolveFont(array $font): int
    {
        foreach ($this->fonts as $i => $existing) {
            if ($existing === $font) {
                return $i;
            }
        }
        $this->fonts[] = $font;

        return count($this->fonts) - 1;
    }

    private function resolveFill(string $color): int
    {
        $color = ltrim($color, '#');
        foreach ($this->fills as $i => $existing) {
            if ($i < 2) {
                continue; // skip built-ins
            }
            if (ltrim((string) $existing, '#') === $color) {
                return $i;
            }
        }
        $this->fills[] = $color;

        return count($this->fills) - 1;
    }
}

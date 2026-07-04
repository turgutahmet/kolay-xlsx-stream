<?php

namespace Kolay\XlsxStream\Readers;

/**
 * @internal
 *
 * Parses xl/styles.xml into a styleId → "renders as a date" bitmap for
 * the reader's opt-in autoDetectDates() feature.
 *
 * External writers (Excel, openpyxl, PhpSpreadsheet, Apache POI) store
 * dates as plain t="n" serials whose ONLY date-ness marker is the
 * cell's s="N" style pointing at a date number format. Two sources
 * feed the mapping:
 *
 *   - <numFmts>: custom formats (id >= 164) — formatCode is scanned
 *     for date/time tokens (see formatCodeIsDate)
 *   - built-in ids: 14-22 (date/datetime/time) and 45-47 (elapsed
 *     mm:ss variants) are dates per the OOXML spec's implied table;
 *     other reserved ids (decimals, currency, percent, scientific,
 *     fractions, text) are not
 *
 * The bitmap indexes by cellXfs position — exactly the s="N" values
 * cells carry. <cellStyleXfs> also holds <xf> elements but is a
 * different lookup space (named-style parents); the scan is scoped to
 * the <cellXfs> section so those can never leak in.
 *
 * Same hand-written forward-scan approach as the package's other XML
 * parsers: no DOM, no regex backtracking, linear and DoS-safe.
 */
class StylesParser
{
    /**
     * Built-in numFmtIds that render as dates/times. OOXML reserves
     * 0-163; of those, 14-22 are the date/datetime/time family and
     * 45-47 the elapsed-time family. (18-21 are time-of-day — they
     * carry meaningful fractional serials, so they count.)
     */
    private const BUILTIN_DATE_IDS = [
        14 => true, 15 => true, 16 => true, 17 => true, 18 => true,
        19 => true, 20 => true, 21 => true, 22 => true,
        45 => true, 46 => true, 47 => true,
    ];

    /**
     * Build the cellXfs-indexed date bitmap.
     *
     * @return list<bool>
     */
    public static function dateStyleBitmap(string $stylesXml): array
    {
        $customIsDate = self::collectCustomNumFmts($stylesXml);

        // Scope to the <cellXfs> section — cellStyleXfs/xf entries use
        // the same tag name but are NOT addressable from cells.
        $sectionStart = strpos($stylesXml, '<cellXfs');
        if ($sectionStart === false) {
            return [];
        }
        $sectionEnd = strpos($stylesXml, '</cellXfs>', $sectionStart);
        $section = substr(
            $stylesXml,
            $sectionStart,
            $sectionEnd === false ? strlen($stylesXml) - $sectionStart : $sectionEnd - $sectionStart
        );

        $bitmap = [];
        $cursor = strlen('<cellXfs'); // skip the container's own tag
        $len = strlen($section);

        while ($cursor < $len) {
            $xfStart = strpos($section, '<xf', $cursor);
            if ($xfStart === false) {
                break;
            }
            $next = $section[$xfStart + 3] ?? '';
            if ($next !== ' ' && $next !== '>' && $next !== '/' && $next !== "\t" && $next !== "\n") {
                $cursor = $xfStart + 3;

                continue;
            }

            $tagEnd = self::findTagEnd($section, $xfStart + 3, $len);
            if ($tagEnd === false) {
                break;
            }

            // Missing numFmtId legally defaults to 0 (General) — not a date.
            $numFmtId = (int) (self::attrValue(substr($section, $xfStart + 3, $tagEnd - $xfStart - 3), 'numFmtId') ?? '0');
            $bitmap[] = $customIsDate[$numFmtId] ?? isset(self::BUILTIN_DATE_IDS[$numFmtId]);

            $cursor = $tagEnd + 1;
        }

        return $bitmap;
    }

    /**
     * Decide whether an Excel format code renders its number as a
     * date/time. A code is a date when it contains a d/m/y/h/s token
     * (either case — Excel treats date tokens case-insensitively)
     * OUTSIDE the spans where those letters are literal text:
     *
     *   "..."   quoted literal          — '0.00" m"' is NOT a date
     *   [...]   color/locale/condition  — '[Red]0.00' is NOT a date
     *   \x      escaped character       — '0\m' is NOT a date
     *   _x, *x  width-skip / fill chars — the next char is layout, not
     *                                     a token ('0.00_m')
     *
     * 'E'/'e' is deliberately NOT a token: '0.00E+00' is scientific
     * notation, and no builtin/common date code relies on the rare
     * era-year 'e' without also carrying y/m/d.
     *
     * An unterminated quote or bracket makes everything after it
     * literal-ish garbage — the scan stops there and reports non-date,
     * the conservative reading of a malformed code.
     */
    public static function formatCodeIsDate(string $code): bool
    {
        $len = strlen($code);
        for ($i = 0; $i < $len; $i++) {
            $c = $code[$i];
            if ($c === '"') {
                $close = strpos($code, '"', $i + 1);
                if ($close === false) {
                    return false;
                }
                $i = $close;
            } elseif ($c === '[') {
                $close = strpos($code, ']', $i + 1);
                if ($close === false) {
                    return false;
                }
                $i = $close;
            } elseif ($c === '\\' || $c === '_' || $c === '*') {
                $i++; // next char is a literal / layout char, never a token
            } elseif ($c === 'd' || $c === 'm' || $c === 'y' || $c === 'h' || $c === 's'
                || $c === 'D' || $c === 'M' || $c === 'Y' || $c === 'H' || $c === 'S') {
                return true;
            }
        }

        return false;
    }

    /**
     * numFmtId => is-date for every custom <numFmt> declaration.
     *
     * @return array<int, bool>
     */
    private static function collectCustomNumFmts(string $stylesXml): array
    {
        $out = [];
        $cursor = 0;
        $len = strlen($stylesXml);

        while ($cursor < $len) {
            $start = strpos($stylesXml, '<numFmt', $cursor);
            if ($start === false) {
                break;
            }
            // Refuse the <numFmts> container tag — only the entry matches.
            $next = $stylesXml[$start + 7] ?? '';
            if ($next !== ' ' && $next !== '/' && $next !== "\t" && $next !== "\n") {
                $cursor = $start + 7;

                continue;
            }

            $tagEnd = self::findTagEnd($stylesXml, $start + 7, $len);
            if ($tagEnd === false) {
                break;
            }

            $attrs = substr($stylesXml, $start + 7, $tagEnd - $start - 7);
            $id = self::attrValue($attrs, 'numFmtId');
            $code = self::attrValue($attrs, 'formatCode');
            if ($id !== null && $code !== null) {
                $out[(int) $id] = self::formatCodeIsDate(self::xmlDecode($code));
            }

            $cursor = $tagEnd + 1;
        }

        return $out;
    }

    /**
     * Extract one attribute's raw value from a tag's attribute blob.
     * Exact-name match on the full trimmed span before '=' — same
     * suffix-collision-proof scan as CellTokenizer::extractCellAttrs.
     */
    private static function attrValue(string $attrs, string $wanted): ?string
    {
        $pos = 0;
        $len = strlen($attrs);

        while ($pos < $len) {
            $eq = strpos($attrs, '=', $pos);
            if ($eq === false) {
                return null;
            }
            $nameStart = $pos;
            while ($nameStart < $eq && ($attrs[$nameStart] === ' ' || $attrs[$nameStart] === "\t" || $attrs[$nameStart] === "\n")) {
                $nameStart++;
            }
            $name = rtrim(substr($attrs, $nameStart, $eq - $nameStart));

            $quote = $attrs[$eq + 1] ?? '';
            if ($quote !== '"' && $quote !== "'") {
                return null;
            }
            $valEnd = strpos($attrs, $quote, $eq + 2);
            if ($valEnd === false) {
                return null;
            }
            if ($name === $wanted) {
                return substr($attrs, $eq + 2, $valEnd - $eq - 2);
            }
            $pos = $valEnd + 1;
        }

        return null;
    }

    /**
     * Same quote-aware tag-end walk as the package's other parsers —
     * formatCode attributes legally contain '>' ('0.00&gt;' after
     * decode, conditions inside brackets, etc.).
     */
    private static function findTagEnd(string $s, int $start, int $len)
    {
        $i = $start;
        while ($i < $len) {
            $i += strcspn($s, '>"\'', $i);
            if ($i >= $len) {
                return false;
            }
            $ch = $s[$i];
            if ($ch === '>') {
                return $i;
            }
            $close = strpos($s, $ch, $i + 1);
            if ($close === false) {
                return false;
            }
            $i = $close + 1;
        }

        return false;
    }

    /**
     * Attribute values arrive XML-encoded; '&quot;' in a formatCode is
     * a real double quote that DELIMITS a literal section, so decoding
     * must happen before the token scan or quoted "m"s would count.
     */
    private static function xmlDecode(string $s): string
    {
        if ($s === '' || strpos($s, '&') === false) {
            return $s;
        }

        return html_entity_decode($s, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }
}

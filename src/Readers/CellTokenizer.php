<?php

namespace Kolay\XlsxStream\Readers;

use function count;
// Function imports matter here: unqualified calls from inside a namespace
// pay a per-call namespace-fallback lookup when opcache can't specialize
// them (CLI runs, disabled opcache). In this hot loop that lookup was a
// measurable ~5-8% of tokenization.
use function html_entity_decode;

use Kolay\XlsxStream\Exceptions\XlsxReadException;

use function ord;
use function rtrim;
use function str_starts_with;
use function strcspn;
use function strlen;
use function strpos;
use function substr;
use function substr_compare;

/**
 * @internal
 *
 * Hand-written cell tokenizer for a single <row>...</row> blob.
 *
 * Why hand-written and not regex / expat:
 *   - regex with `.*?` non-greedy patterns has catastrophic backtracking
 *     potential on adversarial input — DoS surface unacceptable on a
 *     network-facing reader.
 *   - expat works but its push-callback model fights with the parser's
 *     overall row-boundary streaming structure. A 100-line state machine
 *     here is faster, simpler, and bounded by construction (every step
 *     is a forward strpos / substr — strictly linear in input size).
 *
 * Cell shapes recognised:
 *   <c r="A1"/>                                       (sparse, empty)
 *   <c r="A1" t="n"><v>123.45</v></c>                 (numeric)
 *   <c r="A1" t="n" s="3"><v>45292</v></c>            (numeric with style)
 *   <c r="A1" t="b"><v>1</v></c>                      (boolean → true)
 *   <c r="A1" t="b"><v>0</v></c>                      (boolean → false)
 *   <c r="A1" t="inlineStr"><is><t>hello</t></is></c> (inline string)
 *   <c r="A1" t="inlineStr"><is><t xml:space="preserve">  ws  </t></is></c>
 *   <c r="A1" t="inlineStr"><is><r><t>foo</t></r><r><t>bar</t></r></is></c>
 *     (rich text — runs concatenated to "foobar"; styling discarded)
 *   <c r="A1" t="s"><v>42</v></c>                     (sst index — caller resolves)
 *   <c r="A1" t="str"><v>cached</v></c>               (formula cached value)
 *   <c r="A1" t="e"><v>#N/A</v></c>                   (error literal)
 *
 * Sparse rows: column position is taken from the r="A1" attribute, not
 * ordinal sibling order — a writer may skip a cell entirely. Gaps are
 * filled with '' so callers always see a dense row up to the rightmost
 * populated column.
 */
class CellTokenizer
{
    /**
     * Parse a row XML blob into a 0-indexed dense array.
     *
     * Pass a SharedStrings lookup to resolve t="s" cells (external XLSX
     * files written with a deduplicated string table). Files written by
     * SinkableXlsxWriter never use t="s", so $sst is optional.
     *
     * Pass a DateDetection bundle (reader's autoDetectDates() opt-in)
     * to convert numeric cells whose s="N" style renders as a date.
     * Detection dispatches to a dedicated loop up front: folding the
     * style checks into this loop cost the default path a consistent
     * ~1% on the tokenizer A/B corpus (one extra null-compare and a
     * wider destructure per CELL), so the two paths share their
     * helpers but not their hot loop — a per-row ternary is the entire
     * price of the feature when it's off.
     *
     * @return array<int, mixed>
     */
    public static function tokenizeRow(string $rowXml, ?SharedStrings $sst = null, ?DateDetection $dates = null): array
    {
        if ($dates !== null) {
            return self::tokenizeRowDetectingDates($rowXml, $sst, $dates);
        }

        $byIdx = [];
        $maxIdx = -1;
        $cursor = 0;
        $len = strlen($rowXml);
        // Writer-shaped rows arrive with cells in strictly increasing
        // column order and no gaps; track that so the dense rebuild at
        // the bottom (one array alloc + N writes per row) can be skipped
        // for the overwhelmingly common case.
        $dense = true;

        while (true) {
            $cellStart = strpos($rowXml, '<c', $cursor);
            if ($cellStart === false) {
                break;
            }

            // The tag must be either "<c " or "<c/>" — guard against bogus
            // matches like "<color"; accept only when followed by space, slash, or '>'.
            $next = $rowXml[$cellStart + 2] ?? '';
            if ($next !== ' ' && $next !== '/' && $next !== '>' && $next !== "\t" && $next !== "\n") {
                $cursor = $cellStart + 2;
                continue;
            }

            // Find end of opening tag (either '/>' for self-closing or '>' for open).
            $tagEnd = self::findTagEnd($rowXml, $cellStart + 2, $len);
            if ($tagEnd === false) {
                break;
            }

            $isSelfClosing = $rowXml[$tagEnd - 1] === '/';
            $attrs = substr(
                $rowXml,
                $cellStart + 2,
                $tagEnd - ($cellStart + 2) - ($isSelfClosing ? 1 : 0)
            );

            // One left-to-right scan pulls r/t together — previously r
            // and t each cost their own strpos walk over the attr blob.
            // (s= is only captured by the date-detection twin of this
            // loop; skipping it here preserves the scan's early exit.)
            [$colLetters, $type] = self::extractCellAttrs($attrs);

            $idx = $colLetters !== null ? self::columnLettersToIndex($colLetters) : ($maxIdx + 1);
            if ($idx !== $maxIdx + 1) {
                $dense = false;
            }
            if ($idx > $maxIdx) {
                $maxIdx = $idx;
            }

            if ($isSelfClosing) {
                $byIdx[$idx] = '';
                $cursor = $tagEnd + 1;
                continue;
            }

            // Locate matching </c>. We're not inside CDATA (writer never
            // emits CDATA) and `<` cannot appear inside a numeric/bool
            // value or in a properly-encoded inline string body, so plain
            // strpos for '</c>' is safe and linear.
            $bodyStart = $tagEnd + 1;
            $bodyEnd = strpos($rowXml, '</c>', $bodyStart);
            if ($bodyEnd === false) {
                break;
            }
            $body = substr($rowXml, $bodyStart, $bodyEnd - $bodyStart);

            $byIdx[$idx] = self::parseCellBody($body, $type, $sst);

            $cursor = $bodyEnd + 4;
        }

        if ($maxIdx < 0) {
            return [];
        }

        // Dense fast path: cells arrived 0,1,2,…,maxIdx with no gaps and
        // no duplicates, so $byIdx already IS the dense row (packed keys
        // in insertion order). Duplicate refs or out-of-order cells clear
        // $dense above and take the gap-filling rebuild below.
        if ($dense && count($byIdx) === $maxIdx + 1) {
            return $byIdx;
        }

        $row = [];
        for ($i = 0; $i <= $maxIdx; $i++) {
            $row[] = $byIdx[$i] ?? '';
        }

        return $row;
    }

    /**
     * tokenizeRow's date-detecting twin — same scan, plus per-cell
     * style capture and conversion. Structural duplication is the
     * point, not an accident: any shared-loop formulation puts at
     * least one detection branch inside the default path's per-cell
     * hot loop, and that measured ~1% on the A/B corpus (see
     * tokenizeRow). Keep the two loops in lockstep when touching
     * either; the shared helpers (findTagEnd / extractCellAttrs /
     * parseCellBody) carry all the actual parsing rules.
     *
     * A cell converts when ALL hold:
     *   - it carries s= and the bitmap marks that style as a date
     *   - its type is numeric (t="n" or no t at all — Excel's own
     *     shape): shared/inline strings, booleans and error literals
     *     keep their values no matter what the style claims
     *   - its column is not in the bundle's skip set (explicit
     *     castColumn precedence, query-path raw columns)
     *
     * @return array<int, mixed>
     */
    private static function tokenizeRowDetectingDates(string $rowXml, ?SharedStrings $sst, DateDetection $dates): array
    {
        $byIdx = [];
        $maxIdx = -1;
        $cursor = 0;
        $len = strlen($rowXml);
        $dense = true;

        while (true) {
            $cellStart = strpos($rowXml, '<c', $cursor);
            if ($cellStart === false) {
                break;
            }

            $next = $rowXml[$cellStart + 2] ?? '';
            if ($next !== ' ' && $next !== '/' && $next !== '>' && $next !== "\t" && $next !== "\n") {
                $cursor = $cellStart + 2;
                continue;
            }

            $tagEnd = self::findTagEnd($rowXml, $cellStart + 2, $len);
            if ($tagEnd === false) {
                break;
            }

            $isSelfClosing = $rowXml[$tagEnd - 1] === '/';
            $attrs = substr(
                $rowXml,
                $cellStart + 2,
                $tagEnd - ($cellStart + 2) - ($isSelfClosing ? 1 : 0)
            );

            [$colLetters, $type, $style] = self::extractCellAttrs($attrs, true);

            $idx = $colLetters !== null ? self::columnLettersToIndex($colLetters) : ($maxIdx + 1);
            if ($idx !== $maxIdx + 1) {
                $dense = false;
            }
            if ($idx > $maxIdx) {
                $maxIdx = $idx;
            }

            if ($isSelfClosing) {
                $byIdx[$idx] = '';
                $cursor = $tagEnd + 1;
                continue;
            }

            $bodyStart = $tagEnd + 1;
            $bodyEnd = strpos($rowXml, '</c>', $bodyStart);
            if ($bodyEnd === false) {
                break;
            }
            $body = substr($rowXml, $bodyStart, $bodyEnd - $bodyStart);

            $value = self::parseCellBody($body, $type, $sst);
            if ($style !== null
                && ($type === null || $type === 'n')
                && ($dates->isDateByStyle[(int) $style] ?? false)
                && ! isset($dates->skipColumns[$idx])
            ) {
                $value = ($dates->convert)($value);
            }
            $byIdx[$idx] = $value;

            $cursor = $bodyEnd + 4;
        }

        if ($maxIdx < 0) {
            return [];
        }

        if ($dense && count($byIdx) === $maxIdx + 1) {
            return $byIdx;
        }

        $row = [];
        for ($i = 0; $i <= $maxIdx; $i++) {
            $row[] = $byIdx[$i] ?? '';
        }

        return $row;
    }

    /**
     * Walk forward from $start looking for the first '>' that closes the
     * tag. Quoted attribute values may contain '>', so respect quotes.
     * Returns the offset of '>' or false.
     *
     * strcspn jumps at C speed to the next interesting byte ('>', '"' or
     * "'"), so the loop iterates once per quoted attribute value instead
     * of once per character — this is the hottest inner loop of the read
     * path and the char-by-char version dominated read CPU profiles.
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
            // Quote — skip straight to its closing partner.
            $close = strpos($s, $ch, $i + 1);
            if ($close === false) {
                return false;
            }
            $i = $close + 1;
        }

        return false;
    }

    /**
     * Single pass over a cell's attribute blob extracting r="" and t=""
     * (and, on request, s="") together. Attribute names are exact match
     * on the full trimmed span before '=', so suffix collisions
     * (xml:space= vs r=) cannot happen by construction — unlike a
     * strpos-for-"r=" scan, which had to guard against them explicitly.
     *
     * $wantStyle is opt-in because writer-shaped cells carry no s= —
     * demanding it would defeat the early exit and re-scan every attr
     * blob to its end (measured ~8% of tokenization). The styles-aware
     * date-detection layer passes true.
     *
     * @return array{0: ?string, 1: ?string, 2: ?string} [column letters, t, s]
     */
    private static function extractCellAttrs(string $attrs, bool $wantStyle = false): array
    {
        $r = null;
        $t = null;
        $s = null;
        $pos = 0;
        $len = strlen($attrs);

        while ($pos < $len) {
            $eq = strpos($attrs, '=', $pos);
            if ($eq === false) {
                break;
            }
            $nameStart = $pos;
            while ($nameStart < $eq && ($attrs[$nameStart] === ' ' || $attrs[$nameStart] === "\t" || $attrs[$nameStart] === "\n")) {
                $nameStart++;
            }
            $name = rtrim(substr($attrs, $nameStart, $eq - $nameStart));

            $quote = $attrs[$eq + 1] ?? '';
            if ($quote !== '"' && $quote !== "'") {
                break;
            }
            $valEnd = strpos($attrs, $quote, $eq + 2);
            if ($valEnd === false) {
                break;
            }
            $val = substr($attrs, $eq + 2, $valEnd - $eq - 2);

            if ($name === 'r') {
                $r = $val;
            } elseif ($name === 't') {
                $t = $val;
            } elseif ($wantStyle && $name === 's') {
                $s = $val;
            }
            $pos = $valEnd + 1;
            if ($r !== null && $t !== null && (! $wantStyle || $s !== null)) {
                break;
            }
        }

        if ($r === null) {
            return [null, $t, $s];
        }

        // Pull the column letters off the front of "AB42".
        $letters = '';
        $rl = strlen($r);
        for ($i = 0; $i < $rl; $i++) {
            $c = $r[$i];
            if ($c >= 'A' && $c <= 'Z') {
                $letters .= $c;

                continue;
            }
            break;
        }

        return [$letters === '' ? null : $letters, $t, $s];
    }

    /**
     * Decode the body of a <c>...</c> element into the value to yield.
     */
    private static function parseCellBody(string $body, ?string $type, ?SharedStrings $sst): mixed
    {
        if ($body === '') {
            return '';
        }

        if ($type === 'inlineStr') {
            // Fast path: the exact single-run shape this package's writer
            // emits — <is><t>…</t></is> with no attributes. One substr
            // instead of the generic <t>-collecting walk. (Length guard:
            // substr_compare with a negative offset throws on strings
            // shorter than the needle.)
            if (strlen($body) >= 16
                && str_starts_with($body, '<is><t>')
                && substr_compare($body, '</t></is>', -9) === 0
            ) {
                return self::xmlDecode(substr($body, 7, -9));
            }

            // Generic path: concatenate every <t>...</t> within the body
            // (handles <t xml:space="preserve"> and rich-text runs
            // <is><r><t>a</t></r><r><t>b</t></r></is>).
            return self::xmlDecode(self::extractAllT($body));
        }

        if ($type === 'b') {
            $v = self::extractInlineV($body);

            return $v === '1';
        }

        if ($type === 's') {
            $v = self::extractInlineV($body);
            // No sst loaded → return the raw index so a corrupt file degrades
            // gracefully rather than throwing mid-iteration. The reader
            // facade is responsible for loading sst when one is present.
            if ($sst === null || $v === '') {
                return $v;
            }

            return $sst->get((int) $v);
        }

        // numeric (t="n" or absent), formula cached (t="str"),
        // error literal (t="e"): all carry the value in <v>...</v>.
        $v = self::extractInlineV($body);

        return self::xmlDecode($v);
    }

    private static function extractInlineV(string $body): string
    {
        $start = strpos($body, '<v>');
        if ($start === false) {
            return '';
        }
        $end = strpos($body, '</v>', $start + 3);
        if ($end === false) {
            return '';
        }

        return substr($body, $start + 3, $end - $start - 3);
    }

    private static function extractAllT(string $body): string
    {
        $out = '';
        $cursor = 0;
        $len = strlen($body);

        while ($cursor < $len) {
            $tStart = strpos($body, '<t', $cursor);
            if ($tStart === false) {
                break;
            }
            // Need either "<t>" or "<t " (opens with attrs).
            $after = $body[$tStart + 2] ?? '';
            if ($after !== '>' && $after !== ' ' && $after !== "\t" && $after !== "\n") {
                $cursor = $tStart + 2;
                continue;
            }

            $tagClose = strpos($body, '>', $tStart);
            if ($tagClose === false) {
                break;
            }

            // Self-closing <t/> means an empty run.
            if ($body[$tagClose - 1] === '/') {
                $cursor = $tagClose + 1;
                continue;
            }

            $tEnd = strpos($body, '</t>', $tagClose + 1);
            if ($tEnd === false) {
                break;
            }

            $out .= substr($body, $tagClose + 1, $tEnd - $tagClose - 1);
            $cursor = $tEnd + 4;
        }

        return $out;
    }

    /**
     * Excel's hard maximum addressable column. The XFD reference
     * resolves to column 16384 (1-indexed) — index 16383 (0-indexed).
     * Cell references beyond this are spec violations and would push
     * the dense row array allocation past any sensible bound.
     */
    public const MAX_COLUMN_INDEX = 16383;

    /**
     * "A" => 0, "Z" => 25, "AA" => 26, "AB" => 27, ...
     *
     * Crafted cell refs like "ZZZZZZ1" resolve to ~12 million columns
     * which would in turn drive parseRow() to allocate a dense array
     * of that size — a memory DoS surface on adversarial input.
     * Indices past Excel's XFD limit are rejected here so the parser
     * never sees them.
     */
    public static function columnLettersToIndex(string $letters): int
    {
        $n = 0;
        $len = strlen($letters);
        for ($i = 0; $i < $len; $i++) {
            $n = $n * 26 + (ord($letters[$i]) - 64);
        }

        $idx = $n - 1;

        if ($idx > self::MAX_COLUMN_INDEX) {
            throw XlsxReadException::corruptCentralDirectory(
                "cell reference '{$letters}' resolves to column {$idx}, ".
                'past Excel\'s XFD maximum (16383)'
            );
        }

        return $idx;
    }

    private static function xmlDecode(string $s): string
    {
        if ($s === '' || strpos($s, '&') === false) {
            return $s;
        }

        return html_entity_decode($s, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }
}

<?php

namespace Kolay\XlsxStream\Readers;

/**
 * @internal
 *
 * Hand-written cell tokenizer for a single <row>...</row> blob.
 *
 * Why hand-written and not regex / expat:
 *   - regex with `.*?` non-greedy patterns has catastrophic backtracking
 *     potential on adversarial input — DoS surface we will not accept on
 *     a network-facing reader (§3.4 of the spec).
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
 * ordinal sibling order — the writer may skip a cell entirely (§17.3 of
 * the spec). Gaps are filled with '' so callers always see a dense row
 * up to the rightmost populated column.
 */
class CellTokenizer
{
    /**
     * Parse a row XML blob into a 0-indexed dense array.
     *
     * @return array<int, mixed>
     */
    public static function tokenizeRow(string $rowXml): array
    {
        $byIdx = [];
        $maxIdx = -1;
        $cursor = 0;
        $len = strlen($rowXml);

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

            $col = self::extractColumnRef($attrs);
            $idx = $col !== null ? self::columnLettersToIndex($col) : ($maxIdx + 1);
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

            $type = self::extractAttribute($attrs, 't');
            $byIdx[$idx] = self::parseCellBody($body, $type);

            $cursor = $bodyEnd + 4;
        }

        if ($maxIdx < 0) {
            return [];
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
     */
    private static function findTagEnd(string $s, int $start, int $len)
    {
        $inQuote = '';
        for ($i = $start; $i < $len; $i++) {
            $ch = $s[$i];
            if ($inQuote !== '') {
                if ($ch === $inQuote) {
                    $inQuote = '';
                }
                continue;
            }
            if ($ch === '"' || $ch === "'") {
                $inQuote = $ch;
                continue;
            }
            if ($ch === '>') {
                return $i;
            }
        }

        return false;
    }

    /**
     * Decode the body of a <c>...</c> element into the value to yield.
     */
    private static function parseCellBody(string $body, ?string $type): mixed
    {
        if ($body === '') {
            return '';
        }

        if ($type === 'inlineStr') {
            // Concatenate every <t>...</t> within the body (handles plain
            // <is><t>x</t></is> and rich-text <is><r><t>a</t></r><r><t>b</t></r></is>).
            return self::xmlDecode(self::extractAllT($body));
        }

        if ($type === 'b') {
            $v = self::extractInlineV($body);

            return $v === '1';
        }

        // numeric (t="n" or absent), sst ref (t="s"), formula cached (t="str"),
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

    private static function extractAttribute(string $attrs, string $name): ?string
    {
        // Match: <space>name="value"  or starts of string.
        // Scan linearly — attributes are short and we never backtrack.
        $needle = $name.'=';
        $pos = 0;
        $len = strlen($attrs);

        while ($pos < $len) {
            $found = strpos($attrs, $needle, $pos);
            if ($found === false) {
                return null;
            }
            // Must be at start of $attrs OR preceded by whitespace; otherwise
            // it's a suffix of another attribute name (e.g. xml:space=).
            $prev = $found > 0 ? $attrs[$found - 1] : ' ';
            if ($prev !== ' ' && $prev !== "\t" && $prev !== "\n" && $found !== 0) {
                $pos = $found + strlen($needle);
                continue;
            }

            $valStart = $found + strlen($needle);
            $quote = $attrs[$valStart] ?? '';
            if ($quote !== '"' && $quote !== "'") {
                return null;
            }
            $valEnd = strpos($attrs, $quote, $valStart + 1);
            if ($valEnd === false) {
                return null;
            }

            return substr($attrs, $valStart + 1, $valEnd - $valStart - 1);
        }

        return null;
    }

    /**
     * Pull the column letters out of an r="A1" / r="AB42" attribute.
     */
    private static function extractColumnRef(string $attrs): ?string
    {
        $r = self::extractAttribute($attrs, 'r');
        if ($r === null) {
            return null;
        }

        $letters = '';
        $len = strlen($r);
        for ($i = 0; $i < $len; $i++) {
            $c = $r[$i];
            if ($c >= 'A' && $c <= 'Z') {
                $letters .= $c;

                continue;
            }
            break;
        }

        return $letters === '' ? null : $letters;
    }

    /**
     * "A" => 0, "Z" => 25, "AA" => 26, "AB" => 27, ...
     */
    public static function columnLettersToIndex(string $letters): int
    {
        $n = 0;
        $len = strlen($letters);
        for ($i = 0; $i < $len; $i++) {
            $n = $n * 26 + (ord($letters[$i]) - 64);
        }

        return $n - 1;
    }

    private static function xmlDecode(string $s): string
    {
        if ($s === '' || strpos($s, '&') === false) {
            return $s;
        }

        return html_entity_decode($s, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }
}

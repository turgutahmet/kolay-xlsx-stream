<?php

namespace Kolay\XlsxStream\Readers;

/**
 * @internal
 *
 * Parses xl/sharedStrings.xml into an in-memory lookup table.
 *
 * Recognised <si> shapes:
 *
 *   <si><t>plain</t></si>
 *   <si><t xml:space="preserve">  whitespace  </t></si>
 *   <si><r><t>foo</t></r><r><t>bar</t></r></si>          rich text — runs concatenated
 *   <si><r><rPr>...</rPr><t>styled</t></r></si>          rich-text styling discarded
 *   <si/>                                                empty entry
 *
 * Hand-written state machine; no DOM, no regex backtracking. Linear in
 * input size and DoS-safe even on adversarial sst payloads.
 */
class SharedStringsParser
{
    public static function parseInMemory(string $sstXml): InMemorySharedStrings
    {
        return new InMemorySharedStrings(self::extractSiBlocks($sstXml));
    }

    /**
     * @return list<string>
     */
    private static function extractSiBlocks(string $xml): array
    {
        $out = [];
        $cursor = 0;
        $len = strlen($xml);

        while ($cursor < $len) {
            $siStart = self::findSiOpen($xml, $cursor, $len);
            if ($siStart < 0) {
                break;
            }

            $afterTag = self::findTagEnd($xml, $siStart + 3, $len);
            if ($afterTag === false) {
                break;
            }

            if ($xml[$afterTag - 1] === '/') {
                $out[] = '';
                $cursor = $afterTag + 1;

                continue;
            }

            $bodyStart = $afterTag + 1;
            $siEnd = strpos($xml, '</si>', $bodyStart);
            if ($siEnd === false) {
                break;
            }

            $body = substr($xml, $bodyStart, $siEnd - $bodyStart);
            $out[] = self::xmlDecode(self::extractAllT($body));
            $cursor = $siEnd + 5;
        }

        return $out;
    }

    /**
     * Find next "<si " or "<si>" or "<si/>" — refuses prefixes belonging
     * to other tag names like <sheetItem>.
     */
    private static function findSiOpen(string $s, int $start, int $len): int
    {
        $pos = $start;
        while ($pos < $len) {
            $found = strpos($s, '<si', $pos);
            if ($found === false) {
                return -1;
            }
            $next = $s[$found + 3] ?? '';
            if ($next === ' ' || $next === '>' || $next === "\t" || $next === "\n" || $next === '/') {
                return $found;
            }
            $pos = $found + 3;
        }

        return -1;
    }

    /**
     * Walk forward from $start until the first '>' that closes the tag.
     * Quoted attribute values may legally contain '>' so we honour quotes.
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
     * Concatenate every <t>...</t> child within $body. Mirrors the same
     * extraction the cell tokenizer uses for inlineStr cells; duplicated
     * here to keep the two parsers loosely coupled.
     */
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
            $after = $body[$tStart + 2] ?? '';
            if ($after !== '>' && $after !== ' ' && $after !== "\t" && $after !== "\n" && $after !== '/') {
                $cursor = $tStart + 2;

                continue;
            }

            $tagClose = strpos($body, '>', $tStart);
            if ($tagClose === false) {
                break;
            }

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

    private static function xmlDecode(string $s): string
    {
        if ($s === '' || strpos($s, '&') === false) {
            return $s;
        }

        return html_entity_decode($s, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }
}

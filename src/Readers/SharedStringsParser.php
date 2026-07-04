<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * @internal
 *
 * Parses xl/sharedStrings.xml into a PackedSharedStrings lookup table.
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
 *
 * The parser is INCREMENTAL: feed inflated chunks through push() as they
 * come off the ZIP entry stream, then call finish() for the table. The
 * full XML document never has to exist in memory — only the packed
 * output plus at most one in-progress <si> entry's carry buffer. Entry
 * boundaries straddling a chunk edge are handled the same way
 * StreamingSheetReader carries an in-progress <row>: whatever bytes
 * belong to an entry whose closing tag hasn't arrived yet stay in the
 * buffer and are re-scanned when the next chunk lands.
 *
 * One instance parses one document. parseInMemory() wraps the
 * push/finish pair for callers (and tests) that already hold the bytes.
 */
class SharedStringsParser
{
    /**
     * Hard ceiling on the in-progress <si> carry buffer. Excel caps
     * cell text at ~32K characters, so even a rich-text entry with
     * hundreds of runs stays under a few MB — an entry that never
     * closes within 16 MB is malformed or malicious, and letting the
     * carry grow unbounded would defeat the streaming parse's whole
     * point. Mirrors StreamingSheetReader::MAX_ROW_XML_BYTES.
     */
    private const MAX_SI_XML_BYTES = 16 * 1024 * 1024;

    /** Unconsumed bytes: at most one in-progress entry + a 3-byte tag tail. */
    private string $buffer = '';

    /** Concatenated decoded strings (PackedSharedStrings payload). */
    private string $payload = '';

    /** pack('V') start offsets, one per completed entry. */
    private string $offsets = '';

    private int $payloadLen = 0;

    private int $count = 0;

    public static function parseInMemory(string $sstXml): PackedSharedStrings
    {
        $parser = new self();
        $parser->push($sstXml);

        return $parser->finish();
    }

    /**
     * Consume one inflated chunk: drain every <si> entry that completes
     * within buffer+chunk into the packed table, carry the rest.
     */
    public function push(string $chunk): void
    {
        $buffer = $this->buffer.$chunk;
        $len = strlen($buffer);
        $cursor = 0;

        while ($cursor < $len) {
            $siStart = self::findSiOpen($buffer, $cursor, $len);
            if ($siStart < 0) {
                // No entry opens in the remaining bytes — they are
                // inter-entry filler (or the document head/tail) and can
                // be dropped, EXCEPT the last 3 bytes: a '<si' split
                // across the chunk edge is invisible to findSiOpen until
                // its discriminating 4th byte arrives, so the tail must
                // survive into the next push.
                $cursor = max($cursor, $len - 3);
                break;
            }

            $afterTag = self::findTagEnd($buffer, $siStart + 3, $len);
            if ($afterTag === false) {
                // Opening tag not closed in this chunk — carry from '<si'.
                $cursor = $siStart;
                break;
            }

            if ($buffer[$afterTag - 1] === '/') {
                $this->append('');
                $cursor = $afterTag + 1;

                continue;
            }

            $bodyStart = $afterTag + 1;
            $siEnd = strpos($buffer, '</si>', $bodyStart);
            if ($siEnd === false) {
                // Entry body still in flight — carry the whole entry.
                $cursor = $siStart;
                break;
            }

            $body = substr($buffer, $bodyStart, $siEnd - $bodyStart);
            $this->append(self::xmlDecode(self::extractAllT($body)));
            $cursor = $siEnd + 5;
        }

        $this->buffer = $cursor > 0 ? substr($buffer, $cursor) : $buffer;

        if (strlen($this->buffer) > self::MAX_SI_XML_BYTES) {
            throw XlsxReadException::corruptCentralDirectory(
                'shared-strings <si> entry exceeds '.(self::MAX_SI_XML_BYTES / 1024 / 1024).
                ' MB without a closing tag — sst is malformed or malicious'
            );
        }
    }

    /**
     * Seal the table. A truncated trailing entry (opened but never
     * closed before EOF) is dropped — same behaviour as the
     * pre-streaming parser, which broke out of its scan on that shape.
     */
    public function finish(): PackedSharedStrings
    {
        // Sentinel end offset: entry i's length is offsets[i+1] -
        // offsets[i], so the table needs count+1 entries.
        $offsets = $this->offsets.pack('V', $this->payloadLen);

        return new PackedSharedStrings($this->payload, $offsets, $this->count);
    }

    private function append(string $s): void
    {
        $this->offsets .= pack('V', $this->payloadLen);
        $this->payload .= $s;
        $this->payloadLen += strlen($s);
        $this->count++;
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
     *
     * strcspn jumps at C speed to the next interesting byte — same
     * optimization as CellTokenizer::findTagEnd (this loop runs once per
     * <si> entry of the shared-strings table, which can be millions).
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

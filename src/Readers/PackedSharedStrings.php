<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * @internal
 *
 * Packed shared-strings lookup: every decoded string lives in ONE
 * concatenated payload buffer, addressed through a binary offset index
 * (pack('V') uint32 little-endian, count+1 entries — the extra sentinel
 * holds the total payload length so entry i's length is always
 * offset[i+1] - offset[i] with no per-entry length column).
 *
 * Why not a list<string> (InMemorySharedStrings)? PHP arrays cost ~57
 * bytes per element on top of the string payloads (measured on a 1M-entry
 * table: 54 MB of zval/bucket overhead for 18 MB of actual text — 3x the
 * data). Two flat strings carry the same table in payload + 4 bytes per
 * entry, and — just as important — they can be APPENDED to while the sst
 * XML streams through the parser, so the full document never has to be
 * materialised next to the table it produces.
 *
 * get() costs two unpack('V') reads + one substr — measured ~2x slower
 * than an array index per lookup, which is noise inside the read path
 * (tokenization dominates by orders of magnitude).
 *
 * Construction is via SharedStringsParser; the constructor trusts its
 * inputs (offsets string length must be ($count + 1) * 4).
 */
class PackedSharedStrings implements SharedStrings
{
    private string $payload;

    private string $offsets;

    private int $count;

    public function __construct(string $payload, string $offsets, int $count)
    {
        $this->payload = $payload;
        $this->offsets = $offsets;
        $this->count = $count;
    }

    public function get(int $index): string
    {
        // Same contract as InMemorySharedStrings: an out-of-range index
        // means the sheet references a string the table never defined —
        // corrupt or truncated sst — and throws rather than degrading to
        // a silent empty cell.
        if ($index < 0 || $index >= $this->count) {
            throw XlsxReadException::corruptCentralDirectory(
                "shared-string index {$index} out of range (table has {$this->count} entries)"
            );
        }

        $pos = $index * 4;
        $start = unpack('V', $this->offsets, $pos)[1];
        $end = unpack('V', $this->offsets, $pos + 4)[1];

        return substr($this->payload, $start, $end - $start);
    }

    public function count(): int
    {
        return $this->count;
    }
}

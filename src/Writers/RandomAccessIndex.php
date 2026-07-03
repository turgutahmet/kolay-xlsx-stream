<?php

namespace Kolay\XlsxStream\Writers;

/**
 * Binary encoder for the xl/_kxs/index.bin sidecar that pairs with a
 * born-indexed XLSX file.
 *
 * Layout (little-endian throughout):
 *
 *     Header (16 bytes)
 *       4   magic           = "KXSI"
 *       1   version         = 2
 *       1   flags           reserved, = 0
 *       2   sheet_count     uint16
 *       4   sync_period     uint32   approx rows between sync points
 *       4   payload_crc32   uint32   CRC32 of every byte after this header
 *
 *     Core body (variable, repeats per sheet)
 *       2   entry_path_len  uint16
 *       N   entry_path      UTF-8, e.g. "xl/worksheets/sheet1.xml"
 *       4   total_rows      uint32   sheet's total row count incl. header
 *       4   sheet_crc32     uint32   ZIP entry CRC32 of the sheet's
 *                                    uncompressed bytes — lets the reader
 *                                    detect a sheet that was edited and
 *                                    re-saved by another tool (Excel,
 *                                    LibreOffice) and silently fall back
 *                                    to a non-indexed scan.
 *       4   sync_count      uint32
 *       24*K sync_points    K = sync_count
 *           each: 8 row uint64, 8 comp_offset uint64, 8 uncomp_offset uint64
 *
 *     TLV sections (optional, after the core body, repeat to payload end)
 *       4   tag             ASCII, e.g. "STAT"
 *       4   length          uint32, payload bytes
 *       N   payload
 *
 *     The version stays 2 even when TLV sections are present — by
 *     construction, not accident. The v2 decoder shipped in v3.0.x
 *     parses exactly sheet_count core sections and ignores anything
 *     after them, while the payload CRC always covered the full body;
 *     so a stats-bearing sidecar is INVISIBLE to old readers rather
 *     than fatal, and they keep their random access. New decoders
 *     detect TLV data simply by the cursor not being at end-of-body
 *     after the core parse, and MUST skip tags they don't recognise —
 *     future features append sections without a version bump that
 *     would strand older readers. (Parquet/ORC lesson: the container
 *     format outlives any single feature decision.)
 *
 *     "STAT" — per-block column statistics (zone maps). Sheets appear in
 *     core-body order; block k spans the rows between sync points k-1
 *     and k, block sync_count is the tail after the last sync point, so
 *     block_count == sync_count + 1 always holds.
 *       per sheet:
 *         2   tracked_column_count  uint16
 *         per column:
 *           2   column        uint16, 1-based
 *           1   flags         bit0 = sorted ascending across the sheet,
 *                             bit1 = sorted descending (numeric values
 *                             only; non-numerics are invisible to the
 *                             ordering check)
 *           4   block_count   uint32
 *           per block (32 bytes):
 *             8   min       float64 ('e') — meaningless when count == 0
 *             8   max       float64
 *             8   sum       float64
 *             4   count     uint32, numeric values folded into min/max/sum
 *             4   other     uint32, nulls + non-numeric values
 *
 * Offsets are relative to the sheet's *compressed* (resp. uncompressed)
 * data stream — measured from just after the Local File Header. The
 * reader resolves them to absolute file offsets at runtime via the
 * Central Directory entry for that sheet.
 *
 * Format is OOXML-neutral: the part is declared in [Content_Types].xml
 * with content type `application/octet-stream` so OPC validators accept
 * it, but no reference points at it from the workbook rels chain.
 * Vanilla XLSX readers (Excel, PhpSpreadsheet, OpenSpout, …) carry it
 * through as opaque binary; only the matched reader recognises the
 * sidecar and gets O(1) random access.
 *
 * Worked size: 1 sheet, 4 M rows, 10 000 sync_period
 *   core  = 16 + (2 + 24 + 4 + 4 + 4) + 24 * 400            ≈  9.66 KB
 *   +STAT = 8 + 2 + (7 + 32 * 401) per tracked column       ≈ +12.8 KB/col
 *
 * @internal
 */
class RandomAccessIndex
{
    public const MAGIC = 'KXSI';
    public const VERSION = 2;
    public const ENTRY_PATH = 'xl/_kxs/index.bin';

    public const TAG_STATS = 'STAT';

    public const SORTED_ASC = 0x01;
    public const SORTED_DESC = 0x02;

    /**
     * @param  int  $syncPeriod  approximate rows between sync points
     * @param  list<array{
     *     entry: string,
     *     total_rows: int,
     *     sheet_crc32: int,
     *     sync_points: list<array{row: int, comp_offset: int, uncomp_offset: int}>
     * }>  $sheets
     * @param  array<string, list<array{
     *     col: int,
     *     sorted_asc: bool,
     *     sorted_desc: bool,
     *     blocks: list<array{min: float, max: float, sum: float, count: int, other: int}>
     * }>>  $columnStats  entry path => tracked columns (order must match $sheets)
     */
    public static function encode(int $syncPeriod, array $sheets, array $columnStats = []): string
    {
        $body = '';
        foreach ($sheets as $sheet) {
            $path = $sheet['entry'];
            $body .= pack('v', strlen($path));
            $body .= $path;
            $body .= pack('V', $sheet['total_rows']);
            $body .= pack('V', $sheet['sheet_crc32']);
            $body .= pack('V', count($sheet['sync_points']));
            foreach ($sheet['sync_points'] as $sp) {
                $body .= pack(
                    'PPP',
                    $sp['row'],
                    $sp['comp_offset'],
                    $sp['uncomp_offset']
                );
            }
        }

        $hasStats = $columnStats !== [];
        if ($hasStats) {
            $stat = '';
            foreach ($sheets as $sheet) {
                $cols = $columnStats[$sheet['entry']] ?? [];
                $stat .= pack('v', count($cols));
                foreach ($cols as $colStat) {
                    $flags = ($colStat['sorted_asc'] ? self::SORTED_ASC : 0)
                        | ($colStat['sorted_desc'] ? self::SORTED_DESC : 0);
                    $stat .= pack('vCV', $colStat['col'], $flags, count($colStat['blocks']));
                    foreach ($colStat['blocks'] as $block) {
                        $stat .= pack(
                            'eeeVV',
                            $block['min'],
                            $block['max'],
                            $block['sum'],
                            $block['count'],
                            $block['other']
                        );
                    }
                }
            }
            $body .= self::TAG_STATS.pack('V', strlen($stat)).$stat;
        }

        $header = self::MAGIC;
        $header .= pack('CCv', self::VERSION, 0, count($sheets));
        $header .= pack('V', $syncPeriod);
        $header .= pack('V', crc32($body));

        return $header.$body;
    }
}

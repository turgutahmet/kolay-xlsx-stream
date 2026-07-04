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
 *     "SCRC" — per-sheet running CRC32 of the sheet's uncompressed bytes
 *     captured at each sync point. Sheets appear in core-body order,
 *     aligned 1:1 with that sheet's sync_points: value k is
 *     crc32(uncompressed sheet bytes [0, sync_point[k].uncomp_offset)),
 *     i.e. the CRC a fresh crc32 would produce after consuming exactly
 *     the prefix that precedes the sync row. count MUST equal the
 *     sheet's sync_count (mirror of the STAT block_count invariant).
 *     Shared prerequisite for resumable exports, appendable files,
 *     truncation detection and block-granular signing: the writer can
 *     resume its live CRC context from a pinned prefix, and a verifier
 *     can prove a prefix untampered without inflating past it.
 *       per sheet:
 *         4   count         uint32, == sync_count of that sheet
 *         4*K running_crc   uint32 each, K = count
 *
 *     "TDIG" / "CHLL" — file-level approximate-statistics sketches per
 *     tracked column: a merging t-digest (quantiles/median) and a
 *     HyperLogLog (distinct counts). Unlike STAT these are whole-sheet
 *     sketches, not per-block pruning structures — one record per sheet
 *     in core-body order, both sections sharing one generic frame:
 *       per sheet:
 *         2   tracked_column_count  uint16 (0 when none for the sheet)
 *         per column:
 *           2   column       uint16, 1-based
 *           4   payload_len  uint32
 *           N   payload      the sketch's own serialized form
 *                            (TDigest::serialize / HyperLogLog::serialize)
 *     Header row is EXCLUDED from both sketches (estimation bias — the
 *     opposite call from STAT's header fold, which exists for pruning
 *     soundness); see SPEC.md §4.3/§4.4 for the full rationale and the
 *     CHLL canonicalization rule.
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

    public const TAG_SYNC_CRCS = 'SCRC';

    public const TAG_TDIGEST = 'TDIG';

    public const TAG_HLL = 'CHLL';

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
     * @param  array<string, list<int>>  $syncPointCrcs  entry path => running CRC32 per
     *     sync point, aligned 1:1 with that sheet's sync_points. Pass [] to omit
     *     the SCRC section entirely; when present every sheet's list length MUST
     *     equal its sync_point count.
     * @param  array<string, array<int, string>>  $columnDigests  entry path =>
     *     1-based column => serialized TDigest payload. Pass [] to omit the
     *     TDIG section entirely.
     * @param  array<string, array<int, string>>  $columnHlls  entry path =>
     *     1-based column => serialized HyperLogLog payload. Pass [] to omit
     *     the CHLL section entirely.
     */
    public static function encode(int $syncPeriod, array $sheets, array $columnStats = [], array $syncPointCrcs = [], array $columnDigests = [], array $columnHlls = []): string
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

        if ($syncPointCrcs !== []) {
            $scrc = '';
            foreach ($sheets as $sheet) {
                $crcs = $syncPointCrcs[$sheet['entry']] ?? [];
                if (count($crcs) !== count($sheet['sync_points'])) {
                    throw new \InvalidArgumentException(
                        "SCRC list for '{$sheet['entry']}' has ".count($crcs).
                        ' entries but the sheet has '.count($sheet['sync_points']).' sync points'
                    );
                }
                $scrc .= pack('V', count($crcs));
                foreach ($crcs as $crc) {
                    $scrc .= pack('V', $crc);
                }
            }
            $body .= self::TAG_SYNC_CRCS.pack('V', strlen($scrc)).$scrc;
        }

        if ($columnDigests !== []) {
            $tdig = self::encodeSketchSection($sheets, $columnDigests);
            $body .= self::TAG_TDIGEST.pack('V', strlen($tdig)).$tdig;
        }

        if ($columnHlls !== []) {
            $chll = self::encodeSketchSection($sheets, $columnHlls);
            $body .= self::TAG_HLL.pack('V', strlen($chll)).$chll;
        }

        $header = self::MAGIC;
        $header .= pack('CCv', self::VERSION, 0, count($sheets));
        $header .= pack('V', $syncPeriod);
        $header .= pack('V', crc32($body));

        return $header.$body;
    }

    /**
     * Shared frame for the TDIG and CHLL sections: one record per sheet
     * in core-body order (positional alignment, like STAT/SCRC — sheets
     * without sketches contribute a count-0 record), each record listing
     * (column, payload_len, payload) triples in ascending column order.
     *
     * @param  list<array{entry: string}>  $sheets  core-body sheet order (extra keys ignored)
     * @param  array<string, array<int, string>>  $byEntry  entry path => 1-based column => sketch payload
     */
    private static function encodeSketchSection(array $sheets, array $byEntry): string
    {
        $section = '';
        foreach ($sheets as $sheet) {
            $cols = $byEntry[$sheet['entry']] ?? [];
            ksort($cols);
            $section .= pack('v', count($cols));
            foreach ($cols as $col => $payload) {
                $section .= pack('vV', $col, strlen($payload)).$payload;
            }
        }

        return $section;
    }
}

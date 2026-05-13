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
 *     Body (variable, repeats per sheet)
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
 *   = 16 + (2 + 24 + 4 + 4 + 4) + 24 * 400 ≈ 9.66 KB
 *
 * @internal
 */
class RandomAccessIndex
{
    public const MAGIC = 'KXSI';
    public const VERSION = 2;
    public const ENTRY_PATH = 'xl/_kxs/index.bin';

    /**
     * @param  int  $syncPeriod  approximate rows between sync points
     * @param  list<array{
     *     entry: string,
     *     total_rows: int,
     *     sheet_crc32: int,
     *     sync_points: list<array{row: int, comp_offset: int, uncomp_offset: int}>
     * }>  $sheets
     */
    public static function encode(int $syncPeriod, array $sheets): string
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

        $header = self::MAGIC;
        $header .= pack('CCv', self::VERSION, 0, count($sheets));
        $header .= pack('V', $syncPeriod);
        $header .= pack('V', crc32($body));

        return $header.$body;
    }
}

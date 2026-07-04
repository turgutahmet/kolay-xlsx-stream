<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * Binary decoder for the xl/_kxs/index.bin sidecar produced by the
 * writer's withRandomAccessIndex() opt-in.
 *
 * Layout pinned in the writer-side encoder. CRC32 of the body is
 * verified on decode; magic/version mismatches and truncated input
 * raise a clear error so a corrupt sidecar surfaces immediately
 * instead of silently degrading the reader's random-access guarantees.
 *
 * Once decoded the instance answers:
 *   - syncPeriod()                  approx rows between sync points
 *   - totalRows($sheetEntry)        sheet's row count, for O(1) rowCount
 *   - syncPoints($sheetEntry)       full per-sheet sync-point list
 *   - findSyncPoint($entry, $row)   nearest sync point with row <= target
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

    private const HEADER_SIZE = 16;

    /**
     * Sane upper bounds enforced during decode. Real workbooks rarely
     * exceed a handful of sheets; treating anything past 1024 as
     * structurally invalid catches malformed sidecars that would
     * otherwise allocate huge intermediate arrays.
     */
    private const MAX_SHEETS = 1024;

    private const MAX_PATH_LEN = 256;

    private int $syncPeriod;

    /** @var array<string, int> entry path => total rows */
    private array $totalRowsByEntry = [];

    /** @var array<string, int> entry path => sheet content CRC32 */
    private array $sheetCrc32ByEntry = [];

    /** @var array<string, list<array{row: int, comp_offset: int, uncomp_offset: int}>> */
    private array $syncPointsByEntry = [];

    /**
     * Per-block column statistics (zone maps) from the KXSI "STAT"
     * TLV section. Block k spans the rows between sync points k-1 and k;
     * the final block is the tail after the last sync point.
     *
     * @var array<string, array<int, array{
     *     sorted_asc: bool,
     *     sorted_desc: bool,
     *     blocks: list<array{min: float, max: float, sum: float, count: int, other: int}>
     * }>> entry path => 1-based column => stats
     */
    private array $columnStatsByEntry = [];

    private function __construct(int $syncPeriod, array $totals, array $crcs, array $syncs, array $columnStats = [])
    {
        $this->syncPeriod = $syncPeriod;
        $this->totalRowsByEntry = $totals;
        $this->sheetCrc32ByEntry = $crcs;
        $this->syncPointsByEntry = $syncs;
        $this->columnStatsByEntry = $columnStats;
    }

    public static function decode(string $payload): self
    {
        // Sync points are encoded as uint64 (`P` format). On 32-bit PHP
        // unpack silently truncates them into garbage offsets that
        // would seek the inflater to wrong file positions and yield
        // arbitrary rows. Loud rejection beats silent corruption.
        if (PHP_INT_SIZE < 8) {
            throw XlsxReadException::corruptCentralDirectory(
                'random-access index requires 64-bit PHP — detected '.PHP_INT_SIZE.'-byte int. '.
                'Upgrade PHP or use a 64-bit build.'
            );
        }

        if (strlen($payload) < self::HEADER_SIZE) {
            throw XlsxReadException::corruptCentralDirectory(
                'random-access index too short to contain a header'
            );
        }

        if (substr($payload, 0, 4) !== self::MAGIC) {
            throw XlsxReadException::corruptCentralDirectory(
                'random-access index magic bytes do not match KXSI'
            );
        }

        $version = ord($payload[4]);
        if ($version !== self::VERSION) {
            throw XlsxReadException::corruptCentralDirectory(
                "random-access index version {$version} is not supported by this reader"
            );
        }

        $sheetCount = unpack('v', substr($payload, 6, 2))[1];
        $syncPeriod = unpack('V', substr($payload, 8, 4))[1];
        $headerCrc = unpack('V', substr($payload, 12, 4))[1];

        $body = substr($payload, self::HEADER_SIZE);
        if (crc32($body) !== $headerCrc) {
            throw XlsxReadException::corruptCentralDirectory(
                'random-access index CRC32 mismatch'
            );
        }

        // Semantic guards. CRC32 only catches accidental bit flips and
        // benign tooling drift — a sidecar can be intentionally crafted
        // with a valid CRC but nonsense content. Cheap structural checks
        // here turn "silently wrong rowAt result" into a loud rejection.
        if ($syncPeriod === 0) {
            throw XlsxReadException::corruptCentralDirectory(
                'random-access index syncPeriod must be greater than zero'
            );
        }
        if ($sheetCount > self::MAX_SHEETS) {
            throw XlsxReadException::corruptCentralDirectory(
                "random-access index sheet count {$sheetCount} exceeds the supported maximum"
            );
        }

        $totals = [];
        $crcs = [];
        $syncs = [];
        $cursor = 0;
        $bodyLen = strlen($body);

        for ($i = 0; $i < $sheetCount; $i++) {
            if ($cursor + 2 > $bodyLen) {
                throw XlsxReadException::corruptCentralDirectory('truncated sheet section header');
            }
            $pathLen = unpack('v', substr($body, $cursor, 2))[1];
            $cursor += 2;

            if ($pathLen === 0 || $pathLen > self::MAX_PATH_LEN) {
                throw XlsxReadException::corruptCentralDirectory(
                    "random-access index sheet path length out of range: {$pathLen}"
                );
            }

            if ($cursor + $pathLen + 12 > $bodyLen) {
                throw XlsxReadException::corruptCentralDirectory('truncated sheet section payload');
            }
            $entry = substr($body, $cursor, $pathLen);
            $cursor += $pathLen;

            if (isset($totals[$entry])) {
                throw XlsxReadException::corruptCentralDirectory(
                    "random-access index lists sheet '{$entry}' more than once"
                );
            }

            $totalRows = unpack('V', substr($body, $cursor, 4))[1];
            $cursor += 4;
            $sheetCrc32 = unpack('V', substr($body, $cursor, 4))[1];
            $cursor += 4;
            $syncCount = unpack('V', substr($body, $cursor, 4))[1];
            $cursor += 4;

            if ($cursor + 24 * $syncCount > $bodyLen) {
                throw XlsxReadException::corruptCentralDirectory(
                    "truncated sync-point block for sheet {$entry}"
                );
            }

            $points = [];
            $prevRow = 0;
            $prevCompOffset = -1;
            $prevUncompOffset = -1;
            for ($k = 0; $k < $syncCount; $k++) {
                $row = unpack('P', substr($body, $cursor, 8))[1];
                $compOffset = unpack('P', substr($body, $cursor + 8, 8))[1];
                $uncompOffset = unpack('P', substr($body, $cursor + 16, 8))[1];

                // Sync points must increase strictly — a rowAt seek
                // resolves the largest sync point at-or-before the
                // target via a forward walk; non-monotonic input
                // would cause the walk to mis-resolve and inflate
                // from the wrong byte position.
                if ($row <= $prevRow) {
                    throw XlsxReadException::corruptCentralDirectory(
                        "random-access index sync points not strictly increasing on row for sheet '{$entry}'"
                    );
                }
                if ($compOffset <= $prevCompOffset) {
                    throw XlsxReadException::corruptCentralDirectory(
                        "random-access index sync points not strictly increasing on comp_offset for sheet '{$entry}'"
                    );
                }
                if ($uncompOffset <= $prevUncompOffset) {
                    throw XlsxReadException::corruptCentralDirectory(
                        "random-access index sync points not strictly increasing on uncomp_offset for sheet '{$entry}'"
                    );
                }
                // Writers capture each sync point at the next-row boundary
                // — i.e. "after seeking to this comp_offset, the next row
                // to inflate is row N". The flush triggered at the very
                // end of the sheet legitimately produces a sync point
                // with row == totalRows + 1 (the would-be position past
                // the last row). Anything beyond that is corruption.
                if ($row > $totalRows + 1) {
                    throw XlsxReadException::corruptCentralDirectory(
                        "random-access index sync row {$row} exceeds totalRows {$totalRows} for sheet '{$entry}'"
                    );
                }

                $points[] = [
                    'row' => $row,
                    'comp_offset' => $compOffset,
                    'uncomp_offset' => $uncompOffset,
                ];
                $prevRow = $row;
                $prevCompOffset = $compOffset;
                $prevUncompOffset = $uncompOffset;
                $cursor += 24;
            }

            $totals[$entry] = $totalRows;
            $crcs[$entry] = $sheetCrc32;
            $syncs[$entry] = $points;
        }

        // Optional TLV sections may follow the core body (the version
        // stays 2 — pre-v3.1 decoders parse exactly sheet_count core
        // sections and ignore trailing bytes, so extensions degrade to
        // invisible instead of fatal for them; see the writer-side
        // encoder docblock). Unknown tags are skipped by design —
        // that's what lets future writers add sections without
        // stranding this reader.
        $columnStats = [];
        $entryOrder = array_keys($totals);
        while ($cursor + 8 <= $bodyLen) {
            $tag = substr($body, $cursor, 4);
            $length = unpack('V', substr($body, $cursor + 4, 4))[1];
            $cursor += 8;

            if ($cursor + $length > $bodyLen) {
                throw XlsxReadException::corruptCentralDirectory(
                    "random-access index TLV section '{$tag}' overruns the payload"
                );
            }

            if ($tag === self::TAG_STATS) {
                $columnStats = self::decodeStatsSection(
                    substr($body, $cursor, $length),
                    $entryOrder,
                    $syncs
                );
            }

            $cursor += $length;
        }

        return new self($syncPeriod, $totals, $crcs, $syncs, $columnStats);
    }

    /**
     * Parse the "STAT" TLV payload — per-block column statistics in
     * core-body sheet order. Enforces the structural invariant
     * block_count == sync_count + 1 so a mismatched sidecar surfaces
     * loudly instead of mis-mapping blocks to row ranges.
     *
     * @param  list<string>  $entryOrder  sheet entries in core-body order
     * @param  array<string, list<array{row: int, comp_offset: int, uncomp_offset: int}>>  $syncs
     * @return array<string, array<int, array{sorted_asc: bool, sorted_desc: bool, blocks: list<array{min: float, max: float, sum: float, count: int, other: int}>}>>
     */
    private static function decodeStatsSection(string $stat, array $entryOrder, array $syncs): array
    {
        $result = [];
        $cursor = 0;
        $len = strlen($stat);

        foreach ($entryOrder as $entry) {
            if ($cursor + 2 > $len) {
                throw XlsxReadException::corruptCentralDirectory('truncated STAT section sheet header');
            }
            $colCount = unpack('v', substr($stat, $cursor, 2))[1];
            $cursor += 2;

            $expectedBlocks = count($syncs[$entry] ?? []) + 1;

            for ($c = 0; $c < $colCount; $c++) {
                if ($cursor + 7 > $len) {
                    throw XlsxReadException::corruptCentralDirectory('truncated STAT column header');
                }
                $col = unpack('v', substr($stat, $cursor, 2))[1];
                $flags = ord($stat[$cursor + 2]);
                $blockCount = unpack('V', substr($stat, $cursor + 3, 4))[1];
                $cursor += 7;

                if ($col < 1) {
                    throw XlsxReadException::corruptCentralDirectory(
                        'STAT section column index must be 1-based'
                    );
                }
                if ($blockCount !== $expectedBlocks) {
                    throw XlsxReadException::corruptCentralDirectory(
                        "STAT section block count {$blockCount} does not match sync points + 1 ({$expectedBlocks}) for sheet '{$entry}'"
                    );
                }
                if ($cursor + 32 * $blockCount > $len) {
                    throw XlsxReadException::corruptCentralDirectory(
                        "truncated STAT block list for sheet '{$entry}' column {$col}"
                    );
                }

                $blocks = [];
                for ($b = 0; $b < $blockCount; $b++) {
                    $vals = unpack('emin/emax/esum/Vcount/Vother', substr($stat, $cursor, 32));
                    $blocks[] = [
                        'min' => $vals['min'],
                        'max' => $vals['max'],
                        'sum' => $vals['sum'],
                        'count' => $vals['count'],
                        'other' => $vals['other'],
                    ];
                    $cursor += 32;
                }

                $result[$entry][$col] = [
                    'sorted_asc' => (bool) ($flags & self::SORTED_ASC),
                    'sorted_desc' => (bool) ($flags & self::SORTED_DESC),
                    'blocks' => $blocks,
                ];
            }
        }

        return $result;
    }

    /**
     * Sheet content CRC32 captured at write time. Reader compares this
     * against the live ZIP CD CRC for the same entry; mismatch means the
     * sheet was rewritten by a tool that didn't update the sidecar
     * (typical Excel/LibreOffice save), so the sidecar is stale and must
     * not be trusted for random access.
     */
    public function sheetCrc32(string $sheetEntry): ?int
    {
        return $this->sheetCrc32ByEntry[$sheetEntry] ?? null;
    }

    public function syncPeriod(): int
    {
        return $this->syncPeriod;
    }

    public function totalRows(string $sheetEntry): ?int
    {
        return $this->totalRowsByEntry[$sheetEntry] ?? null;
    }

    /**
     * @return list<array{row: int, comp_offset: int, uncomp_offset: int}>
     */
    public function syncPoints(string $sheetEntry): array
    {
        return $this->syncPointsByEntry[$sheetEntry] ?? [];
    }

    /**
     * Column statistics for one sheet (KXSI "STAT" section), or null
     * when the sidecar carries none for that entry/column.
     *
     * @return array{sorted_asc: bool, sorted_desc: bool, blocks: list<array{min: float, max: float, sum: float, count: int, other: int}>}|null
     */
    public function columnStats(string $sheetEntry, int $column): ?array
    {
        return $this->columnStatsByEntry[$sheetEntry][$column] ?? null;
    }

    /** @return list<int> 1-based columns that carry stats for the sheet */
    public function statsColumns(string $sheetEntry): array
    {
        return array_keys($this->columnStatsByEntry[$sheetEntry] ?? []);
    }

    /**
     * Row span + seek target for each index block of a sheet, aligned
     * with the STAT section's block order:
     *
     *   block 0            rows 1 .. sync[0].row - 1, streamed from the
     *                      sheet start (no seek — offsets null)
     *   block k (1..K-1)   rows sync[k-1].row .. sync[k].row - 1
     *   block K            rows sync[K-1].row .. totalRows (tail)
     *
     * first_row/last_row are 1-based sheet rows including the header row
     * (data starts at row 2 when the sheet has headers, matching what
     * rows() yields). last_row of the tail block is totalRows.
     *
     * @return list<array{first_row: int, last_row: int, comp_offset: int|null, uncomp_offset: int|null, start_row_at_offset: int|null}>
     */
    public function blockRanges(string $sheetEntry): array
    {
        $points = $this->syncPointsByEntry[$sheetEntry] ?? [];
        $totalRows = $this->totalRowsByEntry[$sheetEntry] ?? 0;

        $ranges = [];
        $firstRow = 1;
        $offset = null;      // block 0 streams from the sheet start
        $uncomp = null;
        $rowAtOffset = null;

        foreach ($points as $sp) {
            $ranges[] = [
                'first_row' => $firstRow,
                'last_row' => $sp['row'] - 1,
                'comp_offset' => $offset,
                'uncomp_offset' => $uncomp,
                'start_row_at_offset' => $rowAtOffset,
            ];
            $firstRow = $sp['row'];
            $offset = $sp['comp_offset'];
            $uncomp = $sp['uncomp_offset'];
            $rowAtOffset = $sp['row'];
        }

        // Tail block after the last sync point (or the whole sheet when
        // no sync points exist).
        $ranges[] = [
            'first_row' => $firstRow,
            'last_row' => $totalRows,
            'comp_offset' => $offset,
            'uncomp_offset' => $uncomp,
            'start_row_at_offset' => $rowAtOffset,
        ];

        return $ranges;
    }

    /**
     * Largest sync point whose row is <= $targetRow. Returns null when
     * the target precedes every recorded sync point (caller should
     * stream from the start of the sheet) or when no points exist.
     *
     * Sync points are stored in row order by the writer, so a forward
     * linear walk suffices — for typical sheets there are at most a
     * few hundred points which is well under any benchmark threshold.
     *
     * @return array{row: int, comp_offset: int, uncomp_offset: int}|null
     */
    public function findSyncPoint(string $sheetEntry, int $targetRow): ?array
    {
        $points = $this->syncPointsByEntry[$sheetEntry] ?? [];
        $best = null;

        foreach ($points as $sp) {
            if ($sp['row'] <= $targetRow) {
                $best = $sp;
            } else {
                break;
            }
        }

        return $best;
    }
}

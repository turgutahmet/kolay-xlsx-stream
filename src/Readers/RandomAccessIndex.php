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

    private function __construct(int $syncPeriod, array $totals, array $crcs, array $syncs)
    {
        $this->syncPeriod = $syncPeriod;
        $this->totalRowsByEntry = $totals;
        $this->sheetCrc32ByEntry = $crcs;
        $this->syncPointsByEntry = $syncs;
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

        return new self($syncPeriod, $totals, $crcs, $syncs);
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

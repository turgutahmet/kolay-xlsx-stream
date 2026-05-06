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
    public const MAGIC = "KXSI";
    public const VERSION = 1;
    public const ENTRY_PATH = 'xl/_kxs/index.bin';

    private const HEADER_SIZE = 16;

    private int $syncPeriod;

    /** @var array<string, int> entry path => total rows */
    private array $totalRowsByEntry = [];

    /** @var array<string, list<array{row: int, comp_offset: int, uncomp_offset: int}>> */
    private array $syncPointsByEntry = [];

    private function __construct(int $syncPeriod, array $totals, array $syncs)
    {
        $this->syncPeriod = $syncPeriod;
        $this->totalRowsByEntry = $totals;
        $this->syncPointsByEntry = $syncs;
    }

    public static function decode(string $payload): self
    {
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

        $totals = [];
        $syncs = [];
        $cursor = 0;
        $bodyLen = strlen($body);

        for ($i = 0; $i < $sheetCount; $i++) {
            if ($cursor + 2 > $bodyLen) {
                throw XlsxReadException::corruptCentralDirectory('truncated sheet section header');
            }
            $pathLen = unpack('v', substr($body, $cursor, 2))[1];
            $cursor += 2;

            if ($cursor + $pathLen + 8 > $bodyLen) {
                throw XlsxReadException::corruptCentralDirectory('truncated sheet section payload');
            }
            $entry = substr($body, $cursor, $pathLen);
            $cursor += $pathLen;

            $totalRows = unpack('V', substr($body, $cursor, 4))[1];
            $cursor += 4;
            $syncCount = unpack('V', substr($body, $cursor, 4))[1];
            $cursor += 4;

            if ($cursor + 24 * $syncCount > $bodyLen) {
                throw XlsxReadException::corruptCentralDirectory(
                    "truncated sync-point block for sheet {$entry}"
                );
            }

            $points = [];
            for ($k = 0; $k < $syncCount; $k++) {
                $points[] = [
                    'row' => unpack('P', substr($body, $cursor, 8))[1],
                    'comp_offset' => unpack('P', substr($body, $cursor + 8, 8))[1],
                    'uncomp_offset' => unpack('P', substr($body, $cursor + 16, 8))[1],
                ];
                $cursor += 24;
            }

            $totals[$entry] = $totalRows;
            $syncs[$entry] = $points;
        }

        return new self($syncPeriod, $totals, $syncs);
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

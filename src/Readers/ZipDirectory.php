<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * @internal
 *
 * Parses a ZIP archive's End of Central Directory (EOCD) record and
 * Central Directory (CD) entries from a Source. The CD lookup gives us
 * compressed/uncompressed sizes, file offsets and CRC32s without ever
 * touching entry bodies — which is what makes range-fetched reading
 * (S3, HTTP) possible.
 *
 * Strategy: pull the trailing 64 KB of the source into a single buffer.
 * In typical XLSX files this contains the EOCD, the entire CD, and
 * often several small entries (workbook.xml, _rels, styles.xml) as
 * well — so subsequent metadata reads come "for free" from the same
 * buffer with no extra round-trip.
 *
 * ZIP64 (>4 GB or >65535 entries) is detected and rejected with a clear
 * error; not yet implemented.
 */
class ZipDirectory
{
    private const EOCD_SIGNATURE = "PK\x05\x06";
    private const CD_ENTRY_SIGNATURE = "PK\x01\x02";
    private const LFH_SIGNATURE = 0x04034b50;
    private const TAIL_PREFETCH = 65557; // 22 EOCD record + max 65535 ZIP comment

    /**
     * @var array<string, array{
     *     compressed_size: int,
     *     uncompressed_size: int,
     *     offset: int,
     *     method: int,
     *     crc32: int,
     * }>
     */
    public array $entries = [];

    /** @var array<string, int> */
    private array $dataOffsetCache = [];

    /**
     * @return array{
     *     compressed_size: int,
     *     uncompressed_size: int,
     *     offset: int,
     *     method: int,
     *     crc32: int,
     * }|null
     */
    public function entry(string $name): ?array
    {
        return $this->entries[$name] ?? null;
    }

    public function has(string $name): bool
    {
        return isset($this->entries[$name]);
    }

    public static function fromSource(Source $source): self
    {
        $size = $source->size();
        $tailLen = (int) min(self::TAIL_PREFETCH, $size);
        $tail = $source->range($size - $tailLen, $tailLen);

        $eocdPos = self::findEocd($tail);
        if ($eocdPos < 0) {
            throw XlsxReadException::eocdNotFound();
        }

        $eocd = unpack(
            'Vsig/vthisDisk/vcdDisk/vthisEntries/vtotalEntries/VcdSize/VcdOffset/vcommentLen',
            substr($tail, $eocdPos, 22)
        );
        if ($eocd === false) {
            throw XlsxReadException::corruptCentralDirectory('cannot unpack EOCD record');
        }

        // ZIP64 sentinel values — bail with a clear error rather than silently misread.
        if (
            $eocd['totalEntries'] === 0xFFFF
            || $eocd['cdSize'] === 0xFFFFFFFF
            || $eocd['cdOffset'] === 0xFFFFFFFF
        ) {
            throw XlsxReadException::zip64NotSupported();
        }

        $cdAbsOffset = (int) $eocd['cdOffset'];
        $cdSize = (int) $eocd['cdSize'];
        $totalEntries = (int) $eocd['totalEntries'];
        $tailFileOffset = $size - $tailLen;

        // CD often falls inside the tail prefetch — saves a round-trip.
        if ($cdAbsOffset >= $tailFileOffset && $cdAbsOffset + $cdSize <= $size) {
            $cdBytes = substr($tail, $cdAbsOffset - $tailFileOffset, $cdSize);
        } else {
            $cdBytes = $source->range($cdAbsOffset, $cdSize);
        }

        return self::parseCentralDirectory($cdBytes, $totalEntries);
    }

    private static function findEocd(string $tail): int
    {
        // EOCD is 22 bytes minimum; comment field can push it earlier.
        // Scan backwards for the signature.
        for ($i = strlen($tail) - 22; $i >= 0; $i--) {
            if (substr($tail, $i, 4) === self::EOCD_SIGNATURE) {
                return $i;
            }
        }

        return -1;
    }

    private static function parseCentralDirectory(string $cdBytes, int $count): self
    {
        $self = new self();
        $cursor = 0;

        for ($k = 0; $k < $count; $k++) {
            if (substr($cdBytes, $cursor, 4) !== self::CD_ENTRY_SIGNATURE) {
                throw XlsxReadException::corruptCentralDirectory(
                    "missing entry signature at index {$k}"
                );
            }

            $h = unpack(
                'Vsig/vverMade/vverNeed/vflags/vmethod/vmtime/vmdate/Vcrc/VcompSize/'.
                'VuncompSize/vfnameLen/vextraLen/vcommentLen/vdiskStart/vintAttr/'.
                'VextAttr/VlocalOffset',
                substr($cdBytes, $cursor, 46)
            );
            if ($h === false) {
                throw XlsxReadException::corruptCentralDirectory(
                    "cannot unpack entry header at index {$k}"
                );
            }

            $name = substr($cdBytes, $cursor + 46, $h['fnameLen']);
            $self->entries[$name] = [
                'compressed_size' => $h['compSize'],
                'uncompressed_size' => $h['uncompSize'],
                'offset' => $h['localOffset'],
                'method' => $h['method'],
                'crc32' => $h['crc'],
            ];

            $cursor += 46 + $h['fnameLen'] + $h['extraLen'] + $h['commentLen'];
        }

        return $self;
    }

    /**
     * Returns the absolute byte offset where the entry's compressed data
     * begins (just past the Local File Header). Result is cached per
     * entry name so repeated random-access reads only pay the 30-byte
     * range fetch once.
     */
    public function dataOffset(Source $source, string $name): int
    {
        if (isset($this->dataOffsetCache[$name])) {
            return $this->dataOffsetCache[$name];
        }

        $entry = $this->entry($name);
        if ($entry === null) {
            throw XlsxReadException::entryNotFound($name);
        }

        $lfhBytes = $source->range($entry['offset'], 30);
        $lfh = unpack(
            'Vsig/vverNeed/vflags/vmethod/vmtime/vmdate/Vcrc/VcompSize/VuncompSize/'.
            'vfnameLen/vextraLen',
            $lfhBytes
        );
        if ($lfh === false || $lfh['sig'] !== self::LFH_SIGNATURE) {
            throw XlsxReadException::badLocalFileHeader($name);
        }

        return $this->dataOffsetCache[$name] =
            $entry['offset'] + 30 + $lfh['fnameLen'] + $lfh['extraLen'];
    }

    /**
     * Read and inflate a single entry's full payload into a string.
     *
     * Intended for small metadata parts — workbook.xml, the .rels files,
     * styles.xml, sharedStrings.xml when chosen for Strategy 1 RAM-load.
     * Do NOT use for sheet data; that is what StreamingSheetReader is for.
     *
     * Supports DEFLATE (method 8) and STORED (method 0); other methods
     * are not part of the OOXML spec and trigger a clear error.
     */
    public function readEntry(Source $source, string $name): string
    {
        $entry = $this->entry($name);
        if ($entry === null) {
            throw XlsxReadException::entryNotFound($name);
        }

        $offset = $this->dataOffset($source, $name);
        $compressed = $source->range($offset, $entry['compressed_size']);

        if ($entry['method'] === 0) {
            return $compressed;
        }

        if ($entry['method'] !== 8) {
            throw XlsxReadException::inflateFailed(
                "entry {$name} uses unsupported ZIP compression method {$entry['method']}"
            );
        }

        $ctx = inflate_init(ZLIB_ENCODING_RAW);
        if ($ctx === false) {
            throw XlsxReadException::inflateFailed("inflate_init failed for {$name}");
        }

        $inflated = inflate_add($ctx, $compressed, ZLIB_FINISH);
        if ($inflated === false) {
            throw XlsxReadException::inflateFailed("inflate_add failed for {$name}");
        }

        return $inflated;
    }
}

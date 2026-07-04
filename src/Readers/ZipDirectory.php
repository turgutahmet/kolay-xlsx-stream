<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Contracts\SupportsSuffixRange;
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
     * Compressed-size ceiling for the coalesced LFH+body fetch in
     * readEntry(). Metadata parts (workbook.xml, .rels, the index
     * sidecar, most shared-strings tables) sit far below this and get
     * their header and body in one round-trip; anything larger — sheet
     * bodies can be hundreds of MB — keeps the lazy two-step flow so
     * the bounded-RAM contract holds.
     */
    private const COALESCE_CAP = 1024 * 1024;

    /**
     * Slack for the LFH extra field in the coalesced fetch. Its true
     * length is only known after parsing the header, so the one-shot
     * read over-fetches by this much; real-world writers emit 0-36
     * bytes of extra data. A larger field just degrades to a second
     * ranged read for the body remainder — never to a wrong result.
     */
    private const LFH_EXTRA_HEADROOM = 64;

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
        // Sources with suffix-range support deliver the tail bytes and
        // the total size in one round-trip (`Range: bytes=-N` on S3);
        // others pay the size() lookup plus a ranged read.
        if ($source instanceof SupportsSuffixRange) {
            ['data' => $tail, 'size' => $size] = $source->tail(self::TAIL_PREFETCH);
            $tailLen = strlen($tail);
        } else {
            $size = $source->size();
            $tailLen = (int) min(self::TAIL_PREFETCH, $size);
            $tail = $source->range($size - $tailLen, $tailLen);
        }

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
        // The last signature whose record still fits (start <= len-22)
        // wins. strrpos with a negative offset caps the match start at
        // exactly that position — same semantics as the previous
        // byte-by-byte reverse scan, without a substr per position.
        if (strlen($tail) < 22) {
            return -1;
        }

        $pos = strrpos($tail, self::EOCD_SIGNATURE, -22);

        return $pos === false ? -1 : $pos;
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
        $lfh = self::parseLfh($lfhBytes, $name);

        return $this->dataOffsetCache[$name] =
            $entry['offset'] + 30 + $lfh['fnameLen'] + $lfh['extraLen'];
    }

    /**
     * Unpack and validate a 30-byte Local File Header. The CD carries
     * its own copy of most fields; the LFH is only consulted for the
     * name/extra lengths that position the entry body.
     *
     * @return array{fnameLen: int, extraLen: int}
     */
    private static function parseLfh(string $lfhBytes, string $name): array
    {
        $lfh = unpack(
            'Vsig/vverNeed/vflags/vmethod/vmtime/vmdate/Vcrc/VcompSize/VuncompSize/'.
            'vfnameLen/vextraLen',
            $lfhBytes
        );
        if ($lfh === false || $lfh['sig'] !== self::LFH_SIGNATURE) {
            throw XlsxReadException::badLocalFileHeader($name);
        }

        return ['fnameLen' => (int) $lfh['fnameLen'], 'extraLen' => (int) $lfh['extraLen']];
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

        $compressed = $this->fetchEntryBody($source, $name, $entry);

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

    /**
     * Stream an entry's INFLATED bytes as string chunks, never holding
     * more than one compressed read + its inflated output at a time.
     *
     * This is readEntry()'s bounded-RAM sibling: same DEFLATE/STORED
     * support, but for parts too large to materialise as one string —
     * today that's xl/sharedStrings.xml, whose parser consumes chunks
     * incrementally. Sheet bodies keep going through
     * StreamingSheetReader, whose chunk engine additionally supports
     * mid-stream offsets (sync-point seeks) that plain entry streaming
     * never needs; the read loop below deliberately mirrors its
     * validation, FINISH-flush and stalled-source handling so the two
     * paths fail identically on the same malformed archives.
     *
     * @return \Generator<int, string>
     */
    public function streamEntry(Source $source, string $name, int $chunkSize = 65536): \Generator
    {
        $entry = $this->entry($name);
        if ($entry === null) {
            throw XlsxReadException::entryNotFound($name);
        }

        $method = $entry['method'];
        if ($method !== 8 && $method !== 0) {
            throw XlsxReadException::inflateFailed(
                "entry {$name} uses unsupported ZIP compression method {$method}"
            );
        }

        $stream = $source->streamFrom($this->dataOffset($source, $name));
        $compRemaining = $entry['compressed_size'];

        $inflate = null;
        if ($method === 8) {
            $inflate = inflate_init(ZLIB_ENCODING_RAW);
            if ($inflate === false) {
                if (is_resource($stream)) {
                    fclose($stream);
                }
                throw XlsxReadException::inflateFailed("inflate_init failed for {$name}");
            }
        }

        // STORED has no deflate state to flush — treat the FINISH step
        // as already done so the loop exits cleanly when input runs out.
        $finishedFlush = $method === 0;
        $emptyReads = 0;

        try {
            while (true) {
                $inflated = '';

                if ($compRemaining > 0) {
                    $compressed = fread($stream, min($chunkSize, $compRemaining));
                    $n = is_string($compressed) ? strlen($compressed) : 0;

                    if ($n === 0) {
                        if (feof($stream)) {
                            $compRemaining = 0;
                        } elseif (++$emptyReads >= 100) {
                            throw XlsxReadException::sourceUnreadable(
                                'source stream stalled — 100 consecutive empty reads with feof=false'
                            );
                        } else {
                            usleep(10_000); // 10ms backoff before retry

                            continue;
                        }
                    } else {
                        $emptyReads = 0;
                        $compRemaining -= $n;

                        if ($method === 8) {
                            $flag = $compRemaining === 0 ? ZLIB_FINISH : ZLIB_NO_FLUSH;
                            if ($flag === ZLIB_FINISH) {
                                $finishedFlush = true;
                            }
                            $inflated = inflate_add($inflate, $compressed, $flag);
                            if ($inflated === false) {
                                throw XlsxReadException::inflateFailed("mid-stream inflate_add failed for {$name}");
                            }
                        } else {
                            // STORED: raw bytes are the "inflated" output.
                            $inflated = $compressed;
                        }
                    }
                } elseif (! $finishedFlush) {
                    $inflated = inflate_add($inflate, '', ZLIB_FINISH);
                    if ($inflated === false) {
                        throw XlsxReadException::inflateFailed("final inflate_add failed for {$name}");
                    }
                    $finishedFlush = true;
                }

                if ($inflated !== '') {
                    yield $inflated;
                }

                if ($compRemaining === 0 && $finishedFlush) {
                    break;
                }
            }
        } finally {
            if (is_resource($stream)) {
                fclose($stream);
            }
        }
    }

    /**
     * Fetch an entry's compressed body, coalescing the Local File
     * Header lookup and the body read into a single ranged fetch for
     * small entries.
     *
     * The two-step flow (30-byte LFH read to learn fnameLen/extraLen,
     * then a body read) costs two round-trips per entry — noticeable on
     * S3 where each is a full HTTP request. For entries at or below
     * COALESCE_CAP a single over-fetch covers everything:
     *
     *     [ LFH 30 ][ name ][ extra ≤ headroom ][ body compressed_size ]
     *
     * The real extra length is parsed out of the fetched LFH bytes; if
     * a writer used an extra field larger than the headroom, only the
     * missing body tail is fetched separately — a rare shape that
     * degrades to the old two-request cost, never to a wrong result.
     *
     * @param  array{compressed_size: int, offset: int}  $entry
     */
    private function fetchEntryBody(Source $source, string $name, array $entry): string
    {
        // Offset already known (nothing to coalesce) or body too large
        // to prefetch — plain two-step read.
        if (isset($this->dataOffsetCache[$name]) || $entry['compressed_size'] > self::COALESCE_CAP) {
            $offset = $this->dataOffset($source, $name);

            return $source->range($offset, $entry['compressed_size']);
        }

        $want = 30 + strlen($name) + self::LFH_EXTRA_HEADROOM + $entry['compressed_size'];
        $buf = $source->range($entry['offset'], $want);

        $lfh = self::parseLfh(substr($buf, 0, 30), $name);
        $dataStart = 30 + $lfh['fnameLen'] + $lfh['extraLen'];

        // Populate the shared cache so later dataOffset()/readEntry()
        // calls for this entry skip the LFH fetch entirely.
        $this->dataOffsetCache[$name] = $entry['offset'] + $dataStart;

        if ($dataStart + $entry['compressed_size'] <= strlen($buf)) {
            return substr($buf, $dataStart, $entry['compressed_size']);
        }

        // Extra field exceeded the headroom: keep what the over-fetch
        // already delivered and pull only the remainder.
        $have = substr($buf, $dataStart);
        $missing = $entry['compressed_size'] - strlen($have);

        return $have.$source->range($entry['offset'] + $dataStart + strlen($have), $missing);
    }
}

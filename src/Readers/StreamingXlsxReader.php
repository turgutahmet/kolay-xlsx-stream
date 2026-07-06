<?php

namespace Kolay\XlsxStream\Readers;

use Aws\S3\S3Client;
use Kolay\XlsxStream\Contracts\ProvidesCostHints;
use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Sketches\HyperLogLog;
use Kolay\XlsxStream\Sketches\TDigest;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Sources\S3RangeSource;

/**
 * Public-facing streaming reader.
 *
 * Construction is via factory methods, never `new`:
 *
 *     StreamingXlsxReader::fromFile('/path/to/big.xlsx')
 *     StreamingXlsxReader::fromS3($s3Client, 'bucket', 'key.xlsx')
 *     StreamingXlsxReader::from(new LocalFileSource(...))
 *
 * Once constructed the reader exposes:
 *
 *   - sheets()                   list of {name, sheetId, entry}
 *   - onSheet('Reports')         select sheet by name
 *   - onSheetIndex(0)            select sheet by tab order
 *   - header()                   first row of the selected sheet
 *   - rows(skip, limit)          generator over data rows
 *   - chunked(N)                 generator over batches of N rows
 *   - rowCount()                 total rows in selected sheet (boundary scan)
 *
 * RAM is bounded — independent of sheet size. Each call to rows() /
 * chunked() opens a fresh forward-only inflate stream so the API is
 * replayable from the caller's perspective without memoising the
 * dataset in memory.
 *
 * Files written by SinkableXlsxWriter use inline strings exclusively
 * (t="inlineStr") — no sharedStrings.xml is needed. Files written by
 * other XLSX writers (PhpSpreadsheet, openpyxl, Apache POI, …) typically
 * deduplicate strings into xl/sharedStrings.xml; the reader detects this
 * and loads the table transparently — streamed straight from the inflate
 * loop into a packed lookup (one payload buffer + a binary offset index),
 * so the table costs its text plus 4 bytes per entry, never a PHP array
 * or the whole XML document. Archives whose sharedStrings.xml exceeds
 * SST_RAM_THRESHOLD compressed bytes are refused with a clear error — an
 * on-disk variant for very large tables is tracked as a future addition.
 */
class StreamingXlsxReader
{
    /**
     * Compressed-size threshold at which xl/sharedStrings.xml stops
     * fitting comfortably in RAM. ~99% of real-world XLSX files have a
     * sst well below this. Files above the threshold get a clear error
     * pointing at the limitation.
     *
     * 64 MB (up from 20 MB pre-v3.2): the packed table + streaming
     * parse cut the measured resolution peak ~3.5x — a 30 MB sst that
     * used to cost 84 MB of RAM (full XML + list<string> at ~57 B/entry
     * of array overhead) now costs 24 MB (payload + 4 B/entry offset
     * index, XML never materialised) — so the same RAM budget covers a
     * proportionally larger table.
     */
    public const SST_RAM_THRESHOLD = 64 * 1024 * 1024;

    /**
     * Upper bound on the *uncompressed* shared-strings size we will
     * parse into RAM. Highly repetitive XML (a single repeated <si>
     * entry, common in adversarial inputs and accidentally in some
     * exports) compresses 50:1 or higher, so a compressed payload well
     * under SST_RAM_THRESHOLD can balloon far past it. This second
     * guard keeps the bounded-RAM contract intact even when the
     * deflate ratio is pathological.
     *
     * 320 MB (up from 100 MB pre-v3.2), scaled by the same measured
     * ~3.5x reduction: the resolution peak is now the packed table —
     * roughly the string payload (typically 50-70% of the XML size)
     * plus 4 bytes per entry — not a multiple of the document, so a
     * table this size parses in ~200 MB (measured 245 MB XML / 8M
     * entries → 185 MB peak) instead of the ~700 MB the old model
     * would have needed.
     */
    public const SST_UNCOMPRESSED_THRESHOLD = 320 * 1024 * 1024;

    private Source $source;
    private ZipDirectory $cd;

    /** Memoized gap-bridge byte budget (rtt × bandwidth); see maxBridgeBytes(). */
    private ?int $bridgeBytesCache = null;

    /** Observability hook fired when a query degrades to a full scan; see onFullScan(). */
    private ?\Closure $onFullScan = null;

    /** @var list<array{name: string, sheetId: int, entry: string}> */
    private array $sheets;

    private string $currentEntry;
    private ?array $cachedHeader = null;
    private ?SharedStrings $sst = null;
    private bool $sstResolved = false;
    private ?RandomAccessIndex $randomAccessIndex = null;
    private bool $indexResolved = false;

    /**
     * Per-entry memo of RandomAccessIndex::blockRanges(). The index
     * rebuilds the ranges from its sync-point list on every call, and
     * query methods (rowsWhere / groupStats / findRow) may run hundreds
     * of times against the same sheet — hot query workloads pay that
     * rebuild per call for an answer that never changes. The decoded
     * index is immutable and resolved once per reader, so entries can
     * never go stale; keying by entry also makes sheet switches
     * invalidation-free (each entry carries its own slot).
     *
     * @var array<string, list<array{first_row: int, last_row: int, comp_offset: int|null, uncomp_offset: int|null, start_row_at_offset: int|null}>>
     */
    private array $blockRangesByEntry = [];

    /**
     * Physical row count (header included) of a sheet the writer filled
     * to Excel's split threshold before rolling to a continuation sheet
     * (BaseXlsxWriter::ROWS_PER_SHEET). A non-final sheet of an
     * auto-split chain has EXACTLY this many rows by construction —
     * that near-impossible-by-hand precision is what makes fullness a
     * reliable continuation signal.
     */
    private const FULL_SHEET_ROWS = 1_048_575;

    /**
     * Continuation-chain map: entry => ordered members of the chain the
     * entry belongs to. A chain is a maximal run of consecutive sheets
     * where every non-final member is exactly full and every member
     * carries an identical header row — i.e. ONE logical table the
     * writer split only because of Excel's row ceiling. Populated
     * lazily by resolveChains(); single-sheet workbooks and files
     * without a usable sidecar never build it.
     *
     * Member shape: entry, total (physical rows incl. header),
     * dataStartLocal (1 for the first member, 2 for continuations —
     * their repeated header row does not exist logically), globalStart
     * (global row number of local row dataStartLocal).
     *
     * @var array<string, list<array{entry: string, total: int, dataStartLocal: int, globalStart: int}>>|null
     */
    private ?array $chainByEntry = null;

    /**
     * Raw (cast-free) header rows read for chain detection, memoized so
     * each candidate sheet pays its one small read at most once.
     *
     * @var array<string, array<int, mixed>>
     */
    private array $headerByEntry = [];

    /**
     * Header-name resolution cache per sheet entry: 'map' (name =>
     * 0-based index, unique names only), 'dupes' (name => every 0-based
     * position, for the ambiguity error), 'headers' (addressable names
     * in sheet order, for the unknown-name error). Built lazily by
     * resolveColumnName(); entry-keyed, so sheet switches need no
     * invalidation.
     *
     * @var array<string, array{map: array<string, int>, dupes: array<string, list<int>>, headers: list<string>}>
     */
    private array $columnNamesByEntry = [];

    /** @var array<string, TDigest|null> merged chain digests, keyed chain-head|column */
    private array $chainDigestCache = [];

    /** @var array<string, HyperLogLog|null> merged chain HLLs, keyed chain-head|column */
    private array $chainHllCache = [];

    /** @var array<int, callable> */
    private array $columnCasts = [];

    private bool $use1904Epoch = false;

    private bool $autoDetectDates = false;

    private bool $autoDetectWithTime = true;

    /**
     * cellXfs index → is-date bitmap from xl/styles.xml, resolved
     * lazily like the sst (styles are workbook-wide, so no per-sheet
     * invalidation). null when the archive has no styles.xml.
     *
     * @var list<bool>|null
     */
    private ?array $dateStyleBitmap = null;

    private bool $stylesResolved = false;

    /**
     * Excel stores dates as timezone-naive numeric serials. We materialise
     * them as UTC by default for portability — the same file produces the
     * same DateTimeImmutable on every server regardless of
     * date_default_timezone_get(). Callers whose data is authored in a
     * specific timezone opt in via castTimezone().
     */
    private string $castTimezone = 'UTC';

    /**
     * Cached DateTimeZone matching $castTimezone. Constructing a
     * DateTimeZone does a tz-database lookup — paying that per date
     * cell dominates castDate's cost on date-heavy sheets, so the
     * instance is built once and reused. castTimezone() swaps it when
     * the setting changes.
     */
    private ?\DateTimeZone $castTimezoneObj = null;

    private function __construct(Source $source)
    {
        $this->source = $source;
        $this->cd = ZipDirectory::fromSource($source);

        // workbook.xml feeds two construction steps — sheet resolution
        // and date-epoch detection. Fetch + inflate the part once and
        // hand the bytes to both; on S3 sources this saves a redundant
        // round-trip per open. When the part is missing, resolve()
        // raises the canonical error.
        $workbookXml = $this->cd->has('xl/workbook.xml')
            ? $this->cd->readEntry($source, 'xl/workbook.xml')
            : null;

        $this->sheets = WorkbookResolver::resolve($source, $this->cd, $workbookXml);

        if ($this->sheets === []) {
            throw XlsxReadException::corruptCentralDirectory('workbook contains no sheets');
        }

        // Honour the workbook's declared date epoch. Mac-origin XLSX
        // files set workbookPr/@date1904 — without this, every cast
        // date comes back four years and a day too early. Manual
        // use1904Epoch() override still works on top of this default.
        $this->use1904Epoch = $workbookXml !== null
            && WorkbookResolver::parseDate1904Xml($workbookXml);

        $this->currentEntry = $this->sheets[0]['entry'];
    }

    public static function from(Source $source): self
    {
        return new self($source);
    }

    public static function fromFile(string $path): self
    {
        return new self(new LocalFileSource($path));
    }

    public static function fromS3(S3Client $s3, string $bucket, string $key): self
    {
        return new self(new S3RangeSource($s3, $bucket, $key));
    }

    /**
     * Sheets in workbook.xml order (matches the user-visible tab order).
     *
     * @return list<array{name: string, sheetId: int, entry: string}>
     */
    public function sheets(): array
    {
        return $this->sheets;
    }

    public function onSheet(string $name): self
    {
        foreach ($this->sheets as $sheet) {
            if ($sheet['name'] === $name) {
                $this->currentEntry = $sheet['entry'];
                $this->cachedHeader = null;
                $this->columnCasts = [];

                return $this;
            }
        }

        throw XlsxReadException::entryNotFound("sheet named '{$name}'");
    }

    public function onSheetIndex(int $index): self
    {
        if (! isset($this->sheets[$index])) {
            throw XlsxReadException::entryNotFound("sheet at index {$index}");
        }
        $this->currentEntry = $this->sheets[$index]['entry'];
        $this->cachedHeader = null;
        $this->columnCasts = [];

        return $this;
    }

    /**
     * Return the first row of the selected sheet. Cached after first call,
     * so subsequent invocations do not re-open the underlying stream.
     *
     * @return array<int, mixed>
     */
    public function header(): array
    {
        if ($this->cachedHeader !== null) {
            return $this->cachedHeader;
        }

        foreach ($this->openSheetReader()->rows() as $row) {
            return $this->cachedHeader = $row;
        }

        return $this->cachedHeader = [];
    }

    /**
     * Yield rows from the selected sheet. By default returns every row
     * including the header; pass skip=1 to drop the header, or skip=N
     * to start at row N+1. Yielded keys count emitted rows from 0 —
     * they are NOT sheet row numbers (rowRange() carries those).
     *
     * skip is cheap: the reader seeks to the nearest sync point before
     * the target when the file carries a usable index, and covers the
     * remaining gap (or, without an index, the whole prefix) with a
     * '</row>' boundary scan instead of tokenizing rows nobody asked
     * for. Yielded rows are byte-for-byte what the naive "drop the
     * first N yields" loop produced before v3.2 — only the skipping
     * got faster (measured: 3.9s -> 2.5ms for skip=1M indexed, 1.6s ->
     * 56ms for skip=400K without an index).
     *
     * @return \Generator<int, array<int, mixed>>
     */
    public function rows(int $skip = 0, ?int $limit = null): \Generator
    {
        if ($skip < 0) {
            $skip = 0;
        }
        if ($limit !== null && $limit <= 0) {
            return;
        }

        // Chain: iterate the logical table across every member via the
        // chain-aware rowRange (casts applied there), re-keying to this
        // generator's documented 0-based emitted count.
        $chain = $this->chain();
        if ($chain !== null) {
            $globalEnd = $this->chainTotalRows($chain);
            if ($skip + 1 > $globalEnd) {
                return;
            }
            $emitted = 0;
            foreach ($this->rowRange($skip + 1, $globalEnd) as $row) {
                yield $row;
                $emitted++;
                if ($limit !== null && $emitted >= $limit) {
                    return;
                }
            }

            return;
        }

        if ($skip > 0) {
            // rows() counts yields, and rowsFromOffset() yields exactly
            // one row per '</row>' boundary from its starting row — so
            // "drop the first N yields" is the same operation as "start
            // at row N+1" (the countRows() equivalence argument; self-
            // closing rows produce neither a yield nor a boundary).
            // seekTarget degrades to [null, 1] without a usable index,
            // in which case $fastForwardTo alone does the work as a
            // boundary scan from the sheet start.
            [$compOffset, $startingRow] = $this->seekTarget($skip + 1);

            $emitted = 0;
            foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startingRow, $skip + 1) as $row) {
                yield $this->applyCasts($row);
                $emitted++;
                if ($limit !== null && $emitted >= $limit) {
                    return;
                }
            }

            return;
        }

        $emitted = 0;

        foreach ($this->openSheetReader()->rows() as $row) {
            yield $this->applyCasts($row);
            $emitted++;
            if ($limit !== null && $emitted >= $limit) {
                return;
            }
        }
    }

    /**
     * Yield rows in fixed-size batches. Last batch may be shorter than
     * $size if the sheet length is not divisible. Designed for bulk DB
     * inserts where the caller wants amortised round-trip cost.
     *
     * @return \Generator<int, list<array<int, mixed>>>
     */
    public function chunked(int $size, int $skip = 0): \Generator
    {
        if ($size < 1) {
            throw new \InvalidArgumentException('Chunk size must be at least 1');
        }

        $batch = [];
        foreach ($this->rows($skip) as $row) {
            $batch[] = $row;
            if (count($batch) >= $size) {
                yield $batch;
                $batch = [];
            }
        }
        if ($batch !== []) {
            yield $batch;
        }
    }

    /**
     * Total row count including header. O(1) when the file carries a
     * matching xl/_kxs/index.bin sidecar (born-indexed); O(N) inflate
     * scan otherwise — but a boundary-counting scan that skips cell
     * tokenization entirely, not a full parse. Both call sites are
     * covered by the same tests so the result is identical either way.
     */
    public function rowCount(): int
    {
        // Auto-split continuation chain: the file is ONE logical table,
        // so the count spans every member (continuation headers do not
        // exist logically). Pre-v3.2.2 this silently returned the
        // active sheet's count alone.
        $chain = $this->chain();
        if ($chain !== null) {
            return $this->chainTotalRows($chain);
        }

        $index = $this->loadRandomAccessIndex();
        if ($index !== null) {
            $total = $index->totalRows($this->currentEntry);
            if ($total !== null) {
                return $total;
            }
        }

        // No-index fallback: count '</row>' boundaries in the inflated
        // stream — same result as iterator_count over rows() (see
        // StreamingSheetReader::countRows for the equivalence argument)
        // at a fraction of the CPU. The shared-strings table plays no
        // part in counting, so it is deliberately not resolved here; on
        // S3 sources that also skips the sst fetch round-trip.
        return (new StreamingSheetReader($this->source, $this->cd, $this->currentEntry))->countRows();
    }

    /**
     * Verify every sheet's data against the checksums the writer recorded,
     * a single inflate pass per sheet at O(1) memory. On born-indexed files
     * this checks the per-block running CRC (SCRC) pinned at each sync point
     * AND the whole-sheet CRC, so corruption is localized: `corrupt_blocks`
     * lists the 0-based sync-point index of each block whose bytes no longer
     * match — the "which block went bad" report for audit/payroll/HR data.
     * Without a sidecar it falls back to the whole-entry ZIP CRC only (no
     * block granularity). `inflate_ok=false` marks a sheet whose compressed
     * bytes were too damaged to inflate at all.
     *
     * @return array{ok: bool, sheets: list<array{sheet: int, entry: string, ok: bool, blocks_checked: int, corrupt_blocks: list<int>, sheet_crc_ok: bool, inflate_ok: bool}>}
     */
    public function verify(): array
    {
        $index = $this->loadRandomAccessIndex();
        $sheets = [];
        $allOk = true;

        foreach ($this->sheets as $i => $sheet) {
            $entry = $sheet['entry'];
            $indexed = $index !== null && $index->sheetCrc32($entry) !== null;

            if ($indexed) {
                $sheetCrc = $index->sheetCrc32($entry);
                $syncPoints = $index->syncPoints($entry);
                $syncCrcs = $index->syncPointCrcs($entry);
                $checkpoints = [];
                foreach ($syncPoints as $k => $sp) {
                    if (isset($syncCrcs[$k], $sp['uncomp_offset'])) {
                        $checkpoints[] = [$sp['uncomp_offset'], $syncCrcs[$k]];
                    }
                }
            } else {
                // No SCRC pins — verify the whole entry against the ZIP's own CRC.
                $sheetCrc = $this->cd->entry($entry)['crc32'] ?? -1;
                $checkpoints = [];
            }

            $res = (new StreamingSheetReader($this->source, $this->cd, $entry))
                ->verifyCrc($checkpoints, $sheetCrc);

            $sheetOk = $res['inflate_ok'] && $res['sheet_crc_ok'] && $res['corrupt_blocks'] === [];
            $allOk = $allOk && $sheetOk;

            $sheets[] = [
                'sheet' => $i,
                'entry' => $entry,
                'ok' => $sheetOk,
                'blocks_checked' => count($checkpoints),
                'corrupt_blocks' => $res['corrupt_blocks'],
                'sheet_crc_ok' => $res['sheet_crc_ok'],
                'inflate_ok' => $res['inflate_ok'],
            ];
        }

        return ['ok' => $allOk, 'sheets' => $sheets];
    }

    /**
     * Return a single row by 1-based row number — row 1 is the header,
     * row 2 is the first data row. Returns null when $rowNumber is past
     * the end of the sheet.
     *
     * Cost:
     *   - With xl/_kxs/index.bin → O(period) — fresh inflate from the
     *     nearest sync point, boundary-scanning (not tokenizing) the
     *     up-to-$period rows between the sync point and the target.
     *     Effectively constant time bounded by the writer-chosen sync
     *     period; the fast-forward cut that cost ~19x (measured 21ms →
     *     1.1ms per random rowAt at period 10K).
     *   - Without the sidecar → O(N) — inflate from the start, but the
     *     prefix is boundary-scanned rather than tokenized.
     *
     * @return array<int, mixed>|null
     */
    public function rowAt(int $rowNumber): ?array
    {
        if ($rowNumber < 1) {
            return null;
        }

        // Chain: $rowNumber is a GLOBAL row — map it to (member, local)
        // and read from that member's entry directly. State (casts,
        // current sheet) is never touched: chain members share one
        // schema by construction.
        $chain = $this->chain();
        if ($chain !== null) {
            $loc = $this->chainLocate($chain, $rowNumber);
            if ($loc === null) {
                return null;
            }
            [$memberIdx, $local] = $loc;
            $entry = $chain[$memberIdx]['entry'];
            [$compOffset, $startingRow] = $this->seekTarget($local, $entry);

            foreach ($this->openSheetReader([], $entry)->rowsFromOffset($compOffset, $startingRow, $local) as $rn => $row) {
                if ($rn === $local) {
                    return $this->applyCasts($row);
                }
                if ($rn > $local) {
                    return null;
                }
            }

            return null;
        }

        [$compOffset, $startingRow] = $this->seekTarget($rowNumber);

        foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startingRow, $rowNumber) as $rn => $row) {
            if ($rn === $rowNumber) {
                return $this->applyCasts($row);
            }
            if ($rn > $rowNumber) {
                return null;
            }
        }

        return null;
    }

    /**
     * Yield rows by 1-based inclusive range [from, to]. With an index
     * sidecar the cost is O(period + (to - from)); without, the cost
     * matches a row-skip O(from) plus the size of the range. Yielded
     * keys are the 1-based row numbers.
     *
     * @return \Generator<int, array<int, mixed>>
     */
    public function rowRange(int $from, int $to): \Generator
    {
        if ($from < 1) {
            $from = 1;
        }
        if ($to < $from) {
            return;
        }

        // Chain: [from, to] are GLOBAL rows — walk the overlapping
        // members in order, translating each member's local rows back
        // to global keys. Continuation headers (local row 1) are below
        // every member's dataStartLocal and never surface.
        $chain = $this->chain();
        if ($chain !== null) {
            $globalEnd = $this->chainTotalRows($chain);
            if ($from > $globalEnd) {
                return;
            }
            $to = min($to, $globalEnd);

            $loc = $this->chainLocate($chain, $from);
            if ($loc === null) {
                return;
            }
            [$memberIdx, $localFrom] = $loc;

            for ($i = $memberIdx, $n = count($chain); $i <= $n - 1; $i++) {
                $m = $chain[$i];
                $coverEnd = $m['globalStart'] + $m['total'] - $m['dataStartLocal'];
                if ($m['globalStart'] > $to && $i > $memberIdx) {
                    return;
                }
                $startLocal = $i === $memberIdx ? $localFrom : $m['dataStartLocal'];
                $stopLocal = $m['dataStartLocal'] + (min($to, $coverEnd) - $m['globalStart']);

                [$compOffset, $startingRow] = $this->seekTarget($startLocal, $m['entry']);
                foreach ($this->openSheetReader([], $m['entry'])->rowsFromOffset($compOffset, $startingRow, $startLocal) as $rn => $row) {
                    if ($rn < $startLocal) {
                        continue;
                    }
                    if ($rn > $stopLocal) {
                        break;
                    }
                    yield $m['globalStart'] + ($rn - $m['dataStartLocal']) => $this->applyCasts($row);
                }

                if ($coverEnd >= $to) {
                    return;
                }
            }

            return;
        }

        [$compOffset, $startingRow] = $this->seekTarget($from);

        // $from doubles as the fast-forward target: rows between the
        // sync point and the range start are boundary-scanned, never
        // tokenized. The $rn < $from guard stays as a no-op safety net.
        foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startingRow, $from) as $rn => $row) {
            if ($rn < $from) {
                continue;
            }
            if ($rn > $to) {
                return;
            }
            yield $rn => $this->applyCasts($row);
        }
    }

    /**
     * Whole-sheet aggregate for a stats-tracked column, answered from
     * the KXSI sidecar alone — no row data is read. On S3 that means
     * a multi-GB file's column total costs the sidecar fetch and nothing
     * else.
     *
     * $column is 1-based, matching the writer's withColumnStats() —
     * this API pairs with that opt-in (unlike castColumn(), whose
     * 0-based indexing addresses the yielded row arrays). A header
     * NAME works everywhere a column number does — 'Amount' instead
     * of 4 — and makes that base split invisible; naming semantics
     * live on resolveColumnName(). Every query API below (quantile,
     * median, countDistinct, rowsWhere, findRow, groupStats) accepts
     * the same int|string form.
     *
     * Returns null when the file carries no stats for the column (not
     * born-indexed, stale sidecar, or column not tracked) — callers
     * decide whether to fall back to a scan. min/max/avg are null when
     * the column held no numeric values at all. Counts include the
     * header row's cell (as `other` for the usual text header): the
     * header participates in stats so zone-map pruning stays consistent
     * with the full-scan path.
     *
     * @return array{min: float|null, max: float|null, sum: float, avg: float|null, count: int, other: int, sorted: string|null}|null
     */
    public function columnStats(int|string $column): ?array
    {
        if (\is_string($column)) {
            $column = $this->resolveColumnName($column) + 1;
        }
        $index = $this->loadRandomAccessIndex();

        // Chain: fold every member's blocks — the sidecar carries each
        // physical sheet's stats separately and sums/min/max compose
        // exactly. All-or-nothing: if any member lacks the column, the
        // chain answer would silently cover part of the table (the very
        // bug this path fixes), so return null instead. `other` keeps
        // its documented per-physical-sheet header contribution — one
        // header cell per member for the usual text header.
        $chain = $this->chain();
        if ($chain !== null && $index !== null) {
            $min = null;
            $max = null;
            $sum = 0.0;
            $count = 0;
            $other = 0;
            $allAsc = true;
            $allDesc = true;
            $prevMemberMin = null;
            $prevMemberMax = null;
            $boundariesAsc = true;
            $boundariesDesc = true;

            foreach ($chain as $m) {
                $stats = $index->columnStats($m['entry'], $column);
                if ($stats === null) {
                    return null;
                }

                $mMin = null;
                $mMax = null;
                foreach ($stats['blocks'] as $block) {
                    if ($block['count'] > 0) {
                        $mMin = $mMin === null ? $block['min'] : min($mMin, $block['min']);
                        $mMax = $mMax === null ? $block['max'] : max($mMax, $block['max']);
                        $sum += $block['sum'];
                        $count += $block['count'];
                    }
                    $other += $block['other'];
                }

                $allAsc = $allAsc && $stats['sorted_asc'];
                $allDesc = $allDesc && $stats['sorted_desc'];

                // Chain-level sortedness needs the member boundaries in
                // order too. Member min/max include the header folded
                // into block 0, which can only turn a true verdict into
                // null (conservative-safe), never the reverse.
                if ($mMin !== null) {
                    if ($prevMemberMax !== null && $mMin < $prevMemberMax) {
                        $boundariesAsc = false;
                    }
                    if ($prevMemberMin !== null && $mMax > $prevMemberMin) {
                        $boundariesDesc = false;
                    }
                    $prevMemberMin = $mMin;
                    $prevMemberMax = $mMax;
                    $min = $min === null ? $mMin : min($min, $mMin);
                    $max = $max === null ? $mMax : max($max, $mMax);
                }
            }

            return [
                'min' => $min,
                'max' => $max,
                'sum' => $sum,
                'avg' => $count > 0 ? $sum / $count : null,
                'count' => $count,
                'other' => $other,
                'sorted' => ($allAsc && $boundariesAsc) ? 'asc' : (($allDesc && $boundariesDesc) ? 'desc' : null),
            ];
        }

        $stats = $index?->columnStats($this->currentEntry, $column);
        if ($stats === null) {
            return null;
        }

        $min = null;
        $max = null;
        $sum = 0.0;
        $count = 0;
        $other = 0;
        foreach ($stats['blocks'] as $block) {
            if ($block['count'] > 0) {
                $min = $min === null ? $block['min'] : min($min, $block['min']);
                $max = $max === null ? $block['max'] : max($max, $block['max']);
                $sum += $block['sum'];
                $count += $block['count'];
            }
            $other += $block['other'];
        }

        return [
            'min' => $min,
            'max' => $max,
            'sum' => $sum,
            'avg' => $count > 0 ? $sum / $count : null,
            'count' => $count,
            'other' => $other,
            'sorted' => $stats['sorted_asc'] ? 'asc' : ($stats['sorted_desc'] ? 'desc' : null),
        ];
    }

    /**
     * Approximate value at quantile $q (0 = min .. 1 = max) of a
     * column's numeric values, answered from the sidecar's t-digest
     * sketch (KXSI "TDIG") alone — ZERO row data is read and, the index
     * being cached at open, zero additional range requests are issued:
     * on S3 a multi-GB file answers its p99 from bytes already in hand.
     *
     * The value population is the STAT one — cells the writer rendered
     * numerically (int/float, numeric strings, DateTime as Excel serial,
     * bool as 0/1) in DATA rows; the header row is excluded by the
     * format (sketches estimate the data distribution — see SPEC.md
     * §4.3). q=0 and q=1 are the exact min/max; typical rank error at
     * compression 100 is well under 0.1% at the tails.
     *
     * Returns null when the file carries no sketch for the column (not
     * born-sketched, stale sidecar, or column not tracked) — callers
     * decide whether an exact scan is worth it. Also null for a sketch
     * that saw no numeric values (e.g. a text column).
     */
    public function quantile(int|string $column, float $q): ?float
    {
        if (\is_string($column)) {
            $column = $this->resolveColumnName($column) + 1;
        }
        if ($q < 0.0 || $q > 1.0) {
            throw new \InvalidArgumentException("quantile must be within [0, 1]; got {$q}");
        }

        $chain = $this->chain();
        if ($chain !== null) {
            return $this->chainDigest($chain, $column)?->quantile($q);
        }

        $digest = $this->loadRandomAccessIndex()?->columnDigest($this->currentEntry, $column);

        return $digest?->quantile($q);
    }

    /**
     * Merged t-digest across a chain's members — the sketches' merge
     * associativity is exactly what makes per-sheet sections compose
     * into one logical-table answer. All-or-nothing like columnStats:
     * a member without the sketch would make the merged answer cover
     * part of the table, so the result is null instead. Memoized per
     * (chain, column); the first member's digest is cloned so the
     * sidecar's cached instances are never mutated.
     *
     * @param  list<array{entry: string, total: int, dataStartLocal: int, globalStart: int}>  $chain
     */
    private function chainDigest(array $chain, int $column): ?TDigest
    {
        $key = $chain[0]['entry'].'|'.$column;
        if (array_key_exists($key, $this->chainDigestCache)) {
            return $this->chainDigestCache[$key];
        }

        $index = $this->loadRandomAccessIndex();
        $merged = null;
        foreach ($chain as $m) {
            $digest = $index?->columnDigest($m['entry'], $column);
            if ($digest === null) {
                return $this->chainDigestCache[$key] = null;
            }
            if ($merged === null) {
                $merged = clone $digest;
            } else {
                $merged->merge($digest);
            }
        }

        return $this->chainDigestCache[$key] = $merged;
    }

    /**
     * Approximate median — sugar for quantile($column, 0.5).
     */
    public function median(int|string $column): ?float
    {
        return $this->quantile($column, 0.5);
    }

    /**
     * Approximate number of distinct values in a column, answered from
     * the sidecar's HyperLogLog sketch (KXSI "CHLL") alone — like
     * quantile(), zero row data and zero additional requests.
     *
     * Distinctness is over canonical string forms (SPEC.md §4.4), so
     * text columns are fully covered — non-numeric values, invisible to
     * quantile(), count here. Empty cells (null/'') are not values and
     * are never counted; the header row is excluded. Standard error at
     * p=11 is ±~2.3%; small cardinalities are near-exact (linear
     * counting).
     *
     * Returns null when the file carries no sketch for the column;
     * 0 for a tracked column whose data rows were all empty.
     */
    public function countDistinct(int|string $column): ?int
    {
        if (\is_string($column)) {
            $column = $this->resolveColumnName($column) + 1;
        }
        $chain = $this->chain();
        if ($chain !== null) {
            return $this->chainHll($chain, $column)?->count();
        }

        $hll = $this->loadRandomAccessIndex()?->columnHll($this->currentEntry, $column);

        return $hll?->count();
    }

    /**
     * Merged HyperLogLog across a chain's members — register-wise max,
     * the union sketch. Same all-or-nothing and clone-first rules as
     * chainDigest(); a precision mismatch between members (impossible
     * from one writer run, conceivable from a crafted sidecar) answers
     * null rather than throwing mid-query.
     *
     * @param  list<array{entry: string, total: int, dataStartLocal: int, globalStart: int}>  $chain
     */
    private function chainHll(array $chain, int $column): ?HyperLogLog
    {
        $key = $chain[0]['entry'].'|'.$column;
        if (array_key_exists($key, $this->chainHllCache)) {
            return $this->chainHllCache[$key];
        }

        $index = $this->loadRandomAccessIndex();
        $merged = null;
        foreach ($chain as $m) {
            $hll = $index?->columnHll($m['entry'], $column);
            if ($hll === null) {
                return $this->chainHllCache[$key] = null;
            }
            if ($merged === null) {
                $merged = clone $hll;
            } else {
                try {
                    $merged->merge($hll);
                } catch (\InvalidArgumentException) {
                    return $this->chainHllCache[$key] = null;
                }
            }
        }

        return $this->chainHllCache[$key] = $merged;
    }

    /**
     * Register a hook fired whenever a query CANNOT push down and falls
     * back to a full row scan — an unindexed column on rowsWhere /
     * rowsWhereAll, or a groupBy column that is not sorted / not tracked
     * on groupStats. The callback receives a context array
     * `{query, column, reason, entry}` so callers can log or alert when a
     * query silently reads the whole sheet instead of the sidecar. Pass
     * null to clear. On an auto-split chain it may fire once per member.
     *
     * @param  (callable(array{query: string, column: int|null, reason: string, entry: string}): void)|null  $callback
     */
    public function onFullScan(?callable $callback): self
    {
        $this->onFullScan = $callback === null ? null : \Closure::fromCallable($callback);

        return $this;
    }

    private function fireFullScan(string $query, ?int $column, string $reason, string $entry): void
    {
        if ($this->onFullScan !== null) {
            ($this->onFullScan)([
                'query' => $query,
                'column' => $column,
                'reason' => $reason,
                'entry' => $entry,
            ]);
        }
    }

    /**
     * Yield rows whose $column (1-based, see columnStats) satisfies a
     * numeric predicate, using the sidecar's per-block zone maps to skip
     * every block whose [min, max] provably contains no match — the
     * spreadsheet equivalent of Parquet row-group pruning. Exports are
     * usually ID/date-sorted, so range queries typically inflate a
     * handful of blocks out of hundreds.
     *
     * Ops: '=', '<', '<=', '>', '>=', 'between' ($value2 = upper bound,
     * inclusive). The predicate matches numeric cells only (the raw
     * tokenized value, before any castColumn transforms); text cells
     * never match. Yields 1-based row number => row (casts applied).
     *
     * Degrades to a full-scan filter when the sidecar carries no stats
     * for the column — same results, no pruning.
     *
     * @return \Generator<int, array<int, mixed>>
     */
    public function rowsWhere(int|string $column, string $op, int|float $value, int|float|null $value2 = null): \Generator
    {
        if (\is_string($column)) {
            $column = $this->resolveColumnName($column) + 1;
        }

        // Chain: run the per-sheet matcher over every member in order,
        // remapping local row numbers to the chain's global numbering.
        // Continuation members' repeated header row (local 1) is
        // filtered — it does not exist logically, and its local number
        // would collide with the previous member's last data row after
        // remapping.
        $chain = $this->chain();
        if ($chain !== null) {
            foreach ($chain as $i => $m) {
                foreach ($this->matchRowsIn($m['entry'], $column, $op, $value, $value2) as $rn => $row) {
                    if ($i > 0 && $rn === 1) {
                        continue;
                    }
                    yield $m['globalStart'] + ($rn - $m['dataStartLocal']) => $this->applyCasts($row);
                }
            }

            return;
        }

        foreach ($this->matchRowsIn($this->currentEntry, $column, $op, $value, $value2) as $rn => $row) {
            yield $rn => $this->applyCasts($row);
        }
    }

    /**
     * Yield rows satisfying EVERY predicate (a conjunction / AND), pruned
     * by intersecting each predicate's surviving-block set. Because zone
     * maps are sound, a row matching all predicates lives in a block that
     * survives all of them, so the intersection can only shrink the
     * candidate blocks — a query filtering on two differently-clustered
     * columns reads far fewer blocks than either predicate alone.
     *
     * Each predicate is `[column, op, value, value2?]`: column is a
     * 1-based index or a header name (resolved like rowsWhere), op is one
     * of '=', '<', '<=', '>', '>=', 'between' ($value2 = inclusive upper
     * bound). Predicates on columns without zone maps prune nothing but
     * still filter per row. With no statful predicate at all it degrades
     * to a full-scan filter. Yields 1-based row number => row (casts
     * applied), chain-aware exactly like rowsWhere.
     *
     * @param  list<array{0: int|string, 1: string, 2: int|float, 3?: int|float|null}>  $predicates
     * @return \Generator<int, array<int, mixed>>
     */
    public function rowsWhereAll(array $predicates): \Generator
    {
        if ($predicates === []) {
            throw new \InvalidArgumentException('rowsWhereAll() requires at least one predicate');
        }

        $preds = [];
        foreach ($predicates as $p) {
            $col = $p[0];
            if (\is_string($col)) {
                $col = $this->resolveColumnName($col) + 1;
            }
            $preds[] = ['col' => $col, 'op' => $p[1], 'value' => $p[2], 'value2' => $p[3] ?? null];
        }

        $chain = $this->chain();
        if ($chain !== null) {
            foreach ($chain as $i => $m) {
                foreach ($this->matchRowsAllIn($m['entry'], $preds) as $rn => $row) {
                    if ($i > 0 && $rn === 1) {
                        continue;
                    }
                    yield $m['globalStart'] + ($rn - $m['dataStartLocal']) => $this->applyCasts($row);
                }
            }

            return;
        }

        foreach ($this->matchRowsAllIn($this->currentEntry, $preds) as $rn => $row) {
            yield $rn => $this->applyCasts($row);
        }
    }

    /**
     * Estimate how many rows a single predicate selects — WITHOUT reading
     * a row (answered entirely from the sidecar). Returns:
     *   - upper: the summed value-count of the zone-map surviving blocks,
     *     a HARD upper bound (never below the true match count);
     *   - estimate: an approximate count from the t-digest (range ops:
     *     count × ΔCDF, inverted from the in-memory digest by bisection)
     *     or the HLL for '=' (count ÷ distinct), or null when the column
     *     carries no sketch.
     *
     * Returns null when the column has no zone maps (no basis to bound).
     * Chain-aware: bounds fold across continuation members (all-or-nothing).
     *
     * @return array{upper: int, estimate: int|null}|null
     */
    public function estimatedRows(int|string $column, string $op, int|float $value, int|float|null $value2 = null): ?array
    {
        if (\is_string($column)) {
            $column = $this->resolveColumnName($column) + 1;
        }

        $upper = $this->survivingRowCount($column, $op, $value, $value2);
        if ($upper === null) {
            return null;
        }

        return ['upper' => $upper, 'estimate' => $this->estimateSelectivity($column, $op, $value, $value2)];
    }

    /**
     * Describe the query plan rowsWhereAll() would run for a predicate
     * set — WITHOUT executing it. Zero I/O; every number comes from the
     * sidecar. Shape:
     *   - strategy: 'zone-map-prune' (some predicate pruned blocks) or
     *     'full-scan' (no statful predicate / no index);
     *   - candidateBlocks: blocks surviving the intersection;
     *   - runs: contiguous block runs to scan (each = one seek);
     *   - estimatedRows: {upper: int|null, estimate: int|null} — upper is
     *     the row-count bound of the candidate blocks; estimate assumes
     *     predicate independence;
     *   - estimatedBytes: compressed bytes the scan would fetch (the S3
     *     range budget — exactly what bounded streamFrom will request).
     *
     * @param  list<array{0: int|string, 1: string, 2: int|float, 3?: int|float|null}>  $predicates
     * @return array{strategy: string, candidateBlocks: int, runs: int, estimatedRows: array{upper: int|null, estimate: int|null}, estimatedBytes: int}
     */
    public function explain(array $predicates): array
    {
        if ($predicates === []) {
            throw new \InvalidArgumentException('explain() requires at least one predicate');
        }

        $preds = [];
        foreach ($predicates as $p) {
            $col = $p[0];
            if (\is_string($col)) {
                $col = $this->resolveColumnName($col) + 1;
            }
            $preds[] = ['col' => $col, 'op' => $p[1], 'value' => $p[2], 'value2' => $p[3] ?? null];
        }

        $index = $this->loadRandomAccessIndex();
        $chain = $this->chain();
        $entries = $chain !== null ? array_map(fn ($m) => $m['entry'], $chain) : [$this->currentEntry];

        $pruned = false;
        $candidateBlocks = 0;
        $runs = 0;
        $bytes = 0;
        $upper = 0;
        $upperKnown = false;
        foreach ($entries as $entry) {
            $plan = $this->planAllIn($index, $entry, $preds);
            if (! $plan['full']) {
                $pruned = true;
            }
            $candidateBlocks += $plan['candidateBlocks'];
            $runs += $plan['runs'];
            $bytes += $plan['bytes'];
            if ($plan['upper'] !== null) {
                $upper += $plan['upper'];
                $upperKnown = true;
            }
        }

        return [
            'strategy' => $pruned ? 'zone-map-prune' : 'full-scan',
            'candidateBlocks' => $candidateBlocks,
            'runs' => $runs,
            'estimatedRows' => [
                'upper' => $upperKnown ? $upper : null,
                'estimate' => $this->estimateAnd($preds),
            ],
            'estimatedBytes' => $bytes,
        ];
    }

    /**
     * "ORDER BY $column [DESC] LIMIT $k" answered from the sidecar:
     * return the $k rows with the smallest ($desc = false) or largest
     * ($desc = true) values in $column, ranked by that value.
     *
     * On a column the writer flagged sorted, the extremes are the rows at
     * one end of the sheet, so only the blocks at that end are read (one
     * seek, early exit) — an indexed top-N. On an unsorted column it is a
     * single full scan holding an O($k) heap (no pruning is possible, but
     * memory stays bounded). Ranking uses the raw stored numeric value —
     * the same value columnStats/rowsWhere see — while the returned rows
     * carry any registered casts. Non-numeric cells never rank.
     *
     * Returns an ordered map of 1-based row number => row (rank order:
     * position 0 is the top). Empty for $k <= 0 or a header-only sheet;
     * $k is clamped to the available data rows.
     *
     * @return array<int, array<int, mixed>>
     */
    public function topRows(int|string $column, int $k, bool $desc = false): array
    {
        if (\is_string($column)) {
            $column = $this->resolveColumnName($column) + 1;
        }
        if ($k <= 0) {
            return [];
        }

        $total = $this->rowCount();
        $dataRows = $total - 1; // row 1 is the header
        if ($dataRows <= 0) {
            return [];
        }
        $k = min($k, $dataRows);

        $stats = $this->columnStats($column);
        $sorted = $stats['sorted'] ?? null;

        if ($sorted === 'asc' || $sorted === 'desc') {
            return $this->topRowsSorted($k, $desc, $sorted, $total);
        }

        return $this->topRowsScan($column, $k, $desc);
    }

    /**
     * A uniform random sample of $k rows, fetched via the index — so a
     * sample of a multi-GB sheet costs ≈$k block reads, not a full scan.
     * $k distinct row numbers are drawn uniformly (without replacement)
     * from the data range with a seeded Mt19937 Randomizer, then exactly
     * those rows are read; each data row has an equal $k/N chance
     * (unbiasedness proven in poc/d3_sample.php via chi-square). The
     * Randomizer is local — the global mt_rand() state is untouched.
     *
     * Pass $seed for a reproducible sample (same file + seed => same
     * rows); omit it for a fresh secure-random draw. Returns an ascending
     * map of 1-based row number => row (casts applied). Empty for
     * $k <= 0; the whole table (in order) when $k >= the row count.
     * Best for $k far below the total — at large $k a full rows() scan is
     * cheaper than that many seeks.
     *
     * @return array<int, array<int, mixed>>
     */
    public function sampleRows(int $k, ?int $seed = null): array
    {
        if ($k <= 0) {
            return [];
        }

        $total = $this->rowCount();
        if ($total <= 1) {
            return []; // header only, no data rows
        }

        $lo = 2;            // row 1 is the header
        $hi = $total;
        $n = $hi - $lo + 1;
        $k = min($k, $n);

        $rows = [];
        if ($k === $n) {
            foreach ($this->rowRange($lo, $hi) as $rn => $row) {
                $rows[$rn] = $row;
            }

            return $rows;
        }

        $rng = new \Random\Randomizer(
            $seed !== null ? new \Random\Engine\Mt19937($seed) : null
        );

        // Rejection sampling for k DISTINCT rows — getInt() is itself
        // rejection-based, so no modulo bias; the set keys dedup draws.
        // Efficient while k ≪ n, which is the intended use.
        $picked = [];
        while (count($picked) < $k) {
            $picked[$rng->getInt($lo, $hi)] = true;
        }
        ksort($picked);

        foreach (array_keys($picked) as $rn) {
            $row = $this->rowAt($rn);
            if ($row !== null) {
                $rows[$rn] = $row;
            }
        }

        return $rows;
    }

    /**
     * Sorted-column fast path: the k extremes are a contiguous block of
     * rows at one end of the sheet. Read just that row range (via the
     * seeking rowRange) and order it by value — no per-row comparison
     * needed, the sorted flag guarantees value order equals row order.
     *
     * @return array<int, array<int, mixed>>
     */
    private function topRowsSorted(int $k, bool $desc, string $sorted, int $total): array
    {
        // Which end holds the wanted extreme? On an asc column high values
        // sit at the tail; on desc, at the head. We want the high end when
        // $desc (largest) is requested.
        $wantHigh = $desc;
        $highAtTail = $sorted === 'asc';
        $readTail = $wantHigh === $highAtTail;

        [$from, $to] = $readTail ? [$total - $k + 1, $total] : [2, $k + 1];

        $rows = [];
        foreach ($this->rowRange($from, $to) as $rn => $row) {
            $rows[$rn] = $row;
        }

        // $rows is keyed by ascending row number. Value ascends with row
        // number iff the column is asc-sorted. Reorder so position 0 is
        // the requested top (high first when $desc).
        $ascendingByValue = $sorted === 'asc';
        if ($ascendingByValue === $desc) {
            $rows = array_reverse($rows, true);
        }

        return $rows;
    }

    /**
     * Unsorted-column path: a single full scan keeping the running top-k
     * by raw value in an O(k) heap-like array. Casts are applied only to
     * the surviving rows.
     *
     * @return array<int, array<int, mixed>>
     */
    private function topRowsScan(int $column, int $k, bool $desc): array
    {
        $idx = $column - 1;
        $top = [];          // rn => ['v' => float, 'row' => raw row]
        $threshold = null;  // worst value currently kept (evict candidate)

        foreach ($this->scanAllRaw() as $rn => $row) {
            $cell = $row[$idx] ?? null;
            if (! is_numeric($cell)) {
                continue;
            }
            $v = (float) $cell;

            if (count($top) < $k) {
                $top[$rn] = ['v' => $v, 'row' => $row];
                if (count($top) === $k) {
                    $threshold = $this->worstValue($top, $desc);
                }

                continue;
            }

            // Only touch the heap when the new value beats the weakest one.
            if ($desc ? $v > $threshold : $v < $threshold) {
                unset($top[$this->worstRow($top, $desc)]);
                $top[$rn] = ['v' => $v, 'row' => $row];
                $threshold = $this->worstValue($top, $desc);
            }
        }

        uasort($top, fn ($a, $b) => $desc ? ($b['v'] <=> $a['v']) : ($a['v'] <=> $b['v']));

        $out = [];
        foreach ($top as $rn => $e) {
            $out[$rn] = $this->applyCasts($e['row']);
        }

        return $out;
    }

    /**
     * Weakest value in the current top-k (the min when ranking largest,
     * the max when ranking smallest) — the eviction threshold.
     *
     * @param  array<int, array{v: float, row: array<int, mixed>}>  $top
     */
    private function worstValue(array $top, bool $desc): float
    {
        $worst = null;
        foreach ($top as $e) {
            if ($worst === null || ($desc ? $e['v'] < $worst : $e['v'] > $worst)) {
                $worst = $e['v'];
            }
        }

        return $worst ?? 0.0;
    }

    /**
     * Row number of the weakest entry in the current top-k.
     *
     * @param  array<int, array{v: float, row: array<int, mixed>}>  $top
     */
    private function worstRow(array $top, bool $desc): int
    {
        $worstRn = array_key_first($top);
        $worstV = null;
        foreach ($top as $rn => $e) {
            if ($worstV === null || ($desc ? $e['v'] < $worstV : $e['v'] > $worstV)) {
                $worstV = $e['v'];
                $worstRn = $rn;
            }
        }

        return $worstRn;
    }

    /**
     * Chain-aware full scan yielding global 1-based row number => raw row
     * (no casts), header rows skipped. Backs the unsorted topRows path.
     *
     * @return \Generator<int, array<int, mixed>>
     */
    private function scanAllRaw(): \Generator
    {
        $chain = $this->chain();
        if ($chain !== null) {
            foreach ($chain as $m) {
                foreach ($this->openSheetReader([], $m['entry'])->rowsFromOffset(null, 1) as $rn => $row) {
                    if ($rn === 1) {
                        continue; // header (real on member 0, repeated on continuations)
                    }
                    yield $m['globalStart'] + ($rn - $m['dataStartLocal']) => $row;
                }
            }

            return;
        }

        foreach ($this->openSheetReader()->rowsFromOffset(null, 1) as $rn => $row) {
            if ($rn === 1) {
                continue;
            }
            yield $rn => $row;
        }
    }

    /**
     * Single-sheet predicate matcher behind rowsWhere(): zone-map
     * pruning + run-merged scans over one entry, yielding LOCAL 1-based
     * row number => raw matched row (casts belong to the caller). This
     * is the pre-v3.2.2 rowsWhere body parameterized by entry.
     *
     * @return \Generator<int, array<int, mixed>>
     */
    private function matchRowsIn(string $entry, int $column, string $op, int|float $value, int|float|null $value2): \Generator
    {
        [$lo, $hi] = $this->pruneBounds($op, $value, $value2);

        $index = $this->loadRandomAccessIndex();
        $stats = $index?->columnStats($entry, $column);

        if ($stats === null) {
            // No zone maps — scan everything, filter per row. NOT via
            // rows(): that wrapper re-keys sequentially from 0, while
            // this generator's contract (and the pruned path below) is
            // 1-based sheet row numbers.
            $this->fireFullScan('rowsWhere', $column, 'column-not-indexed', $entry);
            foreach ($this->openSheetReader([$column - 1], $entry)->rowsFromOffset(null, 1) as $rn => $row) {
                if ($this->cellMatches($row[$column - 1] ?? null, $op, $value, $value2)) {
                    yield $rn => $row;
                }
            }

            return;
        }

        $ranges = $this->blockRanges($index, $entry);

        // Prune, then merge surviving adjacent blocks into runs so each
        // run costs one seek + one linear scan instead of per-block seeks.
        $runs = $this->planRuns($this->survivingBlocks($stats['blocks'], $lo, $hi), $ranges, $this->maxBridgeBytes());

        foreach ($runs as [$firstBlock, $lastBlock]) {
            $start = $ranges[$firstBlock];
            $stopRow = $ranges[$lastBlock]['last_row'];
            $compLength = $this->runCompLength($ranges, $firstBlock, $lastBlock);

            foreach ($this->openSheetReader([$column - 1], $entry)->rowsFromOffset($start['comp_offset'], $start['start_row_at_offset'] ?? 1, $start['first_row'], $compLength) as $rn => $row) {
                if ($rn < $start['first_row']) {
                    continue;
                }
                if ($rn > $stopRow) {
                    break;
                }
                if ($this->cellMatches($row[$column - 1] ?? null, $op, $value, $value2)) {
                    yield $rn => $row;
                }
            }
        }
    }

    /**
     * Single-sheet AND matcher behind rowsWhereAll(): intersect each
     * statful predicate's surviving blocks, scan the (run-merged)
     * intersection once, and yield rows satisfying every predicate.
     * Yields LOCAL 1-based row => raw row (casts belong to the caller),
     * mirroring matchRowsIn's contract.
     *
     * @param  list<array{col: int, op: string, value: int|float, value2: int|float|null}>  $preds
     * @return \Generator<int, array<int, mixed>>
     */
    private function matchRowsAllIn(string $entry, array $preds): \Generator
    {
        $index = $this->loadRandomAccessIndex();

        // Candidate blocks = intersection of the surviving-block sets of
        // every predicate that HAS zone maps. A predicate on an untracked
        // column can't prune (it is filtered per row below), so it simply
        // doesn't constrain the candidate set. null = still unconstrained.
        $candidate = null;
        $readColumns = [];
        foreach ($preds as $p) {
            $readColumns[$p['col'] - 1] = true;
            $stats = $index?->columnStats($entry, $p['col']);
            if ($stats === null) {
                // Validate the op even when it can't prune, so a bad op or
                // a between without an upper bound fails fast like rowsWhere.
                $this->pruneBounds($p['op'], $p['value'], $p['value2']);

                continue;
            }
            [$lo, $hi] = $this->pruneBounds($p['op'], $p['value'], $p['value2']);
            $survivors = $this->survivingBlocks($stats['blocks'], $lo, $hi);
            $candidate = $candidate === null
                ? $survivors
                : array_values(array_intersect($candidate, $survivors));
        }

        $columns = array_keys($readColumns);

        if ($candidate === null) {
            // No statful predicate — full scan, filter per row.
            $this->fireFullScan('rowsWhereAll', null, 'no-indexed-predicate', $entry);
            foreach ($this->openSheetReader($columns, $entry)->rowsFromOffset(null, 1) as $rn => $row) {
                if ($this->rowMatchesAll($row, $preds)) {
                    yield $rn => $row;
                }
            }

            return;
        }

        if ($candidate === []) {
            return; // every block pruned
        }

        $ranges = $this->blockRanges($index, $entry);
        foreach ($this->planRuns($candidate, $ranges, $this->maxBridgeBytes()) as [$firstBlock, $lastBlock]) {
            $start = $ranges[$firstBlock];
            $stopRow = $ranges[$lastBlock]['last_row'];
            $compLength = $this->runCompLength($ranges, $firstBlock, $lastBlock);

            foreach ($this->openSheetReader($columns, $entry)->rowsFromOffset($start['comp_offset'], $start['start_row_at_offset'] ?? 1, $start['first_row'], $compLength) as $rn => $row) {
                if ($rn < $start['first_row']) {
                    continue;
                }
                if ($rn > $stopRow) {
                    break;
                }
                if ($this->rowMatchesAll($row, $preds)) {
                    yield $rn => $row;
                }
            }
        }
    }

    /**
     * True when a raw row satisfies every predicate (numeric cells only,
     * pre-cast values — same per-cell semantics as rowsWhere).
     *
     * @param  array<int, mixed>  $row
     * @param  list<array{col: int, op: string, value: int|float, value2: int|float|null}>  $preds
     */
    private function rowMatchesAll(array $row, array $preds): bool
    {
        foreach ($preds as $p) {
            if (! $this->cellMatches($row[$p['col'] - 1] ?? null, $p['op'], $p['value'], $p['value2'])) {
                return false;
            }
        }

        return true;
    }

    /**
     * Hard upper bound on the rows a single predicate can select: the
     * summed value-count of its zone-map surviving blocks, folded across
     * chain members. null when any member lacks stats for the column
     * (all-or-nothing — a bound with a hole is not a bound).
     */
    private function survivingRowCount(int $column, string $op, int|float $value, int|float|null $value2): ?int
    {
        $index = $this->loadRandomAccessIndex();
        if ($index === null) {
            return null;
        }

        [$lo, $hi] = $this->pruneBounds($op, $value, $value2);
        $chain = $this->chain();
        $entries = $chain !== null ? array_map(fn ($m) => $m['entry'], $chain) : [$this->currentEntry];

        $total = 0;
        foreach ($entries as $entry) {
            $stats = $index->columnStats($entry, $column);
            if ($stats === null) {
                return null;
            }
            foreach ($this->survivingBlocks($stats['blocks'], $lo, $hi) as $b) {
                $total += $stats['blocks'][$b]['count'];
            }
        }

        return $total;
    }

    /**
     * Approximate row count for a single predicate from the sketches:
     * HLL count/distinct for '=', t-digest ΔCDF × count for ranges. null
     * when the column carries no digest/HLL.
     */
    private function estimateSelectivity(int $column, string $op, int|float $value, int|float|null $value2): ?int
    {
        if ($op === '=') {
            $distinct = $this->countDistinct($column);
            $stats = $this->columnStats($column);
            if ($distinct === null || $distinct <= 0 || $stats === null) {
                return null;
            }

            return (int) max(1, round($stats['count'] / $distinct));
        }

        $stats = $this->columnStats($column);
        if ($stats === null) {
            return null;
        }
        $count = $stats['count'];

        if ($op === 'between') {
            $lo = min((float) $value, (float) $value2);
            $hi = max((float) $value, (float) $value2);
            $cLo = $this->estimateCdf($column, $lo);
            $cHi = $this->estimateCdf($column, $hi);
            if ($cLo === null || $cHi === null) {
                return null;
            }

            return (int) round($count * max(0.0, $cHi - $cLo));
        }

        $cdf = $this->estimateCdf($column, (float) $value);
        if ($cdf === null) {
            return null;
        }
        $frac = ($op === '>=' || $op === '>') ? 1.0 - $cdf : $cdf;

        return (int) round($count * max(0.0, $frac));
    }

    /**
     * CDF(x) = fraction of the column's values ≤ x, answered directly by
     * the t-digest's rank() — a single O(centroids) pass through the same
     * piecewise-linear model quantile() uses (rank and quantile are
     * numerical inverses within the digest's resolution). Zero I/O. null
     * when the column has no digest. Values below min → 0, above max → 1.
     *
     * Previously this inverted quantile() by bisecting it 40 times
     * (40 × O(centroids)); rank() collapses that to one O(centroids) pass,
     * ~80× fewer centroid walks for a 'between' estimate (two CDF lookups).
     */
    private function estimateCdf(int $column, float $x): ?float
    {
        $chain = $this->chain();
        if ($chain !== null) {
            return $this->chainDigest($chain, $column)?->rank($x);
        }

        return $this->loadRandomAccessIndex()?->columnDigest($this->currentEntry, $column)?->rank($x);
    }

    /**
     * Plan a predicate set over ONE entry without scanning: intersect the
     * statful predicates' surviving blocks, then count candidate blocks,
     * merged runs, the compressed byte budget of those runs and the
     * row-count upper bound.
     *
     * @param  list<array{col: int, op: string, value: int|float, value2: int|float|null}>  $preds
     * @return array{full: bool, candidateBlocks: int, runs: int, bytes: int, upper: int|null}
     */
    private function planAllIn(?RandomAccessIndex $index, string $entry, array $preds): array
    {
        $candidate = null;
        if ($index !== null) {
            foreach ($preds as $p) {
                $stats = $index->columnStats($entry, $p['col']);
                if ($stats === null) {
                    continue;
                }
                [$lo, $hi] = $this->pruneBounds($p['op'], $p['value'], $p['value2']);
                $survivors = $this->survivingBlocks($stats['blocks'], $lo, $hi);
                $candidate = $candidate === null
                    ? $survivors
                    : array_values(array_intersect($candidate, $survivors));
            }
        }

        $entrySize = $this->cd->entry($entry)['compressed_size'] ?? 0;

        if ($candidate === null) {
            // Full scan: read the whole entry, row count only if indexed.
            return [
                'full' => true,
                'candidateBlocks' => 0,
                'runs' => 1,
                'bytes' => $entrySize,
                'upper' => $index?->totalRows($entry),
            ];
        }

        if ($candidate === []) {
            return ['full' => false, 'candidateBlocks' => 0, 'runs' => 0, 'bytes' => 0, 'upper' => 0];
        }

        $ranges = $this->blockRanges($index, $entry);
        $runs = $this->planRuns($candidate, $ranges, $this->maxBridgeBytes());
        $bytes = 0;
        $upper = 0;
        foreach ($runs as [$firstBlock, $lastBlock]) {
            $startComp = $ranges[$firstBlock]['comp_offset'] ?? 0;
            $end = $ranges[$lastBlock + 1]['comp_offset'] ?? $entrySize;
            $bytes += max(0, $end - $startComp);
            for ($b = $firstBlock; $b <= $lastBlock; $b++) {
                $upper += $ranges[$b]['last_row'] - $ranges[$b]['first_row'] + 1;
            }
        }

        return [
            'full' => false,
            'candidateBlocks' => count($candidate),
            'runs' => count($runs),
            'bytes' => $bytes,
            'upper' => $upper,
        ];
    }

    /**
     * Independence-assumption estimate for a conjunction: total rows ×
     * Π(per-predicate selectivity). null when any predicate has no sketch
     * to estimate from or the population is unknown.
     *
     * @param  list<array{col: int, op: string, value: int|float, value2: int|float|null}>  $preds
     */
    private function estimateAnd(array $preds): ?int
    {
        $population = $this->rowCount();
        if ($population <= 0) {
            return null;
        }

        $product = 1.0;
        foreach ($preds as $p) {
            $est = $this->estimatedRows($p['col'], $p['op'], $p['value'], $p['value2']);
            if ($est === null || $est['estimate'] === null) {
                return null;
            }
            $product *= min(1.0, $est['estimate'] / $population);
        }

        return (int) round($population * $product);
    }

    /**
     * Zone-map survivor test for one predicate: the ascending list of
     * block indices whose [min, max] range can hold a value in [lo, hi]
     * (empty blocks and blocks entirely outside the band are pruned).
     * Sound by construction — a matching row's block always survives, so
     * pruning never drops a match. Reused by rowsWhere (single predicate),
     * rowsWhereAll (intersection) and estimatedRows (bound).
     *
     * @param  list<array{min: float, max: float, sum: float, count: int, other: int}>  $statBlocks
     * @return list<int>
     */
    private function survivingBlocks(array $statBlocks, float|int $lo, float|int $hi): array
    {
        $survivors = [];
        foreach ($statBlocks as $i => $block) {
            if ($block['count'] > 0 && $block['max'] >= $lo && $block['min'] <= $hi) {
                $survivors[] = $i;
            }
        }

        return $survivors;
    }

    /**
     * Collapse an ascending list of surviving block indices into
     * [firstBlock, lastBlock] runs, so each run costs one seek plus one
     * linear scan instead of a seek per block. Adjacent survivors always
     * merge; a gap of pruned blocks merges too (G5 gap-bridging) when its
     * compressed byte span fits $maxBridgeBytes — the reader keeps one
     * ranged read alive through the gap because transferring those bytes
     * beats a fresh round-trip. The gap's non-matching rows are read and
     * filtered out, never mismatched.
     *
     * $maxBridgeBytes = 0 (a zero-latency / no-hints Source) disables
     * bridging entirely, so run merging is byte-for-byte the contiguous
     * behaviour it always had.
     *
     * @param  list<int>  $blocks  ascending, no duplicates
     * @param  list<array{first_row: int, last_row: int, comp_offset: int|null, uncomp_offset: int|null, start_row_at_offset: int|null}>  $ranges
     * @return list<array{0: int, 1: int}>
     */
    private function planRuns(array $blocks, array $ranges, int $maxBridgeBytes): array
    {
        if ($blocks === []) {
            return [];
        }

        $runs = [];
        $runStart = $blocks[0];
        $runEnd = $blocks[0];
        for ($i = 1, $n = count($blocks); $i < $n; $i++) {
            $b = $blocks[$i];
            if ($b === $runEnd + 1) {
                $runEnd = $b;

                continue;
            }

            // Gap = blocks [runEnd+1 .. b-1]; its compressed span is the
            // distance between the sync point after runEnd and the one
            // starting b. Bridge only when both offsets are known and the
            // span fits the round-trip byte budget.
            $gapFrom = $ranges[$runEnd + 1]['comp_offset'] ?? null;
            $gapTo = $ranges[$b]['comp_offset'] ?? null;
            $gapBytes = ($gapFrom !== null && $gapTo !== null) ? $gapTo - $gapFrom : PHP_INT_MAX;

            if ($maxBridgeBytes > 0 && $gapBytes <= $maxBridgeBytes) {
                $runEnd = $b; // bridge the gap into the current run
            } else {
                $runs[] = [$runStart, $runEnd];
                $runStart = $b;
                $runEnd = $b;
            }
        }
        $runs[] = [$runStart, $runEnd];

        return $runs;
    }

    /**
     * Byte budget for gap-bridging: how many compressed bytes transfer in
     * one round-trip on this Source (rtt × bandwidth). A Source without
     * ProvidesCostHints (local disk) is treated as zero-latency → 0 →
     * bridging off. Memoized; the hints don't change within a read.
     */
    private function maxBridgeBytes(): int
    {
        if ($this->bridgeBytesCache === null) {
            if ($this->source instanceof ProvidesCostHints) {
                $h = $this->source->costHints();
                $this->bridgeBytesCache = (int) (($h['rtt_us'] / 1_000_000) * $h['bandwidth_bps']);
            } else {
                $this->bridgeBytesCache = 0;
            }
        }

        return $this->bridgeBytesCache;
    }

    /**
     * Compressed byte length of a surviving block run [firstBlock,
     * lastBlock], from its starting sync point to the sync point that
     * begins the block AFTER it (a ZLIB_FULL_FLUSH boundary). Passed to
     * rowsFromOffset so a pruned scan fetches only the run's bytes
     * instead of ranging to the entry end.
     *
     * Returns null when the run reaches the entry's final block (read to
     * EOF, historical FINISH path) or when the trailing block sits before
     * the first sync point (no recorded offset to bound at).
     *
     * @param  list<array{first_row: int, last_row: int, comp_offset: int|null, uncomp_offset: int|null, start_row_at_offset: int|null}>  $ranges
     */
    private function runCompLength(array $ranges, int $firstBlock, int $lastBlock): ?int
    {
        $next = $ranges[$lastBlock + 1]['comp_offset'] ?? null;
        if ($next === null) {
            return null;
        }

        return $next - ($ranges[$firstBlock]['comp_offset'] ?? 0);
    }

    /**
     * Point lookup: first row whose $column (1-based) equals $value.
     *
     * Mechanism: zone-map pruning over the per-block [min, max] stats
     * plus first-match early exit — not a binary search, and the sorted
     * flag is not consulted. It doesn't need to be: on a column the
     * writer observed to be sorted, at most two adjacent blocks can
     * straddle any value, so the '=' prune alone bounds the lookup to
     * typically ONE block inflate — two S3 range requests end to end on
     * a multi-GB file. On unsorted columns pruning still applies but
     * may leave several candidate blocks. Falls back to a sequential
     * scan when no stats exist (same result, no pruning).
     *
     * @return array{row: int, values: array<int, mixed>}|null
     */
    public function findRow(int|string $column, int|float $value): ?array
    {
        foreach ($this->rowsWhere($column, '=', $value) as $rn => $row) {
            return ['row' => $rn, 'values' => $row];
        }

        return null;
    }

    /**
     * Sorted-group aggregate pushdown: GROUP BY bucket(group column),
     * aggregating another column — sum/count/min/max per group — with
     * whole blocks answered from the sidecar's zone maps instead of row
     * data whenever the sheet allows it.
     *
     * $groupBy and $aggregate are 1-based columns (see columnStats).
     * $bucket maps a numeric group value to its group key; default is
     * identity. It MUST be monotone non-decreasing over the column's
     * value range — pruning concludes "every row in this block is one
     * group" from bucket(min) == bucket(max), which only follows when
     * bucketing preserves order. Threshold buckets (floor of a
     * division, date serial → calendar month, …) qualify:
     *
     *     // date-serial group column → per-month totals (202401, …)
     *     $reader->groupStats(2, 4, fn (float $d) => (int) gmdate('Ym', (int) (($d - 25569) * 86400)));
     *
     * Pushdown applies when the sidecar tracks BOTH columns and saw the
     * group column sorted (asc or desc) — groups then span contiguous
     * blocks. A block whose min and max bucket to the same group, with
     * no non-numeric group cells (other == 0 — the writer counts
     * missing cells there too), is group-pure: the aggregate column's
     * block stats are exactly that group's contribution, so the block
     * is never inflated. Blocks straddling a group boundary are
     * inflated and folded row by row, as is block 0 (the header row
     * folds into its stats for both columns — a numeric-looking header
     * would otherwise leak into the aggregates, so block 0's rows go
     * through the scan path where the row-1 exclusion is
     * authoritative). Blocks whose group column shows count == 0 hold
     * no numeric group values at all — no row in them can join any
     * group — and are skipped without being read. Unsorted or
     * untracked columns degrade to an honest full scan with the same
     * grouping: identical results, no pruning.
     *
     * Row semantics, identical on every path:
     *   - The header row (row 1) never participates — data rows only.
     *   - Rows whose group cell is non-numeric are EXCLUDED, consistent
     *     with rowsWhere()'s numeric-only matching (block stats cannot
     *     see non-numeric values, so no pruned plan could honour them).
     *   - sum/count/min/max cover the group's numeric aggregate cells;
     *     a group whose aggregate cells are all non-numeric still
     *     appears, with count 0, sum 0.0 and null min/max.
     *   - "Numeric" matches the writer's stats fold (statNumericValue):
     *     booleans count as 0/1, because the block stats the pushdown
     *     consumes already include them.
     *
     * Groups are returned in first-encounter sheet order — ascending
     * group keys for an asc-sorted column, descending for desc.
     *
     * On an auto-split continuation chain the aggregation covers the
     * whole logical table: members fold in workbook order, the same
     * group key accumulates across members, and group order stays
     * first-encounter across the chain. Each member keeps its own
     * pushdown verdict (an unsorted member degrades to a scan alone).
     *
     * @param  callable(float): (int|float)  $bucket
     * @return list<array{group: int|float, sum: float, count: int, min: float|null, max: float|null}>
     */
    public function groupStats(int|string $groupBy, int|string $aggregate, ?callable $bucket = null): array
    {
        if (\is_string($groupBy)) {
            $groupBy = $this->resolveColumnName($groupBy) + 1;
        }
        if (\is_string($aggregate)) {
            $aggregate = $this->resolveColumnName($aggregate) + 1;
        }
        if ($groupBy < 1 || $aggregate < 1) {
            throw new \InvalidArgumentException(
                "groupStats() columns are 1-based; got groupBy={$groupBy}, aggregate={$aggregate}"
            );
        }
        $bucket ??= static fn (float $v): float => $v;

        /** @var array<string, array{group: int|float, sum: float, count: int, min: float|null, max: float|null}> $groups */
        $groups = [];

        $chain = $this->chain();
        if ($chain !== null) {
            foreach ($chain as $m) {
                $this->foldGroupStatsFor($m['entry'], $groupBy, $aggregate, $bucket, $groups);
            }
        } else {
            $this->foldGroupStatsFor($this->currentEntry, $groupBy, $aggregate, $bucket, $groups);
        }

        return array_values($groups);
    }

    /**
     * Single-sheet engine behind groupStats(): the pre-v3.2.2 body
     * parameterized by entry, folding into the caller's group map so
     * chain members compose. The rn === 1 exclusions double as the
     * continuation-header filter — a member's repeated header row is
     * its local row 1.
     *
     * @param  callable(float): (int|float)  $bucket
     * @param  array<string, array{group: int|float, sum: float, count: int, min: float|null, max: float|null}>  $groups
     */
    private function foldGroupStatsFor(string $entry, int $groupBy, int $aggregate, callable $bucket, array &$groups): void
    {
        $index = $this->loadRandomAccessIndex();
        $groupCol = $index?->columnStats($entry, $groupBy);
        $aggCol = $index?->columnStats($entry, $aggregate);

        if ($index === null || $groupCol === null || $aggCol === null
            || (! $groupCol['sorted_asc'] && ! $groupCol['sorted_desc'])) {
            // No pushdown basis — honest full scan with the same
            // grouping. Not via rows(): the scan must see 1-based row
            // numbers to exclude exactly the header.
            $reason = ($groupCol === null || $aggCol === null || $index === null)
                ? 'column-not-indexed'
                : 'groupby-not-sorted';
            $this->fireFullScan('groupStats', $groupBy, $reason, $entry);
            foreach ($this->openSheetReader([$groupBy - 1, $aggregate - 1], $entry)->rowsFromOffset(null, 1) as $rn => $row) {
                if ($rn === 1) {
                    continue;
                }
                $this->foldRowIntoGroups($groups, $row, $groupBy, $aggregate, $bucket);
            }

            return;
        }

        $ranges = $this->blockRanges($index, $entry);
        $aggBlocks = $aggCol['blocks'];

        // Plan pass: classify every block in sheet order. Adjacent scan
        // blocks merge into one run so a group boundary costs one seek,
        // not one per block; the plan keeps sheet order so groups
        // register in first-encounter order no matter which step (block
        // stats or row scan) sees them first.
        /** @var list<array{stats: array{0: int|float, 1: int}}|array{scan: array{0: int, 1: int}}> $plan */
        $plan = [];
        $scanRun = null; // [firstBlockIdx, lastBlockIdx] pending row-level scan
        foreach ($groupCol['blocks'] as $i => $block) {
            // count == 0: the block holds NO numeric group values, and
            // only a numeric group cell can place a row in a group —
            // nothing in this block can reach the result. Skipped
            // without being read, whatever its `other` count says.
            if ($block['count'] === 0) {
                if ($scanRun !== null) {
                    $plan[] = ['scan' => $scanRun];
                    $scanRun = null;
                }

                continue;
            }

            // Group-purity test. other == 0 guarantees EVERY row in the
            // block has a numeric group cell, so with a monotone bucket
            // bucket(min) == bucket(max) puts every row in one group —
            // and the aggregate column's block stats then describe
            // exactly this group's rows. Block 0 is exempt (see the
            // docblock: the header folds into its stats).
            $g = $bucket($block['min']);
            if ($i !== 0 && $block['other'] === 0 && $g == $bucket($block['max'])) {
                if ($scanRun !== null) {
                    $plan[] = ['scan' => $scanRun];
                    $scanRun = null;
                }
                $plan[] = ['stats' => [$g, $i]];

                continue;
            }

            // Boundary block — needs its rows.
            $scanRun === null ? $scanRun = [$i, $i] : $scanRun[1] = $i;
        }
        if ($scanRun !== null) {
            $plan[] = ['scan' => $scanRun];
        }

        // Execute pass.
        foreach ($plan as $step) {
            if (isset($step['stats'])) {
                [$g, $i] = $step['stats'];
                $this->foldBlockIntoGroups($groups, $g, $aggBlocks[$i]);

                continue;
            }

            [$first, $last] = $step['scan'];
            $start = $ranges[$first];
            $stopRow = $ranges[$last]['last_row'];
            $compLength = $this->runCompLength($ranges, $first, $last);
            foreach ($this->openSheetReader([$groupBy - 1, $aggregate - 1], $entry)->rowsFromOffset($start['comp_offset'], $start['start_row_at_offset'] ?? 1, $start['first_row'], $compLength) as $rn => $row) {
                if ($rn > $stopRow) {
                    break;
                }
                if ($rn === 1) {
                    continue; // header rides in block 0 — data rows only
                }
                $this->foldRowIntoGroups($groups, $row, $groupBy, $aggregate, $bucket);
            }
        }
    }

    /**
     * Fold one data row into the running group map. The map is keyed
     * by the string cast of the bucket value — PHP arrays truncate
     * float keys to int, which would merge distinct groups — while
     * 'group' preserves the first-seen numeric value for the caller.
     *
     * @param  array<string, array{group: int|float, sum: float, count: int, min: float|null, max: float|null}>  $groups
     * @param  array<int, mixed>  $row
     */
    private function foldRowIntoGroups(array &$groups, array $row, int $groupBy, int $aggregate, callable $bucket): void
    {
        $gv = $this->statNumericCell($row[$groupBy - 1] ?? null);
        if ($gv === null) {
            return; // non-numeric group cell — row excluded by contract
        }

        $g = $bucket($gv);
        $key = (string) $g;
        if (! isset($groups[$key])) {
            $groups[$key] = ['group' => $g, 'sum' => 0.0, 'count' => 0, 'min' => null, 'max' => null];
        }

        $av = $this->statNumericCell($row[$aggregate - 1] ?? null);
        if ($av === null) {
            return; // group registered; nothing numeric to aggregate
        }

        $acc = &$groups[$key];
        $acc['sum'] += $av;
        $acc['count']++;
        $acc['min'] = $acc['min'] === null ? $av : min($acc['min'], $av);
        $acc['max'] = $acc['max'] === null ? $av : max($acc['max'], $av);
    }

    /**
     * Fold a group-pure block's aggregate-column stats into the group
     * map — the pushdown twin of foldRowIntoGroups(): same accumulator
     * shape, fed from the writer's precomputed block aggregates instead
     * of rows.
     *
     * @param  array<string, array{group: int|float, sum: float, count: int, min: float|null, max: float|null}>  $groups
     * @param  array{min: float, max: float, sum: float, count: int, other: int}  $aggBlock
     */
    private function foldBlockIntoGroups(array &$groups, int|float $g, array $aggBlock): void
    {
        $key = (string) $g;
        if (! isset($groups[$key])) {
            $groups[$key] = ['group' => $g, 'sum' => 0.0, 'count' => 0, 'min' => null, 'max' => null];
        }
        if ($aggBlock['count'] === 0) {
            return; // block's aggregate cells were all non-numeric
        }

        $acc = &$groups[$key];
        $acc['sum'] += $aggBlock['sum'];
        $acc['count'] += $aggBlock['count'];
        $acc['min'] = $acc['min'] === null ? $aggBlock['min'] : min($acc['min'], $aggBlock['min']);
        $acc['max'] = $acc['max'] === null ? $aggBlock['max'] : max($acc['max'], $aggBlock['max']);
    }

    /**
     * Numeric interpretation of a reader-side cell for groupStats,
     * mirroring the writer's statNumericValue(): int/float pass
     * through, numeric strings (the tokenizer's t="n" shape) parse,
     * booleans fold as 0/1. The bool case deliberately DIFFERS from
     * cellMatches(): pushdown contributions come from block stats that
     * already include booleans, so the scanned-block path must apply
     * the identical interpretation or the two paths would disagree on
     * bool-bearing sheets. (DateTime never appears here — groupStats
     * reads raw tokenized cells, before any castColumn transform.)
     */
    private function statNumericCell(mixed $cell): ?float
    {
        if (is_int($cell) || is_float($cell)) {
            return (float) $cell;
        }
        if (is_string($cell)) {
            return $cell !== '' && is_numeric($cell) ? (float) $cell : null;
        }
        if (is_bool($cell)) {
            return $cell ? 1.0 : 0.0;
        }

        return null;
    }

    /**
     * Split the current sheet into up to $count independently readable
     * row ranges. Because every sync point in a born-indexed file is a
     * byte-aligned full-flush boundary, each shard can be inflated from
     * a fresh context with zero knowledge of the others — the plan is
     * plain data (JSON-serializable), so the natural use is queue
     * fan-out: dispatch one job per shard, each opening its own reader:
     *
     *   foreach ($reader->shards(8) as $shard) {
     *       ProcessShard::dispatch($s3Path, $shard);
     *   }
     *   // in the job:
     *   $reader = StreamingXlsxReader::fromS3(...);
     *   foreach ($reader->rowsForShard($shard) as $rn => $row) { ... }
     *
     * Shard boundaries snap to sync points, so shards are balanced to
     * within one sync period (default 10K rows). Without a usable index
     * the whole sheet is one shard — same contract, no parallelism.
     * The header row (row 1) rides in the first shard; skip $rn === 1
     * in workers if headers are handled elsewhere.
     *
     * On an auto-split continuation chain the plan covers EVERY member
     * sheet (pre-v3.2.2 it silently sharded the active sheet alone) —
     * $count splits proportionally across members, at least one shard
     * each. Shard row numbers stay local to their sheet, and every
     * member's first shard carries that sheet's repeated header at
     * local row 1, so the "skip $rn === 1" guidance applies per sheet.
     *
     * @return list<array{sheet: string, comp_offset: int|null, start_row: int, first_row: int, last_row: int}>
     */
    public function shards(int $count): array
    {
        if ($count < 1) {
            throw new \InvalidArgumentException("shards() count must be >= 1; got {$count}");
        }

        $index = $this->loadRandomAccessIndex();

        // Chain: fan-out must cover the whole logical table, so shards
        // tile EVERY member sheet — each member gets a slice of the
        // requested count proportional to its rows (never less than
        // one), and workers cross sheets through rowsForShard()'s
        // existing entry switch. Shard row numbers stay LOCAL to their
        // sheet (the rowsForShard contract): each member's first shard
        // carries that sheet's (repeated) header at local row 1, so the
        // documented "skip rn === 1" worker guidance applies per sheet.
        $chain = $this->chain();
        if ($chain !== null && $index !== null) {
            $memberCount = count($chain);
            $effective = max($count, $memberCount);
            $totals = array_column($chain, 'total');
            $sumTotals = array_sum($totals);

            // Largest-remainder allocation, floor-based with a 1-shard
            // minimum per member; converges to exactly $effective.
            $alloc = [];
            $fracs = [];
            $assigned = 0;
            foreach ($totals as $i => $t) {
                $exact = $sumTotals > 0 ? $effective * $t / $sumTotals : 1.0;
                $alloc[$i] = max(1, (int) floor($exact));
                $fracs[$i] = $exact - floor($exact);
                $assigned += $alloc[$i];
            }
            while ($assigned < $effective) {
                $best = array_search(max($fracs), $fracs, true);
                $alloc[$best]++;
                $fracs[$best] = -1.0;
                $assigned++;
            }
            while ($assigned > $effective) {
                $worst = null;
                foreach ($alloc as $i => $a) {
                    if ($a > 1 && ($worst === null || $a > $alloc[$worst])) {
                        $worst = $i;
                    }
                }
                if ($worst === null) {
                    break; // every member at the 1-shard floor already
                }
                $alloc[$worst]--;
                $assigned--;
            }

            $out = [];
            foreach ($chain as $i => $m) {
                $out = array_merge($out, $this->shardsFor($index, $m['entry'], $alloc[$i]));
            }

            return $out;
        }

        $total = $index?->totalRows($this->currentEntry);

        $wholeSheet = [[
            'sheet' => $this->currentEntry,
            'comp_offset' => null,
            'start_row' => 1,
            'first_row' => 1,
            'last_row' => $total ?? PHP_INT_MAX,
        ]];

        if ($index === null || $total === null || $count === 1) {
            return $wholeSheet;
        }

        return $this->shardsFor($index, $this->currentEntry, $count);
    }

    /**
     * Single-sheet shard planner behind shards(): the pre-v3.2.2 body
     * parameterized by entry. Cut points snap to the entry's sync
     * points; duplicate picks collapse, so tiny sheets yield fewer
     * shards than requested.
     *
     * @return list<array{sheet: string, comp_offset: int|null, start_row: int, first_row: int, last_row: int}>
     */
    private function shardsFor(RandomAccessIndex $index, string $entry, int $count): array
    {
        $total = $index->totalRows($entry);

        $wholeSheet = [[
            'sheet' => $entry,
            'comp_offset' => null,
            'start_row' => 1,
            'first_row' => 1,
            'last_row' => $total ?? PHP_INT_MAX,
        ]];

        if ($total === null || $count === 1) {
            return $wholeSheet;
        }

        $points = $index->syncPoints($entry);
        if ($points === []) {
            return $wholeSheet;
        }

        // Pick count-1 cut points from the sync-point list, spaced evenly
        // by ROW coverage (not point ordinal — periods are approximate).
        // Duplicate picks collapse, so tiny sheets yield fewer shards.
        $cuts = [];
        for ($i = 1; $i < $count; $i++) {
            $idealRow = (int) (1 + ($total * $i) / $count);
            $best = null;
            foreach ($points as $sp) {
                if ($best === null || abs($sp['row'] - $idealRow) < abs($best['row'] - $idealRow)) {
                    $best = $sp;
                }
            }
            $cuts[$best['row']] = $best;
        }
        ksort($cuts);

        $shards = [];
        $firstRow = 1;
        $compOffset = null;
        $startRow = 1;
        foreach ($cuts as $cut) {
            if ($cut['row'] <= $firstRow) {
                continue; // collapsed duplicate — would create an empty shard
            }
            $shards[] = [
                'sheet' => $entry,
                'comp_offset' => $compOffset,
                'start_row' => $startRow,
                'first_row' => $firstRow,
                'last_row' => $cut['row'] - 1,
            ];
            $firstRow = $cut['row'];
            $compOffset = $cut['comp_offset'];
            $startRow = $cut['row'];
        }
        $shards[] = [
            'sheet' => $entry,
            'comp_offset' => $compOffset,
            'start_row' => $startRow,
            'first_row' => $firstRow,
            'last_row' => $total,
        ];

        return $shards;
    }

    /**
     * Stream exactly one shard produced by shards() — typically inside
     * a queue worker holding its own reader instance. Yields 1-based
     * row number => row (casts applied) for rows within the shard's
     * [first_row, last_row] span, then stops without touching the rest
     * of the entry.
     *
     * @param  array{sheet: string, comp_offset: int|null, start_row: int, first_row: int, last_row: int}  $shard
     * @return \Generator<int, array<int, mixed>>
     */
    public function rowsForShard(array $shard): \Generator
    {
        foreach (['sheet', 'comp_offset', 'start_row', 'first_row', 'last_row'] as $key) {
            if (! array_key_exists($key, $shard)) {
                throw new \InvalidArgumentException("rowsForShard() shard is missing '{$key}'");
            }
        }
        if ($shard['sheet'] !== $this->currentEntry) {
            // The shard addresses a specific worksheet entry; make the
            // reader agree instead of silently reading the wrong sheet.
            // The switch must go through the SAME per-sheet resets the
            // explicit onSheet()/onSheetIndex() setters perform —
            // skipping them leaves a stale cached header and leaks
            // column casts registered for the previous sheet into this
            // one (both observed as real bugs before this guard).
            foreach ($this->sheets as $sheet) {
                if ($sheet['entry'] === $shard['sheet']) {
                    $this->currentEntry = $shard['sheet'];
                    $this->cachedHeader = null;
                    $this->columnCasts = [];
                    break;
                }
            }
            if ($shard['sheet'] !== $this->currentEntry) {
                throw new \InvalidArgumentException(
                    "rowsForShard() shard references unknown sheet entry '{$shard['sheet']}'"
                );
            }
        }

        // A stale or absent index invalidates recorded offsets — fall
        // back to streaming from the sheet start; row-number bounds
        // still carve out exactly the shard's span.
        $compOffset = $shard['comp_offset'];
        $startRow = $shard['start_row'];
        if ($compOffset !== null && $this->loadRandomAccessIndex() === null) {
            $compOffset = null;
            $startRow = 1;
        }

        // first_row doubles as the fast-forward target — for shards()
        // output it equals start_row (a no-op), but hand-built shards
        // with first_row > start_row (e.g. "skip the header") get the
        // boundary-scan skip instead of tokenizing the gap.
        foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startRow, $shard['first_row']) as $rn => $row) {
            if ($rn < $shard['first_row']) {
                continue;
            }
            if ($rn > $shard['last_row']) {
                return;
            }
            yield $rn => $this->applyCasts($row);
        }
    }

    /**
     * Inclusive [lo, hi] window used only for BLOCK PRUNING. Strict ops
     * deliberately widen to their inclusive counterpart — pruning must
     * over-approximate (a block kept unnecessarily costs one scan; a
     * block dropped incorrectly loses rows). Exact strictness is applied
     * per cell in cellMatches().
     *
     * @return array{0: float, 1: float}
     */
    private function pruneBounds(string $op, int|float $value, int|float|null $value2): array
    {
        return match ($op) {
            '=' => [(float) $value, (float) $value],
            '<', '<=' => [-INF, (float) $value],
            '>', '>=' => [(float) $value, INF],
            'between' => $value2 === null
                ? throw new \InvalidArgumentException("rowsWhere('between') requires a second bound")
                : [(float) min($value, $value2), (float) max($value, $value2)],
            default => throw new \InvalidArgumentException(
                "rowsWhere() op must be one of =, <, <=, >, >=, between; got '{$op}'"
            ),
        };
    }

    /**
     * Exact per-cell predicate. Numeric cells only — the tokenizer
     * yields t="n" values as numeric strings, so is_numeric covers
     * everything the writer's stats folded in except booleans, which
     * the stats deliberately over-include (widening is safe, see
     * BaseXlsxWriter::accumulateColumnStats).
     */
    private function cellMatches(mixed $cell, string $op, int|float $value, int|float|null $value2): bool
    {
        if (is_int($cell) || is_float($cell)) {
            $v = (float) $cell;
        } elseif (is_string($cell) && $cell !== '' && is_numeric($cell)) {
            $v = (float) $cell;
        } else {
            return false;
        }

        return match ($op) {
            '=' => $v == $value,
            '<' => $v < $value,
            '<=' => $v <= $value,
            '>' => $v > $value,
            '>=' => $v >= $value,
            'between' => $v >= min($value, $value2) && $v <= max($value, $value2),
            default => false, // unreachable — pruneBounds validated $op
        };
    }

    /**
     * Memoized view of RandomAccessIndex::blockRanges() for the current
     * sheet — see $blockRangesByEntry for why the reader caches it
     * instead of the (deliberately stateless) index.
     *
     * @return list<array{first_row: int, last_row: int, comp_offset: int|null, uncomp_offset: int|null, start_row_at_offset: int|null}>
     */
    private function blockRanges(RandomAccessIndex $index, ?string $entry = null): array
    {
        $entry ??= $this->currentEntry;

        return $this->blockRangesByEntry[$entry] ??= $index->blockRanges($entry);
    }

    /**
     * Resolve a target row to a (compressed-offset, starting-row)
     * pair the StreamingSheetReader can inflate from. Returns
     * [null, 1] when no index is present or the target precedes
     * every recorded sync point.
     *
     * @return array{0: int|null, 1: int}
     */
    private function seekTarget(int $rowNumber, ?string $entry = null): array
    {
        $index = $this->loadRandomAccessIndex();
        if ($index === null) {
            return [null, 1];
        }
        $sp = $index->findSyncPoint($entry ?? $this->currentEntry, $rowNumber);
        if ($sp === null) {
            return [null, 1];
        }

        return [$sp['comp_offset'], $sp['row']];
    }

    /**
     * Register a cell-value cast for a column — 0-indexed number, or a
     * header NAME (see resolveColumnName; addressing by name makes the
     * 0-based-here vs 1-based-in-queries split invisible). Built-in cast
     * names: date, datetime, int, float, bool. Pass a callable for
     * custom transformations.
     *
     * Casts are applied lazily as rows() / rowAt() / rowRange() yield —
     * the underlying tokenization is unchanged, so registering a cast
     * after iteration began is allowed (subsequent rows will see it).
     *
     * An explicit cast takes precedence over autoDetectDates(): the
     * column is exempted from detection, so the cast always receives
     * the raw tokenized value, never a DateTimeImmutable.
     */
    public function castColumn(int|string $col, string|callable $cast): self
    {
        if (is_string($col)) {
            $col = $this->resolveColumnName($col); // 0-based, castColumn's native base
        }

        // String is always interpreted as a built-in cast name (date,
        // datetime, int, float, bool). is_callable() must NOT win the
        // dispatch here — "date" is a real PHP function, so a naive
        // is_callable() check would pass it and silently call date($v).
        $this->columnCasts[$col] = is_string($cast)
            ? $this->resolveBuiltinCast($cast)
            : $cast;

        return $this;
    }

    /**
     * Bulk variant of castColumn(). Keys may be 0-based numbers or
     * header names. PHP coerces numeric-string ARRAY KEYS ('5') to
     * ints before we ever see them, so a header literally named "5"
     * cannot be addressed through this bulk form — use
     * castColumn('5', ...) for that (the parameter form keeps the
     * string intact).
     *
     * @param  array<int|string, string|callable>  $casts
     */
    public function castColumns(array $casts): self
    {
        foreach ($casts as $col => $cast) {
            $this->castColumn($col, $cast);
        }

        return $this;
    }

    /**
     * Override the timezone applied to cast dates. Default is UTC.
     * Validation happens at config time so a typo surfaces immediately
     * rather than at the first row.
     */
    public function castTimezone(string $tz): self
    {
        try {
            // The validation instance doubles as the per-cell cache —
            // castDate() reuses it instead of re-resolving the tz
            // database for every date cell.
            $this->castTimezoneObj = new \DateTimeZone($tz);
        } catch (\Exception) {
            throw new \InvalidArgumentException("Unknown timezone: {$tz}");
        }
        $this->castTimezone = $tz;

        return $this;
    }

    /**
     * Switch to the 1904 epoch used by some Mac-origin Excel files.
     * Default is the 1900 epoch with the leap-year quirk preserved.
     */
    public function use1904Epoch(): self
    {
        $this->use1904Epoch = true;

        return $this;
    }

    /**
     * Opt in to numFmt-based date detection: numeric cells whose s="N"
     * style resolves to a date number format (built-in ids 14-22 and
     * 45-47, or a custom formatCode carrying date tokens) come back as
     * DateTimeImmutable instead of raw Excel serials. This is how
     * externally-written files (Excel, openpyxl, PhpSpreadsheet) mark
     * dates — the cell itself is just t="n" — so without the opt-in
     * (or a manual castColumn) their date columns read as numbers.
     *
     * Off by default: detection reads xl/styles.xml (fetched lazily,
     * once) and inspects each numeric cell's style, and the raw-serial
     * output shape is a long-standing contract.
     *
     * Semantics:
     *   - $withTime=true keeps the serial's fractional day (datetime);
     *     false truncates to midnight — the castColumn 'date' vs
     *     'datetime' split.
     *   - Conversion shares castDate(): castTimezone() and the 1904
     *     epoch (declared or forced) apply identically.
     *   - An explicit castColumn() on a column takes precedence: its
     *     cells skip detection and the cast sees the raw serial.
     *   - Values castDate() rejects (non-numeric, out of Excel's date
     *     range) pass through unconverted — detection is an inference
     *     and must not destroy data.
     *   - Workbook-wide, like castTimezone(): survives onSheet()
     *     switches (styles.xml is shared by every sheet).
     *   - rowsWhere()/findRow()/groupStats() keep matching and
     *     aggregating the QUERIED columns' raw serials (their contract
     *     is numeric comparison against writer block stats); other
     *     columns in yielded rows are converted normally.
     */
    public function autoDetectDates(bool $withTime = true): self
    {
        $this->autoDetectDates = true;
        $this->autoDetectWithTime = $withTime;

        return $this;
    }

    public function close(): void
    {
        $this->source->close();
    }

    public function __destruct()
    {
        $this->close();
    }

    private function resolveBuiltinCast(string $type): callable
    {
        return match ($type) {
            'date'     => fn ($v) => $this->castDate($v, false),
            'datetime' => fn ($v) => $this->castDate($v, true),
            'int'      => fn ($v) => is_numeric($v) ? (int) $v : null,
            'float'    => fn ($v) => is_numeric($v) ? (float) $v : null,
            'bool'     => fn ($v) => is_bool($v) ? $v : ($v === '1' || $v === 'true'),
            default    => throw new \InvalidArgumentException(
                "Unknown cast type '{$type}'. Use date|datetime|int|float|bool or a callable."
            ),
        };
    }

    private function castDate(mixed $serial, bool $withTime): ?\DateTimeImmutable
    {
        if (! is_numeric($serial)) {
            return null;
        }
        $serial = (float) $serial;

        // Sanity bound: Excel's date range is roughly [0, 2958465] —
        // serial 0 is 1899-12-30, 2958465 is 9999-12-31. Values outside
        // this band almost always indicate a parse error in the source
        // (empty cell read as 0 by an upstream tool, integer ID column
        // mistakenly cast as a date, etc.). Returning null lets callers
        // detect the bad data instead of silently surfacing a date in
        // the year 12345 or before the Gregorian reform.
        if ($serial < 0 || $serial > 2958465) {
            return null;
        }

        // Canonical formula: ts = epoch + serial * 86400, with anchor
        // 1899-12-30 (1900 mode) or 1904-01-01 (Mac mode). The 1900
        // leap-year quirk is *already encoded* in the serial values
        // Excel writes — applying an extra "-1 if >= 60" subtraction
        // here would produce off-by-one dates for every serial >= 60,
        // including the entire post-1900-03-01 range.
        $epochUnix = $this->use1904Epoch ? -2082844800 : -2209161600;

        $whole = (int) $serial;
        $frac = $serial - $whole;
        $secondsOfDay = $withTime ? (int) round($frac * 86400) : 0;

        $ts = $epochUnix + ($whole * 86400) + $secondsOfDay;

        return (new \DateTimeImmutable('@'.$ts))
            ->setTimezone($this->castTimezoneObj ??= new \DateTimeZone($this->castTimezone));
    }

    /**
     * @param  array<int, mixed>  $row
     * @return array<int, mixed>
     */
    private function applyCasts(array $row): array
    {
        if ($this->columnCasts === []) {
            return $row;
        }
        foreach ($this->columnCasts as $col => $cast) {
            if (array_key_exists($col, $row)) {
                $row[$col] = $cast($row[$col]);
            }
        }

        return $row;
    }

    /**
     * @param  list<int>  $rawColumns  0-based columns whose values must
     *                                 stay raw even under autoDetectDates()
     *                                 — the query paths pass their predicate
     *                                 columns here so numeric matching /
     *                                 aggregation semantics don't change
     *                                 underneath them
     */
    private function openSheetReader(array $rawColumns = [], ?string $entry = null): StreamingSheetReader
    {
        return new StreamingSheetReader(
            $this->source,
            $this->cd,
            $entry ?? $this->currentEntry,
            65536,
            $this->resolveSharedStrings(),
            $this->buildDateDetection($rawColumns),
        );
    }

    /**
     * Assemble the tokenizer's date-detection bundle, or null when the
     * opt-in is off / the workbook has no date styles (null keeps the
     * tokenizer on its unchanged fast path). Built per iteration start
     * so the skip set reflects the casts registered at that moment —
     * castColumn() precedence is snapshotted here.
     *
     * @param  list<int>  $rawColumns
     */
    private function buildDateDetection(array $rawColumns): ?DateDetection
    {
        if (! $this->autoDetectDates) {
            return null;
        }

        $bitmap = $this->resolveDateStyles();
        if ($bitmap === null || ! in_array(true, $bitmap, true)) {
            return null;
        }

        $skip = [];
        foreach (array_keys($this->columnCasts) as $col) {
            $skip[$col] = true;
        }
        foreach ($rawColumns as $col) {
            $skip[$col] = true;
        }

        return new DateDetection(
            $bitmap,
            // ?? falls back to the raw value: castDate returns null for
            // non-numeric or out-of-range serials, and a merely-inferred
            // conversion must never eat data an explicit cast would at
            // least surface as null by caller request.
            fn (mixed $v): mixed => $this->castDate($v, $this->autoDetectWithTime) ?? $v,
            $skip,
        );
    }

    /**
     * Lazy-load the styleId → is-date bitmap from xl/styles.xml.
     * Resolved at most once per reader (the part is workbook-global
     * and immutable for our lifetime); archives without styles.xml —
     * rare, but legal when no cell carries s= — resolve to null and
     * detection stays inert.
     *
     * @return list<bool>|null
     */
    private function resolveDateStyles(): ?array
    {
        if ($this->stylesResolved) {
            return $this->dateStyleBitmap;
        }
        $this->stylesResolved = true;

        if (! $this->cd->has('xl/styles.xml')) {
            return $this->dateStyleBitmap = null;
        }

        return $this->dateStyleBitmap = StylesParser::dateStyleBitmap(
            $this->cd->readEntry($this->source, 'xl/styles.xml')
        );
    }

    /**
     * The continuation chain containing the current sheet, or null when
     * the sheet stands alone (single-sheet workbook, no usable sidecar,
     * or a chain of one). Every query API branches on this ONCE — the
     * null path is the pre-v3.2.2 code, byte-for-byte.
     *
     * @return list<array{entry: string, total: int, dataStartLocal: int, globalStart: int}>|null
     */
    private function chain(): ?array
    {
        if (count($this->sheets) < 2) {
            return null;
        }
        $index = $this->loadRandomAccessIndex();
        if ($index === null) {
            return null;
        }

        $this->resolveChains($index);

        $chain = $this->chainByEntry[$this->currentEntry] ?? null;

        return $chain !== null && count($chain) > 1 ? $chain : null;
    }

    /**
     * Partition the workbook's sheets into continuation chains. Runs at
     * most once per reader. Header reads happen ONLY for sheets
     * adjacent to an exactly-full one (the && short-circuit), so
     * workbooks without any full sheet — the overwhelming majority —
     * never pay a single extra read here.
     */
    private function resolveChains(RandomAccessIndex $index): void
    {
        if ($this->chainByEntry !== null) {
            return;
        }
        $this->chainByEntry = [];

        $runs = [];
        $run = [$this->sheets[0]['entry']];
        for ($i = 1, $n = count($this->sheets); $i < $n; $i++) {
            $prev = $this->sheets[$i - 1]['entry'];
            $cur = $this->sheets[$i]['entry'];

            $continues = $index->totalRows($prev) === self::FULL_SHEET_ROWS
                && $index->totalRows($cur) !== null
                && $this->rawHeaderOf($prev) !== []
                && $this->rawHeaderOf($cur) === $this->rawHeaderOf($prev);

            if ($continues) {
                $run[] = $cur;
            } else {
                $runs[] = $run;
                $run = [$cur];
            }
        }
        $runs[] = $run;

        foreach ($runs as $entries) {
            $members = [];
            $globalStart = 1;
            foreach ($entries as $j => $entry) {
                $total = $index->totalRows($entry) ?? 0;
                $dataStartLocal = $j === 0 ? 1 : 2;
                $members[] = [
                    'entry' => $entry,
                    'total' => $total,
                    'dataStartLocal' => $dataStartLocal,
                    'globalStart' => $globalStart,
                ];
                $globalStart += $total - $dataStartLocal + 1;
            }
            foreach ($entries as $entry) {
                $this->chainByEntry[$entry] = $members;
            }
        }
    }

    /**
     * Raw tokenized header row of an arbitrary entry — casts and date
     * detection deliberately bypassed so equality means "same bytes on
     * disk", the only sound continuation test.
     *
     * @return array<int, mixed>
     */
    private function rawHeaderOf(string $entry): array
    {
        if (! array_key_exists($entry, $this->headerByEntry)) {
            $this->headerByEntry[$entry] = [];
            $reader = new StreamingSheetReader($this->source, $this->cd, $entry, 65536, $this->resolveSharedStrings(), null);
            foreach ($reader->rows() as $row) {
                $this->headerByEntry[$entry] = $row;
                break;
            }
        }

        return $this->headerByEntry[$entry];
    }

    /**
     * Global logical row count of a chain: member 1 in full, plus every
     * continuation member's rows minus its repeated header.
     *
     * @param  list<array{entry: string, total: int, dataStartLocal: int, globalStart: int}>  $chain
     */
    private function chainTotalRows(array $chain): int
    {
        $last = end($chain);

        return $last['globalStart'] + $last['total'] - $last['dataStartLocal'];
    }

    /**
     * Map a global row number to [member index, local row] — the
     * inverse of the member globalStart/dataStartLocal bookkeeping.
     * Null when the target lies past the chain's end (or below 1).
     *
     * @param  list<array{entry: string, total: int, dataStartLocal: int, globalStart: int}>  $chain
     * @return array{0: int, 1: int}|null
     */
    private function chainLocate(array $chain, int $globalRow): ?array
    {
        if ($globalRow < 1) {
            return null;
        }
        foreach ($chain as $i => $m) {
            $coverEnd = $m['globalStart'] + $m['total'] - $m['dataStartLocal'];
            if ($globalRow <= $coverEnd) {
                return [$i, $m['dataStartLocal'] + ($globalRow - $m['globalStart'])];
            }
        }

        return null;
    }

    /**
     * Resolve a header NAME to its 0-based column index — the single
     * naming authority behind every int|string API.
     *
     * Semantics (deliberate, uniform):
     *   - A string is ALWAYS a name. '2024' looks up a header called
     *     "2024"; it is never coerced to column 2024 — silent numeric
     *     reinterpretation is exactly the ambiguity this API removes.
     *   - Names match the RAW tokenized header row (before casts /
     *     date detection), compared on the cells' string form. On a
     *     continuation chain every member carries an identical header,
     *     so the active entry's header speaks for the whole chain.
     *   - Empty header cells are not addressable (skipped, not an error).
     *   - A duplicated name throws ambiguousColumnName — silently
     *     picking one column is the silent-wrong-answer failure class.
     *   - An unknown name throws unknownColumnName listing the sheet's
     *     headers, so a typo is a one-glance fix.
     *
     * The map builds once per sheet entry and is memoized; lookups are
     * O(1) after that.
     */
    private function resolveColumnName(string $name): int
    {
        $entry = $this->currentEntry;
        if (! isset($this->columnNamesByEntry[$entry])) {
            $map = [];
            $dupes = [];
            $headers = [];
            foreach ($this->rawHeaderOf($entry) as $idx => $cell) {
                $label = (string) $cell;
                if ($label === '') {
                    continue;
                }
                $headers[] = $label;
                if (isset($dupes[$label])) {
                    $dupes[$label][] = $idx;
                } elseif (isset($map[$label])) {
                    $dupes[$label] = [$map[$label], $idx];
                } else {
                    $map[$label] = $idx;
                }
            }
            $this->columnNamesByEntry[$entry] = ['map' => $map, 'dupes' => $dupes, 'headers' => $headers];
        }

        $names = $this->columnNamesByEntry[$entry];

        if (isset($names['dupes'][$name])) {
            throw XlsxReadException::ambiguousColumnName(
                $name,
                array_map(static fn (int $i): int => $i + 1, $names['dupes'][$name])
            );
        }
        if (! array_key_exists($name, $names['map'])) {
            throw XlsxReadException::unknownColumnName($name, $names['headers']);
        }

        return $names['map'][$name];
    }

    /**
     * Lazy-load and cache the random-access index sidecar if present.
     * Files written without withRandomAccessIndex() carry no sidecar;
     * resolveRandomAccessIndex() returns null in that case and rowAt /
     * rowRange / rowCount fall back to sequential scans.
     */
    private function loadRandomAccessIndex(): ?RandomAccessIndex
    {
        if ($this->indexResolved) {
            return $this->randomAccessIndex;
        }
        $this->indexResolved = true;

        if (! $this->cd->has(RandomAccessIndex::ENTRY_PATH)) {
            return $this->randomAccessIndex = null;
        }

        $payload = $this->cd->readEntry($this->source, RandomAccessIndex::ENTRY_PATH);
        $index = RandomAccessIndex::decode($payload);

        // Stale-index detection: if the workbook was rewritten by Excel
        // or another OOXML editor between our write and this read, the
        // sheet bytes have changed and the cached comp_offset values now
        // point at arbitrary positions in the new deflate stream. The
        // ZIP CD already carries the live sheet CRC32 — comparing it
        // with the value the writer captured into the sidecar catches
        // the divergence with an O(1) lookup. Mismatch silently falls
        // back to a non-indexed scan rather than yielding garbage rows.
        //
        // Plus a structural bound check: each sync point's comp_offset
        // must lie inside the sheet's compressed_size as recorded in
        // the live CD. A crafted sidecar with a CRC-valid but bogus
        // offset would otherwise drive the inflater past the entry
        // body into adjacent ZIP bytes. Treating any out-of-range
        // offset as invalidation also degrades safely to Mode A.
        foreach ($this->sheets as $sheet) {
            $expected = $index->sheetCrc32($sheet['entry']);
            if ($expected === null) {
                continue;
            }
            $cdEntry = $this->cd->entry($sheet['entry']);
            if ($cdEntry === null || $cdEntry['crc32'] !== $expected) {
                return $this->randomAccessIndex = null;
            }

            $compressedSize = $cdEntry['compressed_size'];
            foreach ($index->syncPoints($sheet['entry']) as $sp) {
                if ($sp['comp_offset'] >= $compressedSize) {
                    return $this->randomAccessIndex = null;
                }
            }
        }

        return $this->randomAccessIndex = $index;
    }

    /**
     * Lazy-load the shared-strings table when first needed. Returns null
     * for archives that don't carry one (the kolay-xlsx-stream output
     * shape — every cell uses inlineStr). Throws a clear error for
     * tables larger than SST_RAM_THRESHOLD so callers know the file is
     * outside the supported range without their RAM filling up first.
     *
     * The table is parsed STREAMING: entry chunks flow straight from
     * the inflate loop into SharedStringsParser::push(), so the peak
     * cost is the packed table plus one chunk — the full sst XML never
     * exists in memory (measured ~7x peak reduction vs the previous
     * inflate-then-parse flow on a 1M-entry table).
     */
    private function resolveSharedStrings(): ?SharedStrings
    {
        if ($this->sstResolved) {
            return $this->sst;
        }

        $entry = $this->cd->entry('xl/sharedStrings.xml');
        if ($entry === null) {
            $this->sstResolved = true;

            return null;
        }

        if ($entry['compressed_size'] > self::SST_RAM_THRESHOLD) {
            $sizeMb = number_format($entry['compressed_size'] / 1024 / 1024, 1);
            throw XlsxReadException::corruptCentralDirectory(
                "xl/sharedStrings.xml is {$sizeMb} MB compressed — beyond the in-memory threshold ".
                'this reader supports. On-disk shared-strings tables are not yet implemented.'
            );
        }

        if ($entry['uncompressed_size'] > self::SST_UNCOMPRESSED_THRESHOLD) {
            $uncMb = number_format($entry['uncompressed_size'] / 1024 / 1024, 1);
            $compMb = number_format($entry['compressed_size'] / 1024 / 1024, 1);
            throw XlsxReadException::corruptCentralDirectory(
                "xl/sharedStrings.xml inflates to {$uncMb} MB ({$compMb} MB compressed) — ".
                'beyond the in-memory threshold. Highly repetitive XML can have extreme deflate '.
                'ratios; on-disk shared-strings tables are not yet implemented.'
            );
        }

        // Defense-in-depth: the metadata guard above trusts the
        // uncompressed_size value the ZIP central directory carries. A
        // crafted archive can lie about that field to slip past the
        // bound while the actual inflate balloons RAM. Counting the
        // real inflated bytes as they stream catches the forgery at
        // the first chunk past the declared size — before the excess
        // is ever buffered, let alone parsed. (A well-formed archive
        // inflates to exactly its declared size, so any overshoot is
        // corruption by definition; honest oversized tables were
        // already rejected above.)
        $declared = $entry['uncompressed_size'];
        $inflatedTotal = 0;
        $parser = new SharedStringsParser();
        foreach ($this->cd->streamEntry($this->source, 'xl/sharedStrings.xml') as $chunk) {
            $inflatedTotal += strlen($chunk);
            if ($inflatedTotal > $declared) {
                $actualMb = number_format($inflatedTotal / 1024 / 1024, 1);
                $declaredMb = number_format($declared / 1024 / 1024, 1);
                throw XlsxReadException::corruptCentralDirectory(
                    "xl/sharedStrings.xml inflated to {$actualMb} MB despite a {$declaredMb} MB ".
                    'declared size — ZIP metadata may be corrupt or forged.'
                );
            }
            $parser->push($chunk);
        }

        $this->sst = $parser->finish();
        $this->sstResolved = true;

        return $this->sst;
    }
}

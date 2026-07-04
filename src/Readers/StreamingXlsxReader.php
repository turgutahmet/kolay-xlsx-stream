<?php

namespace Kolay\XlsxStream\Readers;

use Aws\S3\S3Client;
use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;
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
     * 0-based indexing addresses the yielded row arrays).
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
    public function columnStats(int $column): ?array
    {
        $index = $this->loadRandomAccessIndex();
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
    public function quantile(int $column, float $q): ?float
    {
        if ($q < 0.0 || $q > 1.0) {
            throw new \InvalidArgumentException("quantile must be within [0, 1]; got {$q}");
        }

        $digest = $this->loadRandomAccessIndex()?->columnDigest($this->currentEntry, $column);

        return $digest?->quantile($q);
    }

    /**
     * Approximate median — sugar for quantile($column, 0.5).
     */
    public function median(int $column): ?float
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
    public function countDistinct(int $column): ?int
    {
        $hll = $this->loadRandomAccessIndex()?->columnHll($this->currentEntry, $column);

        return $hll?->count();
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
    public function rowsWhere(int $column, string $op, int|float $value, int|float|null $value2 = null): \Generator
    {
        [$lo, $hi] = $this->pruneBounds($op, $value, $value2);

        $index = $this->loadRandomAccessIndex();
        $stats = $index?->columnStats($this->currentEntry, $column);

        if ($stats === null) {
            // No zone maps — scan everything, filter per row. NOT via
            // rows(): that wrapper re-keys sequentially from 0, while
            // this generator's contract (and the pruned path below) is
            // 1-based sheet row numbers.
            foreach ($this->openSheetReader([$column - 1])->rowsFromOffset(null, 1) as $rn => $row) {
                if ($this->cellMatches($row[$column - 1] ?? null, $op, $value, $value2)) {
                    yield $rn => $this->applyCasts($row);
                }
            }

            return;
        }

        $ranges = $this->blockRanges($index);

        // Prune, then merge surviving adjacent blocks into runs so each
        // run costs one seek + one linear scan instead of per-block seeks.
        $runs = [];
        $run = null;
        foreach ($stats['blocks'] as $i => $block) {
            $survives = $block['count'] > 0 && $block['max'] >= $lo && $block['min'] <= $hi;
            if ($survives) {
                $run === null ? $run = [$i, $i] : $run[1] = $i;
            } elseif ($run !== null) {
                $runs[] = $run;
                $run = null;
            }
        }
        if ($run !== null) {
            $runs[] = $run;
        }

        foreach ($runs as [$firstBlock, $lastBlock]) {
            $start = $ranges[$firstBlock];
            $stopRow = $ranges[$lastBlock]['last_row'];

            foreach ($this->openSheetReader([$column - 1])->rowsFromOffset($start['comp_offset'], $start['start_row_at_offset'] ?? 1, $start['first_row']) as $rn => $row) {
                if ($rn < $start['first_row']) {
                    continue;
                }
                if ($rn > $stopRow) {
                    break;
                }
                if ($this->cellMatches($row[$column - 1] ?? null, $op, $value, $value2)) {
                    yield $rn => $this->applyCasts($row);
                }
            }
        }
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
    public function findRow(int $column, int|float $value): ?array
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
     * @param  callable(float): (int|float)  $bucket
     * @return list<array{group: int|float, sum: float, count: int, min: float|null, max: float|null}>
     */
    public function groupStats(int $groupBy, int $aggregate, ?callable $bucket = null): array
    {
        if ($groupBy < 1 || $aggregate < 1) {
            throw new \InvalidArgumentException(
                "groupStats() columns are 1-based; got groupBy={$groupBy}, aggregate={$aggregate}"
            );
        }
        $bucket ??= static fn (float $v): float => $v;

        $index = $this->loadRandomAccessIndex();
        $groupCol = $index?->columnStats($this->currentEntry, $groupBy);
        $aggCol = $index?->columnStats($this->currentEntry, $aggregate);

        /** @var array<string, array{group: int|float, sum: float, count: int, min: float|null, max: float|null}> $groups */
        $groups = [];

        if ($index === null || $groupCol === null || $aggCol === null
            || (! $groupCol['sorted_asc'] && ! $groupCol['sorted_desc'])) {
            // No pushdown basis — honest full scan with the same
            // grouping. Not via rows(): the scan must see 1-based row
            // numbers to exclude exactly the header.
            foreach ($this->openSheetReader([$groupBy - 1, $aggregate - 1])->rowsFromOffset(null, 1) as $rn => $row) {
                if ($rn === 1) {
                    continue;
                }
                $this->foldRowIntoGroups($groups, $row, $groupBy, $aggregate, $bucket);
            }

            return array_values($groups);
        }

        $ranges = $this->blockRanges($index);
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
            foreach ($this->openSheetReader([$groupBy - 1, $aggregate - 1])->rowsFromOffset($start['comp_offset'], $start['start_row_at_offset'] ?? 1, $start['first_row']) as $rn => $row) {
                if ($rn > $stopRow) {
                    break;
                }
                if ($rn === 1) {
                    continue; // header rides in block 0 — data rows only
                }
                $this->foldRowIntoGroups($groups, $row, $groupBy, $aggregate, $bucket);
            }
        }

        return array_values($groups);
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
     * @return list<array{sheet: string, comp_offset: int|null, start_row: int, first_row: int, last_row: int}>
     */
    public function shards(int $count): array
    {
        if ($count < 1) {
            throw new \InvalidArgumentException("shards() count must be >= 1; got {$count}");
        }

        $index = $this->loadRandomAccessIndex();
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

        $points = $index->syncPoints($this->currentEntry);
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
                'sheet' => $this->currentEntry,
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
            'sheet' => $this->currentEntry,
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
    private function blockRanges(RandomAccessIndex $index): array
    {
        return $this->blockRangesByEntry[$this->currentEntry] ??= $index->blockRanges($this->currentEntry);
    }

    /**
     * Resolve a target row to a (compressed-offset, starting-row)
     * pair the StreamingSheetReader can inflate from. Returns
     * [null, 1] when no index is present or the target precedes
     * every recorded sync point.
     *
     * @return array{0: int|null, 1: int}
     */
    private function seekTarget(int $rowNumber): array
    {
        $index = $this->loadRandomAccessIndex();
        if ($index === null) {
            return [null, 1];
        }
        $sp = $index->findSyncPoint($this->currentEntry, $rowNumber);
        if ($sp === null) {
            return [null, 1];
        }

        return [$sp['comp_offset'], $sp['row']];
    }

    /**
     * Register a cell-value cast for a 0-indexed column. Built-in cast
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
    public function castColumn(int $col, string|callable $cast): self
    {
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
     * Bulk variant of castColumn().
     *
     * @param  array<int, string|callable>  $casts
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
    private function openSheetReader(array $rawColumns = []): StreamingSheetReader
    {
        return new StreamingSheetReader(
            $this->source,
            $this->cd,
            $this->currentEntry,
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

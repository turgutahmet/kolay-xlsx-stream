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
 * and loads the table into memory transparently. Archives whose
 * sharedStrings.xml exceeds SST_RAM_THRESHOLD compressed bytes are
 * refused with a clear error — an on-disk variant for very large tables
 * is tracked as a future addition.
 */
class StreamingXlsxReader
{
    /**
     * Compressed-size threshold at which xl/sharedStrings.xml stops
     * fitting comfortably in RAM. ~99% of real-world XLSX files have a
     * sst well below this. Files above the threshold get a clear error
     * pointing at the limitation.
     */
    public const SST_RAM_THRESHOLD = 20 * 1024 * 1024;

    /**
     * Upper bound on the *uncompressed* shared-strings size we will
     * inflate into RAM. Highly repetitive XML (a single repeated <si>
     * entry, common in adversarial inputs and accidentally in some
     * exports) compresses 50:1 or higher, so a 20 MB compressed payload
     * can balloon to 1 GB+. This second guard keeps the bounded-RAM
     * contract intact even when the deflate ratio is pathological.
     */
    public const SST_UNCOMPRESSED_THRESHOLD = 100 * 1024 * 1024;

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

    /** @var array<int, callable> */
    private array $columnCasts = [];

    private bool $use1904Epoch = false;

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
     * to start at row N+1.
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

        $emitted = 0;
        $skipped = 0;

        foreach ($this->openSheetReader()->rows() as $row) {
            if ($skipped < $skip) {
                $skipped++;

                continue;
            }
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
     *     nearest sync point + scan up to $period rows. Effectively
     *     constant time bounded by the writer-chosen sync period.
     *   - Without the sidecar → O(N) — full inflate scan from the
     *     start until the target row is reached.
     *
     * @return array<int, mixed>|null
     */
    public function rowAt(int $rowNumber): ?array
    {
        if ($rowNumber < 1) {
            return null;
        }

        [$compOffset, $startingRow] = $this->seekTarget($rowNumber);

        foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startingRow) as $rn => $row) {
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

        foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startingRow) as $rn => $row) {
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
            foreach ($this->openSheetReader()->rowsFromOffset(null, 1) as $rn => $row) {
                if ($this->cellMatches($row[$column - 1] ?? null, $op, $value, $value2)) {
                    yield $rn => $this->applyCasts($row);
                }
            }

            return;
        }

        $ranges = $index->blockRanges($this->currentEntry);

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

            foreach ($this->openSheetReader()->rowsFromOffset($start['comp_offset'], $start['start_row_at_offset'] ?? 1) as $rn => $row) {
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
            foreach ($this->sheets as $sheet) {
                if ($sheet['entry'] === $shard['sheet']) {
                    $this->currentEntry = $shard['sheet'];
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

        foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startRow) as $rn => $row) {
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

    private function openSheetReader(): StreamingSheetReader
    {
        return new StreamingSheetReader(
            $this->source,
            $this->cd,
            $this->currentEntry,
            65536,
            $this->resolveSharedStrings(),
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

        $sstXml = $this->cd->readEntry($this->source, 'xl/sharedStrings.xml');

        // Defense-in-depth: the metadata guard above trusts the
        // uncompressed_size value the ZIP central directory carries. A
        // crafted archive can lie about that field to slip past the
        // bound while the actual inflate balloons RAM. Once readEntry
        // returns the real bytes, recheck before passing them to the
        // parser — this catches forged metadata and keeps the parser's
        // allocation profile predictable.
        if (strlen($sstXml) > self::SST_UNCOMPRESSED_THRESHOLD) {
            $actualMb = number_format(strlen($sstXml) / 1024 / 1024, 1);
            $declaredMb = number_format($entry['uncompressed_size'] / 1024 / 1024, 1);
            throw XlsxReadException::corruptCentralDirectory(
                "xl/sharedStrings.xml inflated to {$actualMb} MB despite a {$declaredMb} MB ".
                'declared size — ZIP metadata may be corrupt or forged.'
            );
        }

        $this->sst = SharedStringsParser::parseInMemory($sstXml);
        $this->sstResolved = true;

        return $this->sst;
    }
}

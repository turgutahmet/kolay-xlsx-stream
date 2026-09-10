<?php

namespace Kolay\XlsxStream\Writers;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Readers\SharedStrings;
use Kolay\XlsxStream\Readers\SharedStringsParser;
use Kolay\XlsxStream\Sketches\CoMoments;
use Kolay\XlsxStream\Sketches\HyperLogLog;
use Kolay\XlsxStream\Sketches\MisraGries;
use Kolay\XlsxStream\Sketches\TDigest;
use Kolay\XlsxStream\Styles\StyleRegistry;
use Kolay\XlsxStream\Templates\Template;
use Kolay\XlsxStream\Templates\TemplateSheet;

/**
 * Base XLSX Writer - Core streaming functionality
 *
 * Excel limits:
 * - Max rows per sheet: 1,048,576
 * - Max columns: 16,384 (XFD)
 */
abstract class BaseXlsxWriter
{
    protected array $centralDirectory = [];
    protected int $currentOffset = 0;

    // ZIP constants
    public const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
    public const CENTRAL_FILE_HEADER_SIGNATURE = 0x02014b50;
    public const END_OF_CENTRAL_DIR_SIGNATURE = 0x06054b50;
    public const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;

    // Compression methods
    public const COMPRESSION_STORED = 0;
    public const COMPRESSION_DEFLATED = 8;

    // Version info
    public const VERSION_MADE_BY = 0x001E; // 3.0 UNIX
    public const VERSION_NEEDED = 0x0014;  // 2.0

    // Excel limits
    public const MAX_ROWS_PER_SHEET = 1048576; // Excel's hard limit
    public const ROWS_PER_SHEET = 1048575; // MAX - 1 for header safety
    public const MAX_COLUMNS = 16384; // Excel column limit (XFD)

    // ZIP32 container limits — exceeding any of these requires ZIP64,
    // which the writer does not yet emit. Guarded calls turn silent
    // 32-bit truncation into a loud, actionable exception.
    private const ZIP32_MAX_SIZE = 0xFFFFFFFF;   // 4 GB - 1
    private const ZIP32_MAX_ENTRIES = 0xFFFF;    // 65535

    // Style ids registered in getStylesXml() cellXfs
    public const STYLE_DEFAULT = 0;
    public const STYLE_DATETIME = 1;

    // Excel built-in numFmtIds (0-49 reserved range, no <numFmt> entry).
    // Pass these to setColumnFormat() for locale-aware formatting:
    // BUILTIN_NUMFMT_DATE renders dd.mm.yyyy in tr-TR, mm/dd/yyyy in en-US.
    public const BUILTIN_NUMFMT_GENERAL = 0;
    public const BUILTIN_NUMFMT_INTEGER = 1;     // 0
    public const BUILTIN_NUMFMT_DECIMAL_2 = 2;   // 0.00
    public const BUILTIN_NUMFMT_THOUSANDS = 3;   // #,##0
    public const BUILTIN_NUMFMT_CURRENCY = 5;    // $#,##0_);($#,##0)
    public const BUILTIN_NUMFMT_PERCENT = 9;     // 0%
    public const BUILTIN_NUMFMT_PERCENT_2 = 10;  // 0.00%
    public const BUILTIN_NUMFMT_EXPONENT = 11;   // 0.00E+00
    public const BUILTIN_NUMFMT_FRACTION = 12;   // # ?/?
    public const BUILTIN_NUMFMT_DATE = 14;       // m/d/yyyy (locale-aware)
    public const BUILTIN_NUMFMT_DATE_LONG = 15;  // d-mmm-yy
    public const BUILTIN_NUMFMT_TIME_AMPM = 18;  // h:mm AM/PM
    public const BUILTIN_NUMFMT_TIME = 20;       // h:mm
    public const BUILTIN_NUMFMT_DATETIME = 22;   // m/d/yyyy h:mm (locale-aware)

    // Excel epoch: serial 1 = 1900-01-01, but Excel mistakenly treats 1900 as a leap year.
    // Using 1899-12-30 as base so Unix timestamps map correctly for all post-1900 dates.
    public const EXCEL_EPOCH_TIMESTAMP = -2209161600; // 1899-12-30 00:00:00 UTC

    // XML escaping (hot path). PCRE with a JIT-compiled character class is
    // used as the "does this string need work?" gate instead of strpbrk:
    // strpbrk is a naive O(n*m) double loop and the 34-char needle (escape
    // chars + control bytes) makes clean strings — the overwhelming common
    // case — pay ~5x more than a JIT class, which compiles to a 256-bit
    // bitmap and scans O(n) with a flat constant. Both patterns must cover
    // exactly: & < > " ' plus XML-1.0-invalid control bytes (0x00-0x08,
    // 0x0B, 0x0C, 0x0E-0x1F — tab/LF/CR excluded, they are legal).
    private const XML_ESCAPE_NEEDED = '/[&<>"\'\x00-\x08\x0B\x0C\x0E-\x1F]/';
    private const XML_CTRL_BYTES = '/[\x00-\x08\x0B\x0C\x0E-\x1F]/';
    private const XML_ESCAPE_MAP = [
        '&' => '&amp;',
        '<' => '&lt;',
        '>' => '&gt;',
        '"' => '&quot;',
        "'" => '&apos;',
    ];

    // Sheet management
    protected array $sheets = [];
    protected int $currentSheetIndex = 0;
    protected array $columns = [];

    // Current sheet streaming variables
    protected $deflateCtx = null;
    protected $crcContext = null;
    protected int $sheetCrc = 0;
    protected int $sheetUncompressedSize = 0;
    protected int $sheetCompressedSize = 0;
    protected int $sheetOffset = 0;
    protected int $currentSheetRow = 0;
    protected int $totalRows = 0;

    // Writer state
    protected bool $started = false;
    protected bool $closed = false;
    protected bool $compactMode = false;

    // Performance optimizations
    protected int $bufferFlushInterval = 1000;
    protected string $rowBuffer = '';
    protected int $rowBufferCount = 0;
    // Level 5 is the measured knee of the speed/size curve for XLSX-shaped
    // XML: on a 300K-row mixed workload it produced a file within 0.2% of
    // level 6's size at ~22% less total wall time (level 6 spends its extra
    // effort on entropy — unique cell refs — that doesn't compress anyway).
    protected int $deflateLevel = 5;

    // Column letter cache for performance
    protected array $colLetterCache = [];

    // Progress reporting
    protected ?\Closure $progressCallback = null;
    protected int $progressInterval = 10000;

    // Styling (v2.2+)
    protected StyleRegistry $styles;
    protected ?int $headerStyleId = null;

    /** @var array<int, int> 1-based column index => cellXfs style id */
    protected array $columnStyleIds = [];

    /** @var array<int, string> 1-based column index => format preset name (used for auto-width sizing) */
    protected array $columnFormatNames = [];

    /**
     * Memoized merge of a per-row style with a column number format.
     * Keyed [rowStyleId][columnStyleId] => merged cellXfs id. Keeps the
     * styled hot path off StyleRegistry's O(n) dedup scan after the first
     * time a (row-style, column) pair is seen.
     *
     * @var array<int, array<int, int>>
     */
    protected array $rowStyleMergeCache = [];

    // Sheet view options (v2.2+)
    protected int $freezeRows = 0;
    protected int $freezeColumns = 0;
    protected bool $autoFilterEnabled = false;

    /** @var array<int, float> 1-based column index => width in characters */
    protected array $columnWidths = [];
    protected bool $autoColumnWidth = false;

    // Sample-based auto column width (opt-in via setAutoColumnWidth(sample: N))
    protected ?int $autoWidthSampleSize = null;
    protected bool $autoWidthStrict = false;
    /** @var list<string> Buffered row XML strings while sampling */
    protected array $autoWidthSampleBuffer = [];
    protected int $autoWidthSampleBufferBytes = 0;

    /**
     * Hard cap on accumulated sample-buffer byte size. A misconfigured
     * sample (very wide rows × large sample size) can otherwise hold
     * 100+ MB in memory. When this ceiling is hit we force-finalize
     * early — the sample is "good enough" by then and emitting the
     * preamble lets writeRow exit sample mode and stream normally.
     */
    private const SAMPLE_MAX_BUFFER_BYTES = 8 * 1024 * 1024;
    /** @var array<int, int> 1-based col index => max char length seen */
    protected array $autoWidthMaxLengths = [];
    protected bool $autoWidthFinalized = false;
    protected bool $inSampleMode = false;
    /**
     * Cols whose width was last derived by sample finalize. Cleared at
     * the start of the next sheet so sample widths don't leak between
     * sheets the way user-explicit setColumnWidths() entries do.
     *
     * @var list<int>
     */
    protected array $sampleAutoSetWidthCols = [];

    // Custom sheet name for the next sheet rotation (set by newSheet()).
    protected ?string $nextSheetName = null;

    // Random-access index (opt-in via withRandomAccessIndex). When enabled,
    // ZLIB_FULL_FLUSH is injected at row-buffer boundaries roughly every
    // $indexSyncPeriod rows, and an xl/_kxs/index.bin sidecar is written on
    // finishFile(). Default off — every previously-written byte is identical.
    protected bool $randomAccessIndexEnabled = false;

    protected int $indexSyncPeriod = 10000;
    protected int $rowsSinceSync = 0;

    // Group-boundary sync (opt-in via syncAtGroupBoundaries). When set,
    // a ZLIB_FULL_FLUSH is forced whenever the 1-based $groupSyncColumn's
    // value changes, so each index block holds exactly one group and
    // groupStats() folds it straight from the sidecar with zero row reads.
    protected ?int $groupSyncColumn = null;
    protected ?string $lastGroupKey = null;

    /**
     * Per-sheet sync points keyed by sheet entry path.
     *
     * @var array<string, list<array{row: int, comp_offset: int, uncomp_offset: int}>>
     */
    protected array $indexSyncPoints = [];

    /**
     * Running CRC32 of the sheet's uncompressed bytes at each sync point,
     * keyed by sheet entry path and aligned 1:1 with $indexSyncPoints —
     * value k covers exactly the first uncomp_offset bytes of sync point
     * k. Serialized as the KXSI "SCRC" TLV section.
     *
     * @var array<string, list<int>>
     */
    protected array $indexSyncPointCrcs = [];

    // Column statistics (opt-in via withColumnStats). For each tracked
    // column the writer accumulates min/max/sum/count per index block —
    // the row span between two sync points — plus a per-sheet sorted
    // flag. Serialized as a KXSI "STAT" TLV section; enables the reader to
    // skip whole blocks for range predicates (zone maps, à la Parquet
    // row-group stats), answer column aggregates from the sidecar alone,
    // and binary-search sorted key columns in O(log blocks).

    /** @var list<int> 1-based column indexes tracked for stats */
    protected array $statsColumns = [];

    /** @var array<int, array{min: float, max: float, sum: float, count: int, other: int}> current-block accumulator per column */
    protected array $statsAccum = [];

    /** @var array<int, array{asc: bool, desc: bool, prev: float|null}> per-sheet sortedness tracker per column */
    protected array $statsSorted = [];

    /** @var array<string, array<int, list<array{min: float, max: float, sum: float, count: int, other: int}>>> closed blocks: entry => col => blocks */
    protected array $indexColumnBlocks = [];

    /** @var array<string, array<int, array{asc: bool, desc: bool}>> final sortedness per sheet: entry => col => flags */
    protected array $indexColumnSorted = [];

    // String zone maps (STRZ) — opt-in via withStringStats, orthogonal to
    // withColumnStats. Per-block lexicographic [min, max] string prefixes
    // for string range/prefix pruning. Full min/max are kept per open
    // block, capped to STRING_STAT_CAP at block close, then the truncation
    // length is chosen per column at finishFile (deferred shortest
    // separator — restores pruning on common-prefix corpora).

    /** Max bytes of a stored block min/max prefix. */
    protected const STRING_STAT_CAP = 64;

    /** @var list<int> 1-based columns tracked for string zone maps */
    protected array $stringStatsColumns = [];

    /** @var array<int, array{min: ?string, max: ?string, count: int, other: int}> current-block string accumulator */
    protected array $stringAccum = [];

    /** @var array<int, array{asc: bool, desc: bool, prev: ?string}> per-sheet lexicographic sortedness */
    protected array $stringSorted = [];

    /** @var array<string, array<int, list<array{min: ?string, max: ?string, count: int, other: int}>>> closed string blocks: entry => col => blocks */
    protected array $indexStringBlocks = [];

    /** @var array<string, array<int, array{asc: bool, desc: bool}>> final string sortedness per sheet */
    protected array $indexStringSorted = [];

    // Column sketches (opt-in via withColumnSketches, orthogonal to
    // withColumnStats). For each tracked column the writer feeds one
    // whole-sheet t-digest (numeric values, same inclusion rule as STAT)
    // and one HyperLogLog (canonical string of every non-empty value).
    // Serialized as the KXSI "TDIG" and "CHLL" TLV sections; the reader
    // answers quantile/median/countDistinct from the sidecar alone.
    // The header row is EXCLUDED — see withColumnSketches().

    /** @var list<int> 1-based column indexes tracked for sketches */
    protected array $sketchColumns = [];

    /** @var array<int, TDigest> current-sheet t-digest per tracked column */
    protected array $sketchDigestAccum = [];

    /** @var array<int, HyperLogLog> current-sheet HLL per tracked column */
    protected array $sketchHllAccum = [];

    /**
     * Row-side staging for the sketches: per tracked column, numeric
     * values and canonical strings collect here and bulk-feed the
     * sketches (addMany) every SKETCH_FLUSH_ROWS rows. Batching halves
     * the measured per-cell cost versus per-value add() calls — the
     * dominant expense was PHP call plumbing, not the sketch math.
     *
     * @var array<int, list<float>>
     */
    protected array $sketchNumBuffer = [];

    /** @var array<int, list<string>> */
    protected array $sketchStrBuffer = [];

    protected int $sketchRowsBuffered = 0;

    protected const SKETCH_FLUSH_ROWS = 512;

    /** @var array<string, array<int, string>> finished sheets: entry => col => serialized TDigest */
    protected array $indexColumnDigests = [];

    /** @var array<string, array<int, string>> finished sheets: entry => col => serialized HyperLogLog */
    protected array $indexColumnHlls = [];

    // Range/group quantiles (TDGB) — opt-in via withRangeQuantiles. One
    // t-digest per ROW-SPACE superblock per column: a superblock spans
    // SUPERBLOCK_ROWS rows, its end snapped to the first sync point past
    // that width, so a sheet holds ≤64 superblocks regardless of block
    // count and each digest is built single-pass (no write-time merges).

    /** Superblock row width; end snaps to the first sync point past it. */
    protected const SUPERBLOCK_ROWS = 16384;

    /** @var list<int> 1-based columns tracked for range/group quantiles */
    protected array $rangeQuantileColumns = [];

    /** @var array<int, \Kolay\XlsxStream\Sketches\TDigest> current open superblock digest per column */
    protected array $rangeDigestAccum = [];

    /** Rows accumulated in the current (shared) open superblock. */
    protected int $rangeSuperblockRows = 0;

    /** @var array<string, array<int, list<array{end_row: int, payload: string}>>> closed superblocks: entry => col => list */
    protected array $indexRangeSuperblocks = [];

    /** @var list<int> 1-based columns tracked for top-K frequent values (TOPK) */
    protected array $topValuesColumns = [];

    protected int $topValuesK = MisraGries::DEFAULT_K;

    /** @var array<int, MisraGries> current-sheet Misra-Gries sketch per tracked column */
    protected array $topValueSketches = [];

    /** @var array<string, array<int, string>> finished sheets: entry => col => serialized MisraGries */
    protected array $indexTopValues = [];

    /** @var list<int> 1-based columns tracked for argmin/argmax row pointers (ARGP) */
    protected array $argPointerColumns = [];

    /** @var array<int, bool> O(1) membership set for argPointerColumns */
    protected array $argPointerSet = [];

    /** @var array<int, array{minRow: int, maxRow: int}> current block's arg rows per column (0 = none) */
    protected array $argAccum = [];

    /** @var array<string, array<int, list<array{minRow: int, maxRow: int}>>> finished: entry => col => per-block arg rows */
    protected array $indexArgPointers = [];

    /** @var list<int> 1-based columns tracked for pairwise correlation (CORR) */
    protected array $correlationColumns = [];

    /** @var list<array{0: int, 1: int}> the C(k,2) column pairs (colA < colB) */
    protected array $correlationPairs = [];

    /** @var array<string, CoMoments> current-sheet co-moment accumulator per pair key "a,b" */
    protected array $coMomentAccum = [];

    /** @var array<string, array<string, string>> finished: entry => pairKey => serialized CoMoments */
    protected array $indexCorrelations = [];

    // ── Template mode (v3.5) ─────────────────────────────────────────────
    // The writer keeps a foreign workbook's layout and streams rows into one
    // sheet's <sheetData>. Every field below is null/false on the classic
    // path, which is why a single boolean guards the row builder dispatch.

    protected ?Template $template = null;
    protected bool $templateMode = false;

    /** True when this writer opened the template itself and must close it. */
    protected bool $templateOwned = false;

    protected ?SharedStrings $templateSharedStringsReader = null;
    protected ?TemplateSheet $templateSheet = null;

    /** Sheet currently being streamed: its workbook name and its zip entry. */
    protected ?string $templateSheetName = null;
    protected ?string $templateSheetEntry = null;

    /** @var array<string, true> template sheet entries already streamed */
    protected array $templateStreamedEntries = [];

    protected ?SharedStringTable $sharedStrings = null;

    /**
     * variant => 0-based column index => cellXfs id, and variant => the
     * pre-rendered `<row r="N"` suffix. Both are read once per row, so they
     * are resolved when the sheet opens rather than per row.
     *
     * @var list<array<int, int>>
     */
    protected array $templateStyleMaps = [];

    /** @var list<string> */
    protected array $templateRowSuffixes = [];

    /**
     * cellXfs id for a date written into a column the sample rows never
     * styled. Registered lazily, so a template whose data is all text or
     * numbers keeps its styles.xml byte-identical.
     */
    protected ?int $templateDateStyleId = null;

    public function __construct()
    {
        $this->styles = new StyleRegistry();

        // Laravel-published config defaults. Guarded so the package
        // stays framework-optional: outside Laravel config() either
        // does not exist or has no container behind it (the helper is
        // defined by illuminate/support even without an app booted —
        // hence the try/catch, not just function_exists).
        if (\function_exists('config')) {
            try {
                $cfg = config('xlsx-stream');
            } catch (\Throwable) {
                $cfg = null;
            }
            $this->applyConfigDefaults(is_array($cfg) ? $cfg : null);
        }
    }

    /**
     * Fold the published config's writer defaults into this instance.
     *
     * Precedence: code-level setters > config > package defaults. The
     * config is applied once at construction, so any setter the caller
     * invokes afterwards naturally overrides it.
     *
     * Version gate: copies of config/xlsx-stream.php published before
     * v3.2.2 carried keys the package never read, with stale values
     * that contradict the code defaults (compression_level 1 vs the
     * writer's real default 5). Honouring them retroactively would
     * silently change existing applications' output, so only configs
     * declaring `'version' => 2` (the v3.2.2+ file) are applied —
     * older published copies stay inert, exactly as they always were.
     *
     * Invalid values are silently ignored rather than thrown: this
     * runs in the constructor, and a config file is environment data,
     * not code — a bad env var must not turn every `new Writer` into
     * a 500. The setters keep throwing for code-level misuse.
     */
    protected function applyConfigDefaults(?array $config): void
    {
        if ((int) ($config['version'] ?? 0) < 2) {
            return;
        }

        $level = $config['writer']['compression_level'] ?? null;
        if (is_numeric($level)) {
            try {
                $this->setCompressionLevel((int) $level);
            } catch (\Throwable) {
                // out-of-range config value — keep the package default
            }
        }

        $interval = $config['writer']['buffer_flush_interval'] ?? null;
        if (is_numeric($interval)) {
            try {
                $this->setBufferFlushInterval((int) $interval);
            } catch (\Throwable) {
                // out-of-range config value — keep the package default
            }
        }
    }

    /**
     * Write data to destination (must be implemented by child classes)
     */
    abstract protected function writeToDest(string $data): void;

    /**
     * Set deflate compression level (1-9)
     * 3 = fast (larger files)
     * 5 = balanced (default — within ~0.2% of level 6's size at ~20% less wall time)
     * 6 = marginally smaller, measurably slower
     * 9 = best compression (much slower, ~6% smaller)
     */
    public function setCompressionLevel(int $level): self
    {
        if ($level < 1 || $level > 9) {
            throw XlsxStreamException::invalidCompressionLevel($level);
        }
        $this->deflateLevel = $level;
        return $this;
    }

    /**
     * Opt in to COMPACT output: rows and cells are written WITHOUT the
     * r attribute. ECMA-376 declares both c/@r and row/@r optional —
     * readers assign positions sequentially when they are absent, which
     * is why empty cells still emit a `<c/>` placeholder (the
     * placeholder IS the position carrier; see the compact row
     * builders).
     *
     * Why it pays: `r="AB123456"` is a UNIQUE string per cell — high-
     * entropy noise deflate cannot back-reference. Dropping it turns
     * the row body into a near-pure template. Measured on this machine
     * (100K-row mixed workload, apple-to-apple; ratios vary with data
     * shape): compressed sheet bytes −52 % (up to ~−62 % on sparse
     * data), writing −26 % and reading −33 % faster — the smaller XML
     * is quicker to both build and tokenize.
     *
     * Interoperability, verified per this package's real-app testing
     * discipline: Microsoft Excel, LibreOffice and Apple Numbers open
     * compact files (manual tests, styles/formats/freeze/autofilter
     * included); PhpSpreadsheet 5.8 and OpenSpout 4.28 read them
     * correctly — both implement sequential fallback as first-class
     * code paths — and this package's own reader never consumed r in
     * the first place. Still opt-in while the mode accumulates field
     * mileage; classic output remains byte-identical when off.
     *
     * Must be called before startFile(): flipping mid-sheet would mix
     * referenced and sequential cells and corrupt position assignment.
     */
    public function compact(bool $enabled = true): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode('compact()', 'the sample rows carry r attributes that the compact shape drops');
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        $this->compactMode = $enabled;

        return $this;
    }

    /**
     * Set row buffer flush interval
     * Lower = more responsive streaming
     * Higher = better compression ratio
     */
    /**
     * Enable born-indexed output: write xl/_kxs/index.bin sidecar so a
     * matched reader can do O(1) random-access lookups (rowAt, rowRange,
     * rowCount). Backward-compatible — Excel, PhpSpreadsheet, OpenSpout
     * etc. ignore the unreferenced part and read the file normally.
     *
     * Approximate cost (per the POC benchmark, 4M rows / 10K period):
     *   wall time ≈ ölçüm gürültüsü, RAM ≈ +10 KB, file size ≈ +0.04 %.
     *
     * Calling this method has no effect once startFile() has been called.
     */
    public function withRandomAccessIndex(int $every = 10000): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($every < 1) {
            throw new XlsxStreamException(
                "Index sync period must be at least 1; got {$every}."
            );
        }

        $this->randomAccessIndexEnabled = true;
        $this->indexSyncPeriod = $every;

        return $this;
    }

    /**
     * Align index block boundaries to GROUP changes: force a sync point
     * (ZLIB_FULL_FLUSH) whenever the 1-based $column's value differs from
     * the previous row's, so each block holds exactly one group.
     *
     * Why it pays: groupStats() folds a group-pure block straight from the
     * sidecar's per-block aggregates without reading a row. Aligning
     * blocks to groups makes EVERY block pure, so "GROUP BY on a grouped
     * export" answers from the sidecar alone — zero row reads, even on S3.
     * Pair it with withColumnStats() on the group and aggregate columns.
     *
     * Rows must be written grouped (all of a group's rows consecutive),
     * as grouped exports already are; an interleaved column just produces
     * more, smaller blocks (still correct). Enables the random-access
     * index if it is not already on. Must be called before startFile().
     * The row-count sync period still applies as an upper bound, so a
     * single huge group is capped into several (same-group) blocks.
     */
    public function syncAtGroupBoundaries(int $column): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($column < 1) {
            throw new XlsxStreamException(
                "Group-sync column is 1-based; got {$column}."
            );
        }

        $this->groupSyncColumn = $column;
        $this->randomAccessIndexEnabled = true;

        return $this;
    }

    /**
     * Track per-block column statistics (zone maps) for the given 1-based
     * columns and embed them in the random-access sidecar (a "STAT" TLV
     * section pre-v3.1 readers transparently ignore).
     *
     * What the matched reader gains, all without scanning row data:
     *   - rowsWhere() skips every block whose [min,max] cannot satisfy
     *     the predicate (exports are usually ID/date-sorted, so range
     *     queries typically touch a handful of blocks);
     *   - columnStats() answers min/max/sum/count for the whole sheet
     *     from the ~KB sidecar alone — on S3 that is one range request
     *     for a multi-GB file;
     *   - findRow() binary-searches a column the writer observed to be
     *     monotonically sorted, resolving a point lookup by reading a
     *     single block.
     *
     * Statistics cover values that render as numeric cells: int/float,
     * numeric strings, DateTime (as Excel serial), bool (as 0/1).
     * Anything else counts toward the block's "other" tally and never
     * causes a block to be skipped incorrectly — stats widen, never
     * narrow, so pruning stays sound.
     *
     * Cost: a few comparisons per tracked cell and 32 bytes per
     * (block × column) in the sidecar — ~13 KB for 4M rows / 10K period.
     *
     * Implies withRandomAccessIndex() (blocks are the spans between its
     * sync points); enables it with the default period if not already on.
     */
    public function withColumnStats(array $columns): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($columns === []) {
            throw new XlsxStreamException('withColumnStats() needs at least one column index.');
        }
        foreach ($columns as $col) {
            if (! is_int($col) || $col < 1 || $col > self::MAX_COLUMNS) {
                throw new XlsxStreamException(
                    'withColumnStats() expects 1-based integer column indexes; got: '.var_export($col, true)
                );
            }
        }

        if (! $this->randomAccessIndexEnabled) {
            $this->withRandomAccessIndex();
        }

        $columns = array_values(array_unique($columns));
        sort($columns);
        $this->statsColumns = $columns;

        return $this;
    }

    /**
     * Opt in to per-block STRING zone maps (STRZ) for the given 1-based
     * columns — the lexicographic analogue of withColumnStats(), enabling
     * string range / prefix pruning and string findRow. Comparison is
     * unsigned byte-wise (= Unicode code-point order, NOT a locale
     * collation; see SPEC §4.5). Like STAT, this implies the random-access
     * index. Orthogonal to withColumnStats(): a column may carry numeric
     * stats, string zone maps, both, or neither.
     */
    public function withStringStats(array $columns): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($columns === []) {
            throw new XlsxStreamException('withStringStats() needs at least one column index.');
        }
        foreach ($columns as $col) {
            if (! is_int($col) || $col < 1 || $col > self::MAX_COLUMNS) {
                throw new XlsxStreamException(
                    'withStringStats() expects 1-based integer column indexes; got: '.var_export($col, true)
                );
            }
        }

        if (! $this->randomAccessIndexEnabled) {
            $this->withRandomAccessIndex();
        }

        $columns = array_values(array_unique($columns));
        sort($columns);
        $this->stringStatsColumns = $columns;

        return $this;
    }

    /**
     * Track file-level approximate-statistics sketches for the given
     * 1-based columns and embed them in the random-access sidecar
     * ("TDIG" + "CHLL" TLV sections older readers transparently ignore).
     *
     * Per sheet, per tracked column, the writer maintains:
     *   - a merging t-digest (δ=100, ~1-4 KB serialized) over the
     *     column's numeric values — the reader answers quantile() and
     *     median() from the sidecar alone;
     *   - a HyperLogLog (p=11, ~2 KB, ±~2.3%) over the canonical string
     *     of every non-empty value — countDistinct() likewise. Text
     *     columns are first-class here: non-numeric values are invisible
     *     to the t-digest (like STAT) but fully counted by the HLL.
     *
     * Both sketches merge associatively, so per-sheet (and, later,
     * per-segment) sketches can be combined without touching row data.
     *
     * The header row is EXCLUDED from both sketches — deliberately the
     * opposite of withColumnStats(), which folds the header into block 0.
     * Zone maps are pruning structures: missing a matchable row breaks
     * query soundness, so STAT must over-include. Sketches are estimates
     * of the DATA distribution: folding a header in cannot make any
     * answer "safer", it can only bias quantiles and inflate distinct
     * counts by one phantom value.
     *
     * Orthogonal to withColumnStats() — enable either or both; each
     * implies withRandomAccessIndex() (default period if not already on).
     *
     * Cost: ~0.4 µs per tracked cell (one xxh64 hash + canonical string
     * + a batched digest fold — measured ~10 % wall time with 2 tracked
     * columns on a realistic 12-column export), and ~3-6 KB of sidecar
     * per (sheet × column).
     */
    public function withColumnSketches(array $columns): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($columns === []) {
            throw new XlsxStreamException('withColumnSketches() needs at least one column index.');
        }
        foreach ($columns as $col) {
            if (! is_int($col) || $col < 1 || $col > self::MAX_COLUMNS) {
                throw new XlsxStreamException(
                    'withColumnSketches() expects 1-based integer column indexes; got: '.var_export($col, true)
                );
            }
        }

        if (! $this->randomAccessIndexEnabled) {
            $this->withRandomAccessIndex();
        }

        $columns = array_values(array_unique($columns));
        sort($columns);
        $this->sketchColumns = $columns;

        return $this;
    }

    /**
     * Track the top-K most frequent values of each given column (1-based),
     * answerable from the sidecar (topValues() / the categorical slice of
     * profile()) with zero row reads. Values fold as their canonical string
     * form (the CHLL rule, §4.4), so text and numeric columns are both
     * covered. Bounded to k counters per column (reference k = 64): while a
     * column's cardinality stays ≤ k the stored counts are EXACT — the
     * complete categorical distribution — and above k they become top-k
     * with error ≤ N/k, distinguished by the sketch's saturated bit. The
     * header row is excluded, as for the other sketches. Implies the
     * random-access index. Must be called before startFile().
     */
    public function withTopValues(array $columns, int $k = MisraGries::DEFAULT_K): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($columns === []) {
            throw new XlsxStreamException('withTopValues() needs at least one column index.');
        }
        foreach ($columns as $col) {
            if (! is_int($col) || $col < 1 || $col > self::MAX_COLUMNS) {
                throw new XlsxStreamException(
                    'withTopValues() expects 1-based integer column indexes; got: '.var_export($col, true)
                );
            }
        }

        if (! $this->randomAccessIndexEnabled) {
            $this->withRandomAccessIndex();
        }

        $columns = array_values(array_unique($columns));
        sort($columns);
        $this->topValuesColumns = $columns;
        $this->topValuesK = $k;

        return $this;
    }

    /**
     * Track the sheet row number of each given column's minimum and maximum
     * per block (ARGP), so argMin()/argMax() answer "which row holds the
     * extreme" from the sidecar with no scan. Aligned 1:1 with STAT blocks
     * and defined over the same values STAT sees (it is STAT that ranks the
     * blocks), so these columns are folded into withColumnStats() too —
     * call order does not matter. First-occurrence semantics: the earliest
     * row achieving a block's min/max is the one recorded. Implies the
     * random-access index. Must be called before startFile().
     */
    public function withArgPointers(array $columns): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($columns === []) {
            throw new XlsxStreamException('withArgPointers() needs at least one column index.');
        }
        foreach ($columns as $col) {
            if (! is_int($col) || $col < 1 || $col > self::MAX_COLUMNS) {
                throw new XlsxStreamException(
                    'withArgPointers() expects 1-based integer column indexes; got: '.var_export($col, true)
                );
            }
        }

        $columns = array_values(array_unique($columns));
        sort($columns);
        $this->argPointerColumns = $columns;
        $this->argPointerSet = array_fill_keys($columns, true);

        // argMin/argMax rank blocks by STAT min/max, so every argptr column
        // must also carry STAT — union in, never clobber an existing set.
        $this->statsColumns = array_values(array_unique(array_merge($this->statsColumns, $columns)));
        sort($this->statsColumns);

        if (! $this->randomAccessIndexEnabled) {
            $this->withRandomAccessIndex();
        }

        return $this;
    }

    /**
     * Track pairwise Pearson correlation among the given 1-based columns:
     * every unordered pair gets one CoMoments accumulator (n, Σx, Σy, Σxy,
     * Σx², Σy²), so a reader answers correlation(a, b) from the sidecar with
     * no scan. Only rows where BOTH cells are numeric (the same numeric
     * interpretation STAT uses — DateTime as Excel serial, bool as 0/1,
     * numeric strings as their value) feed a pair; the header is excluded.
     *
     * Cost is quadratic in the column count: k columns hold C(k, 2) pairs at
     * 48 bytes each and cost k²/2 multiply-adds per row, so this is meant for
     * a handful of measure columns, not every column. Accumulators are
     * mergeable, so correlations compose across an auto-split chain like the
     * sketches. Implies the random-access index. Call before startFile().
     */
    public function withCorrelations(array $columns): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        foreach ($columns as $col) {
            if (! is_int($col) || $col < 1 || $col > self::MAX_COLUMNS) {
                throw new XlsxStreamException(
                    'withCorrelations() expects 1-based integer column indexes; got: '.var_export($col, true)
                );
            }
        }

        $columns = array_values(array_unique($columns));
        sort($columns);
        if (count($columns) < 2) {
            throw new XlsxStreamException('withCorrelations() needs at least two distinct columns.');
        }

        $this->correlationColumns = $columns;
        $pairs = [];
        for ($i = 0, $n = count($columns); $i < $n; $i++) {
            for ($j = $i + 1; $j < $n; $j++) {
                $pairs[] = [$columns[$i], $columns[$j]];
            }
        }
        $this->correlationPairs = $pairs;

        if (! $this->randomAccessIndexEnabled) {
            $this->withRandomAccessIndex();
        }

        return $this;
    }

    /**
     * Opt in to per-superblock t-digests (TDGB) for the given 1-based
     * columns — range- and group-scoped approximate quantiles
     * (`quantile(col, q, from, to)`, and group quantiles composed over a
     * group's row range). Unlike withColumnSketches (one whole-sheet
     * digest), this keeps one digest per ROW-SPACE superblock so a query
     * can merge just the superblocks covering its range. Implies the
     * random-access index. Numeric values only (same inclusion rule as
     * withColumnStats). Must be called before startFile().
     */
    public function withRangeQuantiles(array $columns): self
    {
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($columns === []) {
            throw new XlsxStreamException('withRangeQuantiles() needs at least one column index.');
        }
        foreach ($columns as $col) {
            if (! is_int($col) || $col < 1 || $col > self::MAX_COLUMNS) {
                throw new XlsxStreamException(
                    'withRangeQuantiles() expects 1-based integer column indexes; got: '.var_export($col, true)
                );
            }
        }

        if (! $this->randomAccessIndexEnabled) {
            $this->withRandomAccessIndex();
        }

        $columns = array_values(array_unique($columns));
        sort($columns);
        $this->rangeQuantileColumns = $columns;

        return $this;
    }

    /**
     * One-call preset for a fully queryable export: turn on the
     * random-access index, per-block zone maps (columnStats) and, unless
     * $withSketches is false, the t-digest/HLL sketches — all for the
     * given 1-based $columns. Equivalent to chaining withRandomAccessIndex()
     * + withColumnStats() + withColumnSketches(), and byte-identical to
     * doing so with the same arguments.
     *
     * Must be called before startFile().
     */
    public function queryable(array $columns, int $every = 10000, bool $withSketches = true): self
    {
        $this->withRandomAccessIndex($every);
        $this->withColumnStats($columns);
        if ($withSketches) {
            $this->withColumnSketches($columns);
        }

        return $this;
    }

    public function setBufferFlushInterval(int $rows): self
    {
        if ($rows < 1) {
            throw XlsxStreamException::invalidBufferSize($rows);
        }
        $this->bufferFlushInterval = $rows;

        return $this;
    }

    /**
     * Register a progress callback fired every $progressInterval rows.
     *
     * Signature: function(int $rowsWritten, int $bytesWritten): void
     *
     * Useful for queue jobs that update a progress UI:
     *
     *   $writer->onProgress(fn($rows, $bytes) =>
     *       Cache::put("export:$jobId", compact('rows','bytes'))
     *   );
     */
    public function onProgress(callable $callback): self
    {
        $this->progressCallback = $callback instanceof \Closure
            ? $callback
            : \Closure::fromCallable($callback);

        return $this;
    }

    /**
     * Set how often the progress callback fires (in rows).
     */
    public function setProgressInterval(int $rows): self
    {
        if ($rows < 1) {
            throw XlsxStreamException::invalidBufferSize($rows);
        }
        $this->progressInterval = $rows;

        return $this;
    }

    /**
     * Style the header row (the first row written by startFile).
     *
     * Options:
     *  - bold (bool)
     *  - color (#RRGGBB) — text color
     *  - fill (#RRGGBB)  — background color
     *  - size (int)      — font size, default 11
     *
     * Example:
     *
     *   $writer->setHeaderStyle([
     *       'bold' => true,
     *       'fill' => '#4F81BD',
     *       'color' => '#FFFFFF',
     *   ]);
     */
    public function setHeaderStyle(array $options): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode('setHeaderStyle()', 'the template owns its header rows and their styles');
        $this->finalizePendingSample();
        // Allowed before startFile() (for sheet 1) and between newSheet() calls
        // (to give each manually-rotated sheet its own header style).
        $this->headerStyleId = $this->styles->registerHeaderStyle($options);

        return $this;
    }

    /**
     * Register a reusable per-row style and return its id, to be passed as
     * the second argument of writeRow().
     *
     * Options are identical to setHeaderStyle(): fill (background #RRGGBB),
     * color (text #RRGGBB), bold, size, name. Register each logical style
     * once up front, then stamp it on whichever rows you choose:
     *
     *   $failed = $writer->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
     *   $vip    = $writer->registerRowStyle(['fill' => '#1F4E78', 'color' => '#FFFFFF']);
     *
     *   foreach ($rows as $r) {
     *       $writer->writeRow($r->toArray(), $r->failed ? $failed : null);
     *   }
     *
     * Styles are dedup'd, so one logical style is a single styles.xml entry
     * no matter how many rows use it — memory stays flat. Passing null (the
     * writeRow default) keeps a row on the unstyled fast path at zero cost.
     *
     * A styled row still honours column number formats: a currency or date
     * column keeps its format and gains the row's fill/color (the two are
     * merged on the fly, also dedup'd).
     */
    public function registerRowStyle(array $options): int
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode(
            'registerRowStyle()',
            'a template row is styled by its variant, so the id could never be used'
        );

        return $this->styles->registerRowStyle($options);
    }

    /**
     * Freeze the first row so it stays visible while scrolling.
     *
     * Equivalent to Excel's "View > Freeze Top Row".
     */
    public function freezeFirstRow(): self
    {
        $this->refuseInTemplateMode('freezeFirstRow()', 'the template owns the <sheetViews> block');
        return $this->freezeRowsAndColumns(1, 0);
    }

    /**
     * Freeze a custom number of rows and/or columns.
     *
     *   $writer->freezeRowsAndColumns(rows: 1, columns: 2);
     */
    public function freezeRowsAndColumns(int $rows = 1, int $columns = 0): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode('freezeRowsAndColumns()', 'the template owns the <sheetViews> block');
        $this->finalizePendingSample();
        if ($rows < 0 || $columns < 0) {
            throw new XlsxStreamException('Freeze rows/columns must be >= 0.');
        }
        // Settable any time before close — applies to the next sheet started
        // via writeRow() or newSheet().
        $this->freezeRows = $rows;
        $this->freezeColumns = $columns;

        return $this;
    }

    /**
     * Set explicit column widths in Excel character units.
     *
     *   $writer->setColumnWidths([1 => 8, 2 => 30, 3 => 15]);
     *
     * Keys are 1-based column indexes. Columns omitted from the array
     * fall back to Excel's default width.
     *
     * @param array<int, float|int> $widths
     */
    public function setColumnWidths(array $widths): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode('setColumnWidths()', 'the template owns the <cols> block');
        $this->finalizePendingSample();
        foreach ($widths as $col => $width) {
            if ($col < 1) {
                throw new XlsxStreamException("Column index must be >= 1, got {$col}.");
            }
            if ($width <= 0) {
                throw new XlsxStreamException("Column width must be > 0, got {$width}.");
            }
            $this->columnWidths[$col] = (float) $width;
        }

        return $this;
    }

    /**
     * Auto-size columns.
     *
     * Two modes — choose based on the precision/cost tradeoff:
     *
     *   Heuristic (default, cost: zero per row):
     *     $writer->setAutoColumnWidth();
     *   Width = max(format-min, mb_strlen(header) + 2, 8.43).
     *
     *   Sample-based (opt-in, cost: O(sample) bytes RAM):
     *     $writer->setAutoColumnWidth(sample: 1000);
     *   Buffers the first N data rows, computes per-column max char
     *   length, then drains. Catches columns where the data is wider
     *   than the header — phone numbers, descriptions, IDs.
     *
     *   Strict mode for sample (default lenient):
     *     $writer->setAutoColumnWidth(sample: 1000, strict: true);
     *   Strict propagates any internal failure during width tracking.
     *   Lenient (default) catches it, logs to error_log, and falls back
     *   to the heuristic — no broken file shipped under load.
     *
     * Manual setColumnWidths() entries always win over both modes.
     */
    public function setAutoColumnWidth(bool|int $sample = true, bool $strict = false): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode('setAutoColumnWidth()', 'the template owns the <cols> block');
        $this->finalizePendingSample();

        if (is_int($sample)) {
            if ($sample < 0) {
                throw new XlsxStreamException(
                    "Auto-column-width sample size must be >= 0, got {$sample}."
                );
            }
            $this->autoColumnWidth = true;
            $this->autoWidthSampleSize = $sample > 0 ? $sample : null;
        } else {
            $this->autoColumnWidth = $sample;
            $this->autoWidthSampleSize = null;
        }
        $this->autoWidthStrict = $strict;

        return $this;
    }

    /**
     * Add Excel's auto-filter dropdowns to the header row.
     *
     * The filter range is computed automatically from the sheet's columns
     * and final row count when the sheet is closed.
     */
    public function enableAutoFilter(bool $enabled = true): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode('enableAutoFilter()', 'the template owns the auto filter');
        $this->autoFilterEnabled = $enabled;

        return $this;
    }

    /**
     * Apply a number format to a column (1-based).
     *
     * Accepts:
     *   - Preset name (see StyleRegistry::PRESETS) — `'date'`, `'currency_try'`
     *   - Raw Excel format code — `'0.000'`, `'#,##0.00'`
     *   - Built-in numFmtId int (0-49 reserved range) — `BUILTIN_NUMFMT_DATE`
     *
     * Built-in ids render with the *reader's* locale (e.g. dd.mm.yyyy in
     * tr-TR, mm/dd/yyyy in en-US) — useful when the same export ships to
     * users in multiple regions. Custom format codes are locale-stable.
     *
     *   $writer->setColumnFormat(2, 'date');                                // YYYY-MM-DD literal
     *   $writer->setColumnFormat(3, 'currency_try');                        // #,##0.00 ₺
     *   $writer->setColumnFormat(4, '0.000');                               // raw code
     *   $writer->setColumnFormat(5, BaseXlsxWriter::BUILTIN_NUMFMT_DATE);   // locale-aware
     *   $writer->setColumnFormat(6, 'General', raw: true);                  // verbatim, skip preset lookup
     *
     * Unknown preset-shaped strings throw (a typo'd preset written as a
     * literal formatCode sends Excel into repair mode); pure date-token
     * runs like 'dddd'/'mmss' pass as raw codes, and $raw = true skips
     * preset resolution and the guard entirely — the escape hatch for
     * arbitrary codes coming from external constant tables (e.g.
     * PhpSpreadsheet NumberFormat values).
     *
     * Numeric values in that column are wrapped with the chosen format.
     * String values pass through unchanged.
     */
    public function setColumnFormat(int $column, string|int $format, bool $raw = false): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode('setColumnFormat()', 'the sample row is the format oracle and the template number format wins');
        $this->finalizePendingSample();
        if ($column < 1) {
            throw new XlsxStreamException("Column index must be >= 1, got {$column}.");
        }
        // Range validation is deferred to startNewSheet() — at this point the
        // user may be pre-configuring formats for an upcoming newSheet() call
        // whose column count is different from the current sheet's.

        if (is_int($format)) {
            if ($format < 0 || $format > 49) {
                throw new XlsxStreamException(
                    "Built-in numFmtId must be 0-49 (Excel reserved range); got {$format}. ".
                    'Pass a string preset or raw format code for custom formats.'
                );
            }
            $this->columnStyleIds[$column] = $this->styles->registerBuiltinNumFmt($format);
            // Tag the format name so the auto-width heuristic recognises it
            // as a known formatted column (defaults to the format-min path).
            $this->columnFormatNames[$column] = 'builtin:'.$format;

            return $this;
        }

        $this->columnStyleIds[$column] = $this->styles->registerColumnFormat($format, $raw);
        // Stored separately so the auto-width heuristic can pick a sensible
        // minimum based on the format (e.g. currency cells need ~14 chars
        // even when the header is shorter).
        $this->columnFormatNames[$column] = $format;

        return $this;
    }

    /**
     * Clear all column-level formats and widths.
     *
     * Useful between newSheet() calls when the next sheet has a different
     * column layout — without this, formats set on the previous sheet leak
     * into the next sheet's numeric cells and auto-width calculations.
     */
    public function clearColumnFormats(): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        $this->refuseInTemplateMode('clearColumnFormats()', 'the sample row is the format oracle and the template number format wins');
        $this->finalizePendingSample();
        $this->columnStyleIds = [];
        $this->columnFormatNames = [];
        $this->columnWidths = [];

        return $this;
    }

    /**
     * Convert Unix timestamp to DOS time and date (separate fields)
     */
    protected function dosTimeParts(int $timestamp): array
    {
        $d = getdate($timestamp);

        // DOS Time: bits 15-11: hours, 10-5: minutes, 4-0: seconds/2
        $dosTime = (($d['hours'] & 0x1F) << 11) |
                   (($d['minutes'] & 0x3F) << 5) |
                   (($d['seconds'] >> 1) & 0x1F);

        // DOS Date: bits 15-9: year-1980, 8-5: month, 4-0: day
        $dosDate = ((($d['year'] - 1980) & 0x7F) << 9) |
                   (($d['mon'] & 0x0F) << 5) |
                   (($d['mday'] & 0x1F));

        return [$dosTime, $dosDate];
    }

    /**
     * Start XLSX file with headers and static files
     */
    public function startFile(array $headers): void
    {
        $this->refuseInTemplateMode('startFile()', 'the template already supplies the header rows; choose a sheet with sheet($name, $dataStartRow)');
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        if (count($headers) > self::MAX_COLUMNS) {
            throw XlsxStreamException::tooManyColumns(count($headers), self::MAX_COLUMNS);
        }

        $this->columns = $headers;
        $this->started = true;

        // Static files that don't depend on sheet count or registry state.
        // styles.xml is deferred to finishFile() so styles registered between
        // newSheet() calls (or during streaming) are all included.
        $this->writeStaticFile('_rels/.rels', $this->getRelsXml());
    }

    /**
     * Write a single row (handles multi-sheet automatically).
     *
     * $styleId stamps a style registered with registerRowStyle() over the
     * whole row. $variant selects which of a template's sample rows this row
     * should look like, and is meaningful only in template mode. They are
     * separate arguments on purpose: one parameter never changes meaning
     * with the writer's mode, so a style id can never be read as a variant.
     */
    public function writeRow(array $row, ?int $styleId = null, int $variant = 0): void
    {
        if (!$this->started) {
            throw $this->templateMode
                ? XlsxStreamException::templateSheetNotSelected()
                : XlsxStreamException::headersNotSet();
        }
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        if (count($row) > self::MAX_COLUMNS) {
            throw XlsxStreamException::tooManyColumns(count($row), self::MAX_COLUMNS);
        }

        // Check if we need to start a new sheet
        if ($this->currentSheetRow === 0 || $this->currentSheetRow >= self::ROWS_PER_SHEET) {
            $this->rollSheet();
        }

        // Group-boundary sync: when this row starts a new group, flush the
        // buffered previous group WITH a sync point so its block ends here.
        // Runs before currentSheetRow++ so the sync point (recorded as
        // currentSheetRow + 1) lands on this row — the first of the new
        // group's block. The empty-buffer guard skips the first row and
        // fresh-sheet starts.
        if ($this->groupSyncColumn !== null) {
            $value = $row[$this->groupSyncColumn - 1] ?? null;
            $key = is_scalar($value) ? (string) $value : '';
            if ($this->lastGroupKey !== null && $key !== $this->lastGroupKey && $this->rowBuffer !== '') {
                $this->flushRowBuffer(true);
            }
            $this->lastGroupKey = $key;
        }

        $this->currentSheetRow++;
        $this->totalRows++;

        if ($this->statsColumns !== []) {
            $this->accumulateColumnStats($row);
        }
        if ($this->stringStatsColumns !== []) {
            $this->accumulateStringStats($row);
        }
        if ($this->sketchColumns !== []) {
            $this->accumulateColumnSketches($row);
        }
        if ($this->rangeQuantileColumns !== []) {
            $this->accumulateRangeQuantiles($row);
        }
        if ($this->topValuesColumns !== []) {
            $this->accumulateTopValues($row);
        }
        if ($this->correlationPairs !== []) {
            $this->accumulateCorrelations($row);
        }

        // One boolean picks the builder family, the same price the compact
        // shape already pays; the classic builders stay untouched so their
        // golden-file bytes cannot drift.
        $rowXml = $this->templateMode
            ? ($this->sharedStrings !== null
                ? $this->buildRowXmlTemplateShared($this->currentSheetRow, $row, $variant, $styleId)
                : $this->buildRowXmlTemplate($this->currentSheetRow, $row, $variant, $styleId))
            : $this->buildRowXml($this->currentSheetRow, $row, $styleId);

        // Sample-mode width tracking (opt-in). Buffers row XML + records
        // per-column max char length until the sample size is reached or
        // finishFile() drains a partial sample.
        if ($this->inSampleMode && ! $this->autoWidthFinalized) {
            $this->trackSampledRow($row, $rowXml);
            // Two finalize triggers — sample-count target reached, or
            // accumulated XML bytes hit the safety cap. The byte cap
            // protects against a misconfigured wide-row × large-sample
            // combo that would otherwise hold tens of MB in memory.
            if (
                count($this->autoWidthSampleBuffer) >= $this->autoWidthSampleSize
                || $this->autoWidthSampleBufferBytes >= self::SAMPLE_MAX_BUFFER_BYTES
            ) {
                $this->finalizeAutoWidthSample();
            }
            if ($this->progressCallback !== null && $this->totalRows % $this->progressInterval === 0) {
                ($this->progressCallback)($this->totalRows, $this->currentOffset);
            }

            return;
        }

        // Normal path: append to buffer + periodic flush.
        $this->rowBuffer .= $rowXml;
        $this->rowBufferCount++;

        if ($this->rowBufferCount >= $this->bufferFlushInterval) {
            $this->flushRowBuffer();
        }

        if ($this->progressCallback !== null && $this->totalRows % $this->progressInterval === 0) {
            ($this->progressCallback)($this->totalRows, $this->currentOffset);
        }
    }

    /**
     * Sheet boundary reached: finalize the current sheet and open the next.
     *
     * Template mode never auto-splits — an overflow sheet would have no
     * layout to inherit — so it either returns (a template whose data starts
     * on row 1 legitimately leaves the counter at zero before its first row)
     * or refuses at Excel's row ceiling.
     */
    protected function rollSheet(): void
    {
        if ($this->templateMode) {
            if ($this->currentSheetRow >= self::ROWS_PER_SHEET) {
                throw XlsxStreamException::templateSheetRowLimit(
                    (string) $this->templateSheetName,
                    self::ROWS_PER_SHEET
                );
            }

            return;
        }

        if ($this->currentSheetRow > 0) {
            $this->flushRowBuffer();
            $this->finishCurrentSheet();
        }
        $this->currentSheetIndex++;
        $this->startNewSheet();
    }

    /**
     * Sample-mode width tracker. Records each cell's char length in
     * autoWidthMaxLengths and buffers the row's XML for later replay.
     *
     * Lenient mode (default) catches any internal failure here, logs to
     * error_log, and bails out of sample mode — preferring a valid file
     * with heuristic widths over an HTTP 500. Strict mode propagates
     * the exception so callers see the failure during testing.
     */
    protected function trackSampledRow(array $row, string $rowXml): void
    {
        try {
            $this->updateAutoWidthMaxLengths($row);
            $this->autoWidthSampleBuffer[] = $rowXml;
            $this->autoWidthSampleBufferBytes += strlen($rowXml);
        } catch (\Throwable $e) {
            if ($this->autoWidthStrict) {
                throw $e;
            }
            error_log(sprintf(
                'kolay-xlsx-stream: auto-width sample failed at row %d (%s) — falling back to heuristic',
                count($this->autoWidthSampleBuffer),
                $e->getMessage()
            ));
            $this->autoWidthSampleBuffer[] = $rowXml;
            $this->autoWidthSampleBufferBytes += strlen($rowXml);
            $this->bailFromSampleMode();
        }
    }

    /**
     * Inner width-tracker. Factored out so subclasses (and tests) can
     * override the failure surface without re-implementing the
     * try/catch + bail logic in trackSampledRow().
     */
    protected function updateAutoWidthMaxLengths(array $row): void
    {
        foreach ($row as $i => $cell) {
            $col = $i + 1;

            // DateTime objects can't be cast to string and would crash a
            // naive (string) coercion. buildRowXml renders them as Excel
            // serial numbers — the user-visible width in Excel is bounded
            // by "yyyy-mm-dd hh:mm:ss" (19 characters), so use that as a
            // conservative upper bound without invoking the formatter.
            if ($cell instanceof \DateTimeInterface) {
                $len = 19;
            } else {
                $len = mb_strlen((string) $cell);
            }

            if (! isset($this->autoWidthMaxLengths[$col]) || $len > $this->autoWidthMaxLengths[$col]) {
                $this->autoWidthMaxLengths[$col] = $len;
            }
        }
    }

    /**
     * Sample size reached — compute widths, emit preamble (now with the
     * computed <cols>) + header, drain the sample buffer.
     */
    /**
     * Finalize the active sheet's pending auto-width sample, if any.
     *
     * Config mutations must never act retroactively on an already-
     * written (or pending) sheet preamble. In sample mode the preamble
     * is deferred until the sample drains — without this guard, a
     * mutator called while the sample is pending (the natural "prepare
     * the NEXT sheet, then newSheet()" flow) would rewrite the CURRENT
     * sheet's header style, wipe its explicit widths, or leak the next
     * sheet's formats into it. Every config mutator that feeds the
     * preamble calls this first, so a pending sample is finalized with
     * exactly the state it was sampled under before the mutation lands.
     *
     * Cost when no sample is pending: three property reads.
     */
    protected function finalizePendingSample(): void
    {
        if ($this->started && $this->inSampleMode && ! $this->autoWidthFinalized) {
            $this->finalizeAutoWidthSample();
        }
    }

    protected function finalizeAutoWidthSample(): void
    {
        foreach ($this->autoWidthMaxLengths as $col => $maxLen) {
            // Honour any explicit setColumnWidths() entry the user already set.
            if (isset($this->columnWidths[$col])) {
                continue;
            }
            $width = min(255.0, max(8.43, (float) ($maxLen + 2)));
            $this->columnWidths[$col] = $width;
            $this->sampleAutoSetWidthCols[] = $col;
        }
        $this->autoWidthFinalized = true;
        $this->inSampleMode = false;

        $this->writeSheetData($this->buildSheetPreambleXml());

        foreach ($this->autoWidthSampleBuffer as $bufferedRowXml) {
            $this->rowBuffer .= $bufferedRowXml;
            $this->rowBufferCount++;
        }
        $this->autoWidthSampleBuffer = [];
        $this->autoWidthSampleBufferBytes = 0;

        if ($this->rowBufferCount >= $this->bufferFlushInterval) {
            $this->flushRowBuffer();
        }
    }

    /**
     * Lenient-mode fallback path. Disables sample mode, emits the
     * preamble using whatever widths the heuristic produces, and drains
     * any rows already buffered. Triggered by trackSampledRow() when
     * width recording fails and strict mode is off.
     */
    protected function bailFromSampleMode(): void
    {
        $this->autoWidthSampleSize = null;
        $this->autoWidthFinalized = true;
        $this->inSampleMode = false;

        $this->writeSheetData($this->buildSheetPreambleXml());

        foreach ($this->autoWidthSampleBuffer as $bufferedRowXml) {
            $this->rowBuffer .= $bufferedRowXml;
            $this->rowBufferCount++;
        }
        $this->autoWidthSampleBuffer = [];
        $this->autoWidthSampleBufferBytes = 0;

        if ($this->rowBufferCount >= $this->bufferFlushInterval) {
            $this->flushRowBuffer();
        }
    }

    /**
     * Write multiple rows efficiently.
     *
     * Accepts any iterable: array, Generator, Iterator, IteratorAggregate.
     * Streaming-friendly with lazy collections (e.g. Eloquent's `lazy()` cursor):
     *
     *   $writer->writeRows(User::query()->lazy(1000));
     */
    public function writeRows(iterable $rows, ?callable $variantFor = null): void
    {
        if ($variantFor === null) {
            foreach ($rows as $row) {
                $this->writeRow($row);
            }

            return;
        }

        // Template mode: the caller decides which sample row each row should
        // look like. The ordinal is passed alongside the row so zebra
        // striping needs no counter of its own.
        $ordinal = 0;
        foreach ($rows as $row) {
            $this->writeRow($row, null, (int) $variantFor($row, $ordinal++));
        }
    }

    /**
     * Start a new named sheet, optionally with its own header row.
     *
     *   $writer->startFile(['ID', 'Name']);
     *   $writer->writeRow([1, 'Alice']);
     *   $writer->newSheet('Orders', ['ID', 'Customer', 'Total']);
     *   $writer->writeRow([100, 1, 49.90]);
     *
     * The current sheet is finalized first (if it has any rows). The new
     * sheet is created eagerly so it shows up in the output even if no
     * data rows follow.
     *
     * If $headers is null the previous header row is reused.
     */
    public function newSheet(string $name, ?array $headers = null): self
    {
        $this->refuseInTemplateMode('newSheet()', 'sheets come from the template; move to the next one with sheet($name, $dataStartRow)');
        if (! $this->started) {
            throw XlsxStreamException::headersNotSet();
        }
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        if ($name === '') {
            throw new XlsxStreamException('Sheet name cannot be empty.');
        }

        // Finalize the current sheet so its data descriptor and central-dir
        // entry are committed before we start writing the next sheet.
        if ($this->currentSheetRow > 0) {
            $this->flushRowBuffer();
            $this->finishCurrentSheet();
        }

        if ($headers !== null) {
            if (count($headers) > self::MAX_COLUMNS) {
                throw XlsxStreamException::tooManyColumns(count($headers), self::MAX_COLUMNS);
            }
            $this->columns = $headers;
        }

        $this->nextSheetName = $name;
        $this->currentSheetIndex++;
        $this->startNewSheet();

        return $this;
    }

    /**
     * Flush row buffer to stream. When the random-access index is enabled
     * and the cumulative rows-since-last-sync threshold is reached,
     * ZLIB_FULL_FLUSH is attached to the deflate_add call carrying the
     * buffer — which produces a byte-aligned 0x00 0x00 0xFF 0xFF marker
     * a downstream reader can resume inflation from with a fresh
     * inflate_init context (no inflatePrime needed).
     *
     * Critical nuance: ZLIB_FULL_FLUSH MUST be passed alongside real
     * input on the same deflate_add call. PHP's zlib does not drain the
     * encoder's pending output if the flush flag is on a separate empty
     * deflate_add('', ZLIB_FULL_FLUSH) — the flush still happens but the
     * bytes only escape on the next input. Tying it to the row-buffer
     * call keeps every emitted sync marker between two complete <row>
     * elements, the alignment invariant the reader's row tokenizer
     * depends on.
     */
    protected function flushRowBuffer(bool $forceSync = false): void
    {
        if ($this->rowBuffer === '') {
            return;
        }

        $rowsInBuffer = $this->rowBufferCount;
        $shouldSync = $this->randomAccessIndexEnabled
            && ($forceSync || ($this->rowsSinceSync + $rowsInBuffer) >= $this->indexSyncPeriod);

        if (! $shouldSync) {
            $this->writeSheetData($this->rowBuffer);
            if ($this->randomAccessIndexEnabled) {
                $this->rowsSinceSync += $rowsInBuffer;
            }
            $this->rowBuffer = '';
            $this->rowBufferCount = 0;

            return;
        }

        // Sync path: replicate writeSheetData but with ZLIB_FULL_FLUSH so
        // the deflate stream has a byte-aligned resume marker right after
        // the last row in this buffer.
        hash_update($this->crcContext, $this->rowBuffer);
        $this->sheetUncompressedSize += strlen($this->rowBuffer);

        // Running CRC of the uncompressed prefix this sync point pins —
        // exactly the first uncomp_offset bytes. hash_copy leaves the
        // live context untouched (finalizing it would kill the sheet
        // CRC); hexdec mirrors how finishCurrentSheet derives sheetCrc.
        // Feeds the KXSI "SCRC" TLV section. Cost: ~70 ns per sync
        // point, ~0.4 % of the hash_update it rides along with.
        $runningCrc = hexdec(hash_final(hash_copy($this->crcContext)));

        $compressed = deflate_add($this->deflateCtx, $this->rowBuffer, ZLIB_FULL_FLUSH);
        if ($compressed !== false && strlen($compressed) > 0) {
            $this->writeToDest($compressed);
            $this->sheetCompressedSize += strlen($compressed);
        }

        // Sync point points at the FIRST row of the NEXT batch — that is
        // the row a reader will encounter after seeking to comp_offset
        // and starting a fresh inflate context.
        $entry = $this->currentSheetEntry();
        $this->indexSyncPoints[$entry][] = [
            'row' => $this->currentSheetRow + 1,
            'comp_offset' => $this->sheetCompressedSize,
            'uncomp_offset' => $this->sheetUncompressedSize,
        ];
        $this->indexSyncPointCrcs[$entry][] = $runningCrc;

        // The rows flushed above complete an index block; snapshot the
        // column-stat accumulators so block k always spans exactly the
        // rows between sync points k-1 and k.
        if ($this->statsColumns !== [] || $this->stringStatsColumns !== []) {
            $this->closeStatsBlock($entry);
        }

        // Snap the superblock to this sync boundary once it has spanned at
        // least SUPERBLOCK_ROWS rows (syncPeriod > SUPERBLOCK_ROWS ⇒ one
        // block per superblock — the natural degenerate case).
        if ($this->rangeQuantileColumns !== [] && $this->rangeSuperblockRows >= self::SUPERBLOCK_ROWS) {
            $this->closeRangeSuperblock($entry, $this->currentSheetRow);
        }

        $this->rowsSinceSync = 0;
        $this->rowBuffer = '';
        $this->rowBufferCount = 0;
    }

    /**
     * The zip entry the sheet currently being written lives in.
     *
     * Classic sheets are numbered by the writer, so the path follows the
     * sheet index. A template sheet keeps the entry the template gave it,
     * which need not match the order it is streamed in — index sections are
     * keyed by entry, so getting this wrong would file a sheet's sync points
     * and zone maps under a sheet the reader never looks at.
     */
    protected function currentSheetEntry(): string
    {
        return $this->templateSheetEntry ?? "xl/worksheets/sheet{$this->currentSheetIndex}.xml";
    }

    /**
     * Fold one row's tracked-column values into the current block's
     * accumulators and (for data rows) the per-sheet sortedness tracker.
     *
     * Inclusion rule mirrors what a reader will see as a numeric cell —
     * over-inclusion is deliberate: a value that widens min/max can only
     * make block pruning less selective, never incorrect, whereas a
     * value the stats missed could cause a matching block to be skipped.
     *
     * $trackOrder is false for the header row: it participates in the
     * block stats (a numeric-looking header is matchable by rowsWhere's
     * full-scan path, so pruning must account for it) but says nothing
     * about the DATA ordering the sorted flag describes.
     */
    protected function accumulateColumnStats(array $row, bool $trackOrder = true): void
    {
        if (! array_is_list($row)) {
            $row = array_values($row);
        }

        // Row number for any argmin/argmax pointer set in this call. Data
        // rows carry their sheet row in currentSheetRow (≥ 2); the header
        // folds in with currentSheetRow == 0, so it claims its true row 1.
        $argRow = $this->currentSheetRow ?: 1;

        foreach ($this->statsColumns as $col) {
            $v = $this->statNumericValue($row[$col - 1] ?? null);
            $acc = &$this->statsAccum[$col];

            if ($v === null) {
                $acc['other']++;
                continue;
            }

            $trackArg = isset($this->argPointerSet[$col]);
            if ($acc['count'] === 0) {
                $acc['min'] = $v;
                $acc['max'] = $v;
                if ($trackArg) {
                    $this->argAccum[$col] = ['minRow' => $argRow, 'maxRow' => $argRow];
                }
            } else {
                if ($v < $acc['min']) {
                    $acc['min'] = $v;
                    if ($trackArg) {
                        $this->argAccum[$col]['minRow'] = $argRow;
                    }
                }
                if ($v > $acc['max']) {
                    $acc['max'] = $v;
                    if ($trackArg) {
                        $this->argAccum[$col]['maxRow'] = $argRow;
                    }
                }
            }
            $acc['sum'] += $v;
            $acc['count']++;

            if (! $trackOrder) {
                continue;
            }

            $s = &$this->statsSorted[$col];
            if ($s['prev'] !== null) {
                if ($v < $s['prev']) {
                    $s['asc'] = false;
                }
                if ($v > $s['prev']) {
                    $s['desc'] = false;
                }
            }
            $s['prev'] = $v;
        }
    }

    /**
     * Stage one data row's tracked-column values for the current sheet's
     * sketches. Numeric interpretation for the t-digest is EXACTLY the
     * STAT rule (statNumericValue), so quantiles describe the same value
     * population columnStats() aggregates — except non-finite floats,
     * which would poison every centroid mean and are skipped (STAT's
     * min/max tolerate them; a t-digest cannot). The HLL hashes the
     * canonical string of every non-empty value, numeric or not.
     *
     * Values are appended to per-column buffers and bulk-fed to the
     * sketches every SKETCH_FLUSH_ROWS rows (flushSketchBuffers) — the
     * common string/int/float cases are dispatched inline because the
     * helper-method calls were the measured hot cost, not the sketch
     * math; rare types (bool/DateTime/objects) take the readable path.
     *
     * Called only from writeRow() — never for the header row, which the
     * preamble emits. That exclusion is the documented TDIG/CHLL
     * semantic, not an accident of plumbing (see withColumnSketches).
     */
    protected function accumulateColumnSketches(array $row): void
    {
        if (! array_is_list($row)) {
            $row = array_values($row);
        }

        foreach ($this->sketchColumns as $col) {
            $value = $row[$col - 1] ?? null;

            if ($value === null) {
                continue;
            }
            if (is_string($value)) {
                if ($value === '') {
                    continue;
                }
                $this->sketchStrBuffer[$col][] = $value;
                if (is_numeric($value)) {
                    $v = (float) $value;
                    if (is_finite($v)) { // '1e999' is numeric but overflows to INF
                        $this->sketchNumBuffer[$col][] = $v;
                    }
                }
                continue;
            }
            if (is_int($value)) {
                $this->sketchNumBuffer[$col][] = (float) $value;
                $this->sketchStrBuffer[$col][] = (string) $value;
                continue;
            }
            if (is_float($value)) {
                if (is_finite($value)) {
                    $this->sketchNumBuffer[$col][] = $value;
                }
                $this->sketchStrBuffer[$col][] = (string) $value;
                continue;
            }

            // Rare types: bool, DateTime, Stringable objects.
            $v = $this->statNumericValue($value);
            if ($v !== null && is_finite($v)) {
                $this->sketchNumBuffer[$col][] = $v;
            }
            $canonical = $this->sketchCanonicalString($value);
            if ($canonical !== null) {
                $this->sketchStrBuffer[$col][] = $canonical;
            }
        }

        if (++$this->sketchRowsBuffered >= self::SKETCH_FLUSH_ROWS) {
            $this->flushSketchBuffers();
        }
    }

    /**
     * Fold one DATA row into every tracked column pair's co-moments (CORR).
     * Each column's numeric value is resolved once per row (a column sits in
     * up to k−1 pairs), then a pair accumulates only when BOTH sides are
     * numeric and finite — the aligned-observation rule Pearson requires.
     */
    protected function accumulateCorrelations(array $row): void
    {
        if (! array_is_list($row)) {
            $row = array_values($row);
        }

        $vals = [];
        foreach ($this->correlationColumns as $col) {
            $v = $this->statNumericValue($row[$col - 1] ?? null);
            $vals[$col] = ($v !== null && is_finite($v)) ? $v : null;
        }

        foreach ($this->correlationPairs as [$a, $b]) {
            if ($vals[$a] !== null && $vals[$b] !== null) {
                $this->coMomentAccum[$a.','.$b]->add($vals[$a], $vals[$b]);
            }
        }
    }

    /**
     * Fold one data row's tracked numeric values into the current open
     * superblock digest (TDGB). One shared row counter drives the
     * sync-snapped superblock close in flushRowBuffer.
     */
    protected function accumulateRangeQuantiles(array $row): void
    {
        if (! array_is_list($row)) {
            $row = array_values($row);
        }
        foreach ($this->rangeQuantileColumns as $col) {
            $v = $this->statNumericValue($row[$col - 1] ?? null);
            if ($v !== null && is_finite($v)) {
                $this->rangeDigestAccum[$col]->add($v);
            }
        }
        $this->rangeSuperblockRows++;
    }

    /**
     * Fold one data row's tracked values into their Misra-Gries sketches
     * (TOPK), keyed by the same canonical string form the HyperLogLog uses
     * (§4.4). Empty cells (null / '') are not values and are skipped, so a
     * column of mostly-empty cells does not spend a counter on ''.
     */
    protected function accumulateTopValues(array $row): void
    {
        if (! array_is_list($row)) {
            $row = array_values($row);
        }
        foreach ($this->topValuesColumns as $col) {
            $canonical = $this->sketchCanonicalString($row[$col - 1] ?? null);
            if ($canonical !== null) {
                $this->topValueSketches[$col]->add($canonical);
            }
        }
    }

    /**
     * Close the current open superblock at a sync point: serialize each
     * column's digest with the superblock's last row, then reset. Called
     * from flushRowBuffer once ≥ SUPERBLOCK_ROWS have accumulated (snap to
     * the sync boundary), and once at sheet finalize for the remainder.
     */
    protected function closeRangeSuperblock(string $entry, int $endRow): void
    {
        foreach ($this->rangeQuantileColumns as $col) {
            $this->indexRangeSuperblocks[$entry][$col][] = [
                'end_row' => $endRow,
                'payload' => $this->rangeDigestAccum[$col]->serialize(),
            ];
            $this->rangeDigestAccum[$col] = new TDigest();
        }
        $this->rangeSuperblockRows = 0;
    }

    /**
     * Drain the per-column staging buffers into the sheet's sketches.
     * Runs every SKETCH_FLUSH_ROWS rows and right before the sheet's
     * sketches are serialized (finishCurrentSheet).
     */
    protected function flushSketchBuffers(): void
    {
        foreach ($this->sketchColumns as $col) {
            if (($this->sketchNumBuffer[$col] ?? []) !== []) {
                $this->sketchDigestAccum[$col]->addMany($this->sketchNumBuffer[$col]);
                $this->sketchNumBuffer[$col] = [];
            }
            if (($this->sketchStrBuffer[$col] ?? []) !== []) {
                $this->sketchHllAccum[$col]->addMany($this->sketchStrBuffer[$col]);
                $this->sketchStrBuffer[$col] = [];
            }
        }
        $this->sketchRowsBuffered = 0;
    }

    /**
     * Canonical string a cell value is hashed under for distinct
     * counting (the KXSI "CHLL" canonicalization rule, SPEC.md §4.4):
     *
     *   - null and '' → excluded entirely (an empty cell is the absence
     *     of a value, not a distinct value);
     *   - strings → as-is, byte-exact, before any XML escaping;
     *   - int/float → PHP's decimal rendering `(string) $value` — the
     *     same text buildRowXml writes into the cell's <v>;
     *   - bool → '1' / '0' (matches the <v> of a t="b" cell);
     *   - DateTimeInterface → decimal rendering of its Excel serial
     *     (matches the <v> of the date cell);
     *   - other objects → `(string) $value` (the inlineStr rendering).
     *
     * countDistinct is therefore over canonical FORMS, not typed
     * identity: int 7 and the string '7' collapse (both render '7'),
     * while '1.50' stays distinct from 1.5 (strings hash as-is even
     * though the number cell renders '1.5').
     */
    protected function sketchCanonicalString(mixed $value): ?string
    {
        if ($value === null) {
            return null;
        }
        if (is_string($value)) {
            return $value === '' ? null : $value;
        }
        if (is_int($value) || is_float($value)) {
            return (string) $value;
        }
        if (is_bool($value)) {
            return $value ? '1' : '0';
        }
        if ($value instanceof \DateTimeInterface) {
            return (string) (($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400);
        }

        return (string) $value;
    }

    /**
     * Numeric interpretation of a cell for statistics purposes, aligned
     * with how buildRowXml renders the value (DateTime -> Excel serial,
     * bool -> 0/1, numeric strings -> their numeric value). Null means
     * "not numeric" and lands in the block's `other` count.
     */
    protected function statNumericValue(mixed $value): ?float
    {
        if (is_int($value) || is_float($value)) {
            return (float) $value;
        }
        if (is_string($value)) {
            return is_numeric($value) ? (float) $value : null;
        }
        if ($value instanceof \DateTimeInterface) {
            return ($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400;
        }
        if (is_bool($value)) {
            return $value ? 1.0 : 0.0;
        }

        return null;
    }

    /**
     * String analogue of accumulateColumnStats: fold each tracked cell's
     * canonical string into the current block's lexicographic [min, max]
     * (unsigned byte compare = strcmp). Empty/null cells count as `other`
     * (mirrors the sketch canonical rule), so a string predicate that can
     * never match them is not widened by them. $trackOrder = false for the
     * header fold, exactly like the numeric path.
     */
    protected function accumulateStringStats(array $row, bool $trackOrder = true): void
    {
        if (! array_is_list($row)) {
            $row = array_values($row);
        }

        foreach ($this->stringStatsColumns as $col) {
            $s = $this->sketchCanonicalString($row[$col - 1] ?? null);
            $acc = &$this->stringAccum[$col];

            if ($s === null) {
                $acc['other']++;
                unset($acc);

                continue;
            }

            if ($acc['count'] === 0) {
                $acc['min'] = $s;
                $acc['max'] = $s;
            } else {
                if (strcmp($s, $acc['min']) < 0) {
                    $acc['min'] = $s;
                }
                if (strcmp($s, $acc['max']) > 0) {
                    $acc['max'] = $s;
                }
            }
            $acc['count']++;
            unset($acc);

            if (! $trackOrder) {
                continue;
            }

            $so = &$this->stringSorted[$col];
            if ($so['prev'] !== null) {
                if (strcmp($s, $so['prev']) < 0) {
                    $so['asc'] = false;
                }
                if (strcmp($s, $so['prev']) > 0) {
                    $so['desc'] = false;
                }
            }
            $so['prev'] = $s;
            unset($so);
        }
    }

    /** Byte-wise common prefix length of two strings. */
    protected static function commonPrefixLen(string $a, string $b): int
    {
        $n = min(strlen($a), strlen($b));
        $i = 0;
        while ($i < $n && $a[$i] === $b[$i]) {
            $i++;
        }

        return $i;
    }

    /**
     * Deferred shortest-separator truncation length for one column: one
     * byte past the DEEPEST divergence between adjacent boundary strings
     * (the classic B-tree separator), so every block boundary stays
     * distinguishable after truncation — restoring pruning on a
     * common-prefix corpus, capped at STRING_STAT_CAP.
     *
     * The adjacent-pair (sorted) measure is deliberately robust to
     * outliers: a lone off-prefix boundary (e.g. a differently-cased
     * header folded into block 0) does NOT drag the length down the way a
     * whole-set common prefix would. When two boundaries are identical
     * (a key straddling a block edge) no length separates them; the cap
     * bounds the cost and pruning honestly weakens at that one edge.
     *
     * @param  list<string>  $boundaries  every block's stored min and max
     */
    protected static function stringTruncationLength(array $boundaries): int
    {
        if (count($boundaries) <= 1) {
            return self::STRING_STAT_CAP;
        }
        sort($boundaries, SORT_STRING); // byte order = strcmp = the STRZ collation

        $maxAdjacent = 0;
        for ($i = 1, $n = count($boundaries); $i < $n; $i++) {
            $lcp = self::commonPrefixLen($boundaries[$i - 1], $boundaries[$i]);
            if ($lcp > $maxAdjacent) {
                $maxAdjacent = $lcp;
            }
        }

        // +4 past the deepest adjacent divergence: enough to separate every
        // boundary plus a little slack for query values sharing the prefix.
        return max(1, min(self::STRING_STAT_CAP, $maxAdjacent + 4));
    }

    /** Sound LOWER bound: the first $len bytes are a prefix, always ≤ $s. */
    protected static function truncateStringMin(string $s, int $len): string
    {
        return substr($s, 0, $len);
    }

    /**
     * Sound UPPER bound: the smallest string ≥ every string sharing $s's
     * $len-byte prefix. Increment the rightmost non-0xFF byte of the
     * prefix and drop the tail; if $s already fits in $len it is exact; an
     * all-0xFF prefix cannot be bumped, so $s is kept whole.
     */
    protected static function truncateStringMax(string $s, int $len): string
    {
        if (strlen($s) <= $len) {
            return $s;
        }
        $p = substr($s, 0, $len);
        for ($i = $len - 1; $i >= 0; $i--) {
            $b = ord($p[$i]);
            if ($b < 0xFF) {
                return substr($p, 0, $i).chr($b + 1);
            }
        }

        return $s;
    }

    /**
     * Snapshot the per-column accumulators as one closed index block for
     * the sheet and reset them for the next block.
     */
    protected function closeStatsBlock(string $entry): void
    {
        foreach ($this->statsColumns as $col) {
            $this->indexColumnBlocks[$entry][$col][] = $this->statsAccum[$col];
            $this->statsAccum[$col] = ['min' => 0.0, 'max' => 0.0, 'sum' => 0.0, 'count' => 0, 'other' => 0];
        }

        // Argmin/argmax row pointers (ARGP), 1:1 with the STAT blocks just
        // pushed. A block with no numeric values keeps {0, 0} — row 0 does
        // not exist, so it reads as "no pointer".
        foreach ($this->argPointerColumns as $col) {
            $this->indexArgPointers[$entry][$col][] = $this->argAccum[$col] ?? ['minRow' => 0, 'maxRow' => 0];
            $this->argAccum[$col] = ['minRow' => 0, 'maxRow' => 0];
        }

        // String blocks: cap the full min/max to STRING_STAT_CAP now (sound
        // bounds — min truncates down, max up); the shorter per-column
        // separator length is chosen later at finishFile.
        foreach ($this->stringStatsColumns as $col) {
            $acc = $this->stringAccum[$col];
            $this->indexStringBlocks[$entry][$col][] = [
                'min' => $acc['min'] === null ? null : self::truncateStringMin($acc['min'], self::STRING_STAT_CAP),
                'max' => $acc['max'] === null ? null : self::truncateStringMax($acc['max'], self::STRING_STAT_CAP),
                'count' => $acc['count'],
                'other' => $acc['other'],
            ];
            $this->stringAccum[$col] = ['min' => null, 'max' => null, 'count' => 0, 'other' => 0];
        }
    }

    /**
     * Serialize the per-sheet sync points into the binary sidecar payload.
     * Sheets are written in workbook order so the sidecar's section order
     * matches the workbook.xml sheet listing.
     */
    protected function buildRandomAccessIndexPayload(): string
    {
        $sheetSections = [];
        $columnStats = [];
        $columnStringStats = [];
        $syncPointCrcs = [];
        foreach ($this->sheets as $sheet) {
            $entry = $sheet['filename'];
            $sheetSections[] = [
                'entry' => $entry,
                'total_rows' => $sheet['rows'],
                'sheet_crc32' => $sheet['crc32'] ?? 0,
                'sync_points' => $this->indexSyncPoints[$entry] ?? [],
            ];
            // Always present when the index is enabled — a sheet without
            // sync points contributes an empty (count = 0) SCRC record so
            // the section stays aligned with the core body.
            $syncPointCrcs[$entry] = $this->indexSyncPointCrcs[$entry] ?? [];

            if ($this->statsColumns !== []) {
                $cols = [];
                foreach ($this->statsColumns as $col) {
                    $sorted = $this->indexColumnSorted[$entry][$col] ?? ['asc' => false, 'desc' => false];
                    $cols[] = [
                        'col' => $col,
                        'sorted_asc' => $sorted['asc'],
                        'sorted_desc' => $sorted['desc'],
                        'blocks' => $this->indexColumnBlocks[$entry][$col] ?? [],
                    ];
                }
                $columnStats[$entry] = $cols;
            }

            if ($this->stringStatsColumns !== []) {
                $cols = [];
                foreach ($this->stringStatsColumns as $col) {
                    $blocks = $this->indexStringBlocks[$entry][$col] ?? [];

                    // Deferred separator truncation: choose one length for
                    // the whole column from its block boundaries, then
                    // truncate every block's min (down) / max (up) to it.
                    $boundaries = [];
                    foreach ($blocks as $b) {
                        if ($b['min'] !== null) {
                            $boundaries[] = $b['min'];
                        }
                        if ($b['max'] !== null) {
                            $boundaries[] = $b['max'];
                        }
                    }
                    $len = self::stringTruncationLength($boundaries);
                    foreach ($blocks as &$b) {
                        if ($b['min'] !== null) {
                            $b['min'] = self::truncateStringMin($b['min'], $len);
                        }
                        if ($b['max'] !== null) {
                            $b['max'] = self::truncateStringMax($b['max'], $len);
                        }
                    }
                    unset($b);

                    $sorted = $this->indexStringSorted[$entry][$col] ?? ['asc' => false, 'desc' => false];
                    $cols[] = [
                        'col' => $col,
                        'sorted_asc' => $sorted['asc'],
                        'sorted_desc' => $sorted['desc'],
                        'blocks' => $blocks,
                    ];
                }
                $columnStringStats[$entry] = $cols;
            }
        }

        return RandomAccessIndex::encode(
            $this->indexSyncPeriod,
            $sheetSections,
            $columnStats,
            $syncPointCrcs,
            $this->indexColumnDigests,
            $this->indexColumnHlls,
            $columnStringStats,
            $this->indexRangeSuperblocks,
            $this->indexTopValues,
            $this->indexArgPointers,
            $this->indexCorrelations
        );
    }

    /**
     * Sanitize sheet name for Excel compatibility
     */
    protected function sanitizeSheetName(string $name): string
    {
        $name = preg_replace('/[:*?\/\\\[\]]/', '_', $name);

        if (mb_strlen($name, 'UTF-8') > 31) {
            $name = mb_substr($name, 0, 31, 'UTF-8');
        }

        return $name === '' ? 'Sheet' : $name;
    }

    /**
     * Start a new sheet
     */
    protected function startNewSheet(): void
    {
        // Validate any column formats / widths against the now-known column
        // count. Catches setColumnFormat(99, ...) typos at the point where
        // the columns are actually committed for the sheet.
        $columnCount = count($this->columns);
        if ($columnCount > 0) {
            foreach (array_keys($this->columnStyleIds) as $col) {
                if ($col > $columnCount) {
                    throw XlsxStreamException::columnIndexOutOfRange($col, $columnCount);
                }
            }
        }

        // Custom name (from newSheet()) wins, else preserve the legacy default
        // ("Report" for the first sheet, "SheetN" for auto-split overflow).
        if ($this->nextSheetName !== null) {
            $sheetName = $this->nextSheetName;
            $this->nextSheetName = null;
        } elseif ($this->currentSheetIndex === 1) {
            $sheetName = 'Report';
        } else {
            $sheetName = "Sheet{$this->currentSheetIndex}";
        }

        $sheetName = $this->sanitizeSheetName($sheetName);
        $filename = "xl/worksheets/sheet{$this->currentSheetIndex}.xml";

        $this->sheets[] = [
            'index' => $this->currentSheetIndex,
            'name' => $sheetName,
            'filename' => $filename,
            'rows' => 0,
        ];

        $this->sheetOffset = $this->currentOffset;
        $this->currentSheetRow = 0;

        $this->resetPerSheetAccumulators();

        $this->openSheetStream($filename);

        // Reset per-sheet sample state (multi-sheet workbooks re-sample
        // because each sheet may have different column widths). Clear
        // any widths the previous sheet's sample wrote into columnWidths
        // so this sheet starts from a clean slate — user-explicit widths
        // (set via setColumnWidths) survive intentionally.
        foreach ($this->sampleAutoSetWidthCols as $col) {
            unset($this->columnWidths[$col]);
        }
        $this->sampleAutoSetWidthCols = [];
        $this->autoWidthSampleBuffer = [];
        $this->autoWidthSampleBufferBytes = 0;
        $this->autoWidthMaxLengths = [];
        $this->autoWidthFinalized = false;

        $sampleMode = $this->autoColumnWidth
            && $this->autoWidthSampleSize !== null
            && $this->autoWidthSampleSize > 0;

        if ($sampleMode) {
            // Sample mode: defer preamble + header emission until we know
            // the per-column widths. Seed the width tracker with the
            // header lengths so a long header still influences the result.
            $this->inSampleMode = true;
            if (! empty($this->columns)) {
                foreach ($this->columns as $i => $headerCell) {
                    $col = $i + 1;
                    $len = mb_strlen((string) $headerCell);
                    $this->autoWidthMaxLengths[$col] = $len;
                }
                $this->currentSheetRow = 1; // claim row 1 for the eventual header
            }

            return;
        }

        // Normal path: emit preamble + header right away.
        $this->inSampleMode = false;
        $this->writeSheetData($this->buildSheetPreambleXml());
        if (! empty($this->columns)) {
            $this->currentSheetRow = 1;
        }
    }

    /**
     * Reset every per-sheet accumulator so a freshly started sheet carries
     * its own verdicts. Shared by the classic and the template sheet-open
     * paths; the header fold below is guarded on $this->columns, which
     * template mode leaves empty because the template writes its own header.
     */
    protected function resetPerSheetAccumulators(): void
    {
        // Fresh sheet -> fresh stat accumulators and sortedness trackers
        // (sortedness is a per-sheet property; auto-split sheets each get
        // their own verdict).
        if ($this->statsColumns !== []) {
            foreach ($this->statsColumns as $col) {
                $this->statsAccum[$col] = ['min' => 0.0, 'max' => 0.0, 'sum' => 0.0, 'count' => 0, 'other' => 0];
                $this->statsSorted[$col] = ['asc' => true, 'desc' => true, 'prev' => null];
            }

            // The header row is emitted via the sheet preamble, never
            // through writeRow() — but it IS a row the reader's
            // full-scan path can match (a numeric-looking header passes
            // is_numeric). Fold it into block 0 so zone-map pruning can
            // never hide a row the un-pruned path would return; without
            // this, rowsWhere() gave different results with and without
            // stats for out-of-data-range values matching the header.
            // Fresh block-0 arg accumulators before the header folds in
            // (so a numeric header can claim row 1 as its extreme).
            foreach ($this->argPointerColumns as $col) {
                $this->argAccum[$col] = ['minRow' => 0, 'maxRow' => 0];
            }
            if ($this->columns !== []) {
                $this->accumulateColumnStats($this->columns, trackOrder: false);
            }
        }

        // Same fresh-sheet reset + header fold for string zone maps.
        if ($this->stringStatsColumns !== []) {
            foreach ($this->stringStatsColumns as $col) {
                $this->stringAccum[$col] = ['min' => null, 'max' => null, 'count' => 0, 'other' => 0];
                $this->stringSorted[$col] = ['asc' => true, 'desc' => true, 'prev' => null];
            }
            if ($this->columns !== []) {
                $this->accumulateStringStats($this->columns, trackOrder: false);
            }
        }

        // Fresh sheet -> fresh sketches (auto-split sheets each carry
        // their own — the sections are per-sheet and merge-friendly).
        // The header row is deliberately NOT folded in here: sketches
        // estimate the data distribution, they prune nothing, so the
        // STAT header-fold soundness argument does not apply and folding
        // would only bias the estimates (see withColumnSketches).
        if ($this->sketchColumns !== []) {
            foreach ($this->sketchColumns as $col) {
                $this->sketchDigestAccum[$col] = new TDigest();
                $this->sketchHllAccum[$col] = new HyperLogLog();
                $this->sketchNumBuffer[$col] = [];
                $this->sketchStrBuffer[$col] = [];
            }
            $this->sketchRowsBuffered = 0;
        }

        // Fresh sheet -> fresh open superblock digests (per-sheet, like the
        // sketches). No header fold: quantiles estimate the data.
        if ($this->rangeQuantileColumns !== []) {
            foreach ($this->rangeQuantileColumns as $col) {
                $this->rangeDigestAccum[$col] = new TDigest();
            }
            $this->rangeSuperblockRows = 0;
        }

        // Fresh sheet -> fresh Misra-Gries sketches (per-sheet, mergeable;
        // header not folded, same reasoning as the other sketches).
        if ($this->topValuesColumns !== []) {
            foreach ($this->topValuesColumns as $col) {
                $this->topValueSketches[$col] = new MisraGries($this->topValuesK);
            }
        }

        // Fresh sheet -> fresh co-moment accumulators (per-sheet, mergeable).
        if ($this->correlationPairs !== []) {
            foreach ($this->correlationPairs as [$a, $b]) {
                $this->coMomentAccum[$a.','.$b] = new CoMoments();
            }
        }
    }

    /**
     * Open a sheet's ZIP entry: a local file header whose sizes are deferred
     * to a data descriptor, then a fresh deflate + CRC context and empty row
     * buffers. These are the bytes the golden-file tests pin.
     */
    protected function openSheetStream(string $filename): void
    {
        [$mtime, $mdate] = $this->dosTimeParts(time());

        // Write local file header
        $header = pack('V', self::LOCAL_FILE_HEADER_SIGNATURE);
        $header .= pack('v', self::VERSION_NEEDED);
        $header .= pack('v', 0x0008);
        $header .= pack('v', self::COMPRESSION_DEFLATED);
        $header .= pack('v', $mtime);
        $header .= pack('v', $mdate);
        $header .= pack('V', 0);
        $header .= pack('V', 0);
        $header .= pack('V', 0);
        $header .= pack('v', strlen($filename));
        $header .= pack('v', 0);
        $header .= $filename;

        $this->writeToDest($header);

        // Initialize deflate context
        $this->deflateCtx = deflate_init(ZLIB_ENCODING_RAW, ['level' => $this->deflateLevel]);
        $this->crcContext = hash_init('crc32b');
        $this->sheetCrc = 0;
        $this->sheetUncompressedSize = 0;
        $this->sheetCompressedSize = 0;
        $this->rowBuffer = '';
        $this->rowBufferCount = 0;
        $this->rowsSinceSync = 0;
    }

    /**
     * Build the deferred preamble (xml decl + worksheet open + sheetViews
     * + cols + sheetData open + header row when present). Used both by
     * the immediate-emission path in startNewSheet() and by the
     * sample-mode finalize/bail paths.
     */
    protected function buildSheetPreambleXml(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $xml .= '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        $xml .= $this->buildSheetViewsXml();
        $xml .= $this->buildColsXml();
        $xml .= '<sheetData>';

        if (! empty($this->columns)) {
            $headerStyleAttr = $this->headerStyleId !== null ? ' s="'.$this->headerStyleId.'"' : '';
            // Compact mode drops the optional r attributes here too —
            // once per sheet, so a plain conditional (not a twin) does.
            if ($this->compactMode) {
                $headerRow = '<row>';
                foreach ($this->columns as $header) {
                    $escaped = $this->fastXmlEscape((string) $header);
                    $headerRow .= '<c'.$headerStyleAttr.' t="inlineStr"><is><t>'.$escaped.'</t></is></c>';
                }
            } else {
                $headerRow = '<row r="1">';
                foreach ($this->columns as $i => $header) {
                    $cellRef = $this->getColumnLetter($i + 1).'1';
                    $escaped = $this->fastXmlEscape((string) $header);
                    $headerRow .= '<c r="'.$cellRef.'"'.$headerStyleAttr.' t="inlineStr"><is><t>'.$escaped.'</t></is></c>';
                }
            }
            $headerRow .= '</row>';
            $xml .= $headerRow;
        }

        return $xml;
    }

    /**
     * Write data to current sheet with streaming compression
     */
    protected function writeSheetData(string $data): void
    {
        hash_update($this->crcContext, $data);
        $this->sheetUncompressedSize += strlen($data);

        $compressed = deflate_add($this->deflateCtx, $data, ZLIB_NO_FLUSH);
        if ($compressed !== false && strlen($compressed) > 0) {
            $this->writeToDest($compressed);
            $this->sheetCompressedSize += strlen($compressed);
        }
    }

    /**
     * Build the <sheetViews> block when freeze panes are configured.
     * Must appear BEFORE <sheetData> per OOXML schema order.
     */
    protected function buildSheetViewsXml(): string
    {
        if ($this->freezeRows === 0 && $this->freezeColumns === 0) {
            return '';
        }

        $topLeftCol = $this->getColumnLetter($this->freezeColumns + 1);
        $topLeftRow = $this->freezeRows + 1;
        $topLeftCell = $topLeftCol.$topLeftRow;

        $activePane = match (true) {
            $this->freezeRows > 0 && $this->freezeColumns > 0 => 'bottomRight',
            $this->freezeRows > 0 => 'bottomLeft',
            default => 'topRight',
        };

        $pane = '<pane';
        if ($this->freezeColumns > 0) {
            $pane .= ' xSplit="'.$this->freezeColumns.'"';
        }
        if ($this->freezeRows > 0) {
            $pane .= ' ySplit="'.$this->freezeRows.'"';
        }
        $pane .= ' topLeftCell="'.$topLeftCell.'" activePane="'.$activePane.'" state="frozen"/>';

        return '<sheetViews><sheetView workbookViewId="0">'.$pane.'</sheetView></sheetViews>';
    }

    /**
     * Build the <cols> block emitted between <sheetViews> and <sheetData>.
     *
     * Resolution order (per column):
     *   1. Explicit width from setColumnWidths()
     *   2. setAutoColumnWidth() heuristic: max(format_min, header_len + 2, 8)
     *      where format_min reflects the rendered width of values formatted
     *      as date / datetime / currency / percent / decimal — fixes the
     *      "Salary header is 6 chars but ₺50,000.00 needs ~14 chars" case
     *   3. Excel default (no <col> entry)
     */
    protected function buildColsXml(): string
    {
        $resolved = [];
        $columnCount = count($this->columns);
        $maxWidthCol = $this->columnWidths === [] ? 0 : max(array_keys($this->columnWidths));
        $upperBound = $columnCount > $maxWidthCol ? $columnCount : $maxWidthCol;

        // Plain for-loop (vs. range()) avoids allocating a 1..N array on
        // every sheet startup — matters at the 16,384 column limit.
        for ($col = 1; $col <= $upperBound; $col++) {
            if (isset($this->columnWidths[$col])) {
                $resolved[$col] = $this->columnWidths[$col];

                continue;
            }
            if ($this->autoColumnWidth && isset($this->columns[$col - 1])) {
                $headerLen = mb_strlen((string) $this->columns[$col - 1], 'UTF-8');
                $formatMin = $this->minWidthForFormat($this->columnFormatNames[$col] ?? null);
                $resolved[$col] = (float) max(8, $headerLen + 2, $formatMin);
            }
        }

        if (empty($resolved)) {
            return '';
        }

        $xml = '<cols>';
        foreach ($resolved as $col => $width) {
            $xml .= '<col min="'.$col.'" max="'.$col.'" width="'.$width.'" customWidth="1"/>';
        }
        $xml .= '</cols>';

        return $xml;
    }

    /**
     * Minimum sensible column width (in chars) for a given format preset.
     * Returns 0 for unknown / unformatted columns so the header heuristic wins.
     */
    protected function minWidthForFormat(?string $name): int
    {
        return match ($name) {
            'date' => 12,                    // 2026-01-15
            'datetime', 'datetime_iso' => 20, // 2026-01-15 10:30:00
            'time' => 10,                    // 10:30:00
            'currency_try',
            'currency_usd',
            'currency_eur',
            'currency_gbp' => 14,            // ₺1,234,567.89
            'percent' => 10,                 // 100.00%
            'decimal' => 14,                 // 1,234,567.89
            'integer' => 14,                 // 1,234,567,890 (10-digit grouped)
            'builtin:'.self::BUILTIN_NUMFMT_INTEGER => 14,       // 1
            'builtin:'.self::BUILTIN_NUMFMT_DECIMAL_2 => 14,     // 2
            'builtin:'.self::BUILTIN_NUMFMT_THOUSANDS => 14,     // 3
            'builtin:'.self::BUILTIN_NUMFMT_CURRENCY => 14,      // 5
            'builtin:'.self::BUILTIN_NUMFMT_PERCENT => 10,       // 9
            'builtin:'.self::BUILTIN_NUMFMT_PERCENT_2 => 10,     // 10
            'builtin:'.self::BUILTIN_NUMFMT_EXPONENT => 11,      // 11
            'builtin:'.self::BUILTIN_NUMFMT_FRACTION => 10,      // 12
            'builtin:'.self::BUILTIN_NUMFMT_DATE => 12,          // 14
            'builtin:'.self::BUILTIN_NUMFMT_DATE_LONG => 13,     // 15
            'builtin:'.self::BUILTIN_NUMFMT_TIME_AMPM => 11,     // 18
            'builtin:'.self::BUILTIN_NUMFMT_TIME => 10,          // 20
            'builtin:'.self::BUILTIN_NUMFMT_DATETIME => 20,      // 22
            default => 0,
        };
    }

    /**
     * Build the <autoFilter> element written AFTER </sheetData>.
     * Range covers all populated rows and the configured column count.
     */
    protected function buildAutoFilterXml(): string
    {
        if (! $this->autoFilterEnabled || empty($this->columns) || $this->currentSheetRow < 1) {
            return '';
        }

        $lastCol = $this->getColumnLetter(count($this->columns));
        $range = 'A1:'.$lastCol.$this->currentSheetRow;

        return '<autoFilter ref="'.$range.'"/>';
    }

    /**
     * Finish current sheet
     */
    protected function finishCurrentSheet(): void
    {
        // If a sample is still pending (e.g. fewer rows than the sample
        // size were ever written), finalize so the deferred preamble +
        // header still get emitted before we close the sheet.
        if ($this->inSampleMode && ! $this->autoWidthFinalized) {
            $this->finalizeAutoWidthSample();
        }

        $this->flushRowBuffer();

        // Template mode closes with the template's own tail — mergeCells,
        // conditional formatting, page setup, everything below the data —
        // carried across verbatim.
        $sheetFooter = $this->templateSheet !== null
            ? $this->templateSheet->tail()
            : '</sheetData>'.$this->buildAutoFilterXml().'</worksheet>';

        hash_update($this->crcContext, $sheetFooter);
        $this->sheetUncompressedSize += strlen($sheetFooter);

        $this->sheetCrc = hexdec(hash_final($this->crcContext));

        $compressed = deflate_add($this->deflateCtx, $sheetFooter, ZLIB_FINISH);
        if ($compressed !== false) {
            $this->writeToDest($compressed);
            $this->sheetCompressedSize += strlen($compressed);
        }

        $sheetInfo = end($this->sheets);
        $this->assertZip32Compatible($this->sheetCompressedSize, "sheet '{$sheetInfo['filename']}' compressed size");
        $this->assertZip32Compatible($this->sheetUncompressedSize, "sheet '{$sheetInfo['filename']}' uncompressed size");
        $this->assertZip32Compatible($this->currentOffset, 'cumulative archive offset at end of sheet');

        // Write data descriptor
        $descriptor = pack('V', self::DATA_DESCRIPTOR_SIGNATURE);
        $descriptor .= pack('V', $this->sheetCrc);
        $descriptor .= pack('V', $this->sheetCompressedSize);
        $descriptor .= pack('V', $this->sheetUncompressedSize);

        $this->writeToDest($descriptor);

        // Add to central directory
        $sheetInfo = end($this->sheets);
        $this->centralDirectory[] = [
            'filename' => $sheetInfo['filename'],
            'crc32' => $this->sheetCrc,
            'compressed_size' => $this->sheetCompressedSize,
            'uncompressed_size' => $this->sheetUncompressedSize,
            'offset' => $this->sheetOffset,
            'compression' => self::COMPRESSION_DEFLATED,
            'flags' => 0x0008,
            'timestamp' => time(),
        ];

        $this->sheets[count($this->sheets) - 1]['rows'] = $this->currentSheetRow;
        // Mirror the ZIP-CD CRC into the sheet record so the random-access
        // index payload can pin the sheet content. Reader compares this
        // with the live CD CRC at open time and silently invalidates the
        // index when an external editor rewrote the sheet.
        $this->sheets[count($this->sheets) - 1]['crc32'] = $this->sheetCrc;

        // Close the tail block (rows after the last sync point — possibly
        // empty, still emitted so block_count == sync_count + 1 holds for
        // every sheet) and pin the sheet's sortedness verdict.
        if ($this->statsColumns !== [] || $this->stringStatsColumns !== []) {
            $entry = $sheetInfo['filename'];
            $this->closeStatsBlock($entry);
            foreach ($this->statsColumns as $col) {
                $s = $this->statsSorted[$col];
                $this->indexColumnSorted[$entry][$col] = ['asc' => $s['asc'], 'desc' => $s['desc']];
            }
            foreach ($this->stringStatsColumns as $col) {
                $so = $this->stringSorted[$col];
                $this->indexStringSorted[$entry][$col] = ['asc' => $so['asc'], 'desc' => $so['desc']];
            }
        }

        // Snapshot the sheet's sketches in serialized form — drain the
        // row-side staging buffers first; serialize() then folds any
        // values still buffered inside the digest, so the stored payload
        // is the final canonical sketch for the sheet.
        if ($this->sketchColumns !== []) {
            $this->flushSketchBuffers();
            $entry = $sheetInfo['filename'];
            foreach ($this->sketchColumns as $col) {
                $this->indexColumnDigests[$entry][$col] = $this->sketchDigestAccum[$col]->serialize();
                $this->indexColumnHlls[$entry][$col] = $this->sketchHllAccum[$col]->serialize();
            }
        }

        // Close the final (partial) superblock — the rows since the last
        // sync-snap. Guard on the row counter so a sheet that snapped
        // exactly on its last sync point does not emit an empty trailer.
        if ($this->rangeQuantileColumns !== [] && $this->rangeSuperblockRows > 0) {
            $this->closeRangeSuperblock($sheetInfo['filename'], $this->currentSheetRow);
        }

        // Snapshot the sheet's Misra-Gries sketches in serialized form.
        if ($this->topValuesColumns !== []) {
            $entry = $sheetInfo['filename'];
            foreach ($this->topValuesColumns as $col) {
                $this->indexTopValues[$entry][$col] = $this->topValueSketches[$col]->serialize();
            }
        }

        // Snapshot the sheet's co-moment accumulators in serialized form.
        if ($this->correlationPairs !== []) {
            $entry = $sheetInfo['filename'];
            foreach ($this->correlationPairs as [$a, $b]) {
                $this->indexCorrelations[$entry][$a.','.$b] = $this->coMomentAccum[$a.','.$b]->serialize();
            }
        }

        // Retire the template sheet: its entry is written, so the copy pass
        // must skip it and a second sheet() call on the same name is refused.
        if ($this->templateSheet !== null) {
            $this->templateStreamedEntries[(string) $this->templateSheetEntry] = true;
            $this->templateSheet = null;
            $this->templateSheetEntry = null;
            $this->templateSheetName = null;
        }
    }

    // ─────────────────────────────────────────────────────────────────────
    // Template mode (v3.5)
    // ─────────────────────────────────────────────────────────────────────

    /**
     * Stream rows into a layout another producer authored.
     *
     * The template supplies everything above and below the data: header
     * rows, column widths, frozen panes, merges, conditional formatting,
     * page setup, and the style table its sample rows point into. This
     * writer cuts the chosen sheet once at `<sheetData>`, streams rows into
     * the cut, and carries every other ZIP entry across without inflating
     * it. Classic layout calls are refused rather than half-applied,
     * because the template already owns what they would set.
     *
     *   $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $path);
     *   $writer->sheet('Leaves', dataStartRow: 3);
     *   $writer->writeRow(['Ada', $start, 1200.0]);
     *   $writer->writeRow(['Grace', $start, 900.0], variant: 1);
     *   $writer->finishFile();
     *
     * $maxUniqueStrings caps the shared string dictionary, the one place
     * template mode holds memory proportional to the data. Past the ceiling
     * new strings are written inline instead, so a column of unique free text
     * degrades the file's size rather than the process's memory.
     */
    public function useTemplate(
        Template $template,
        int $maxUniqueStrings = SharedStringTable::DEFAULT_MAX_UNIQUE
    ): self {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        if ($this->started) {
            throw XlsxStreamException::alreadyStarted();
        }
        if ($this->compactMode) {
            throw XlsxStreamException::templateConflictsWith('compact()');
        }
        foreach ([
            'setHeaderStyle()' => $this->headerStyleId !== null,
            'setColumnWidths()' => $this->columnWidths !== [],
            'setAutoColumnWidth()' => $this->autoColumnWidth,
            'freezeRowsAndColumns()' => $this->freezeRows > 0 || $this->freezeColumns > 0,
            'enableAutoFilter()' => $this->autoFilterEnabled,
            'setColumnFormat()' => $this->columnStyleIds !== [],
        ] as $what => $configured) {
            if ($configured) {
                throw XlsxStreamException::templateConflictsWith($what);
            }
        }

        $this->template = $template;
        $this->templateMode = true;

        // The template's own tables become the authority: style ids the
        // sample rows hand out must keep meaning what they meant, and a
        // string we intern must land after the seed's own entries.
        $stylesEntry = 'xl/styles.xml';
        if ($template->hasEntry($stylesEntry)) {
            $this->styles = StyleRegistry::fromStylesXml($template->readEntry($stylesEntry));
        }

        // Shared strings are used only when the template already carries the
        // part. Creating one would need a new workbook relationship and a new
        // content-type override — a fifth seam this feature does not open —
        // so a template without one takes inline strings instead.
        $sstEntry = 'xl/sharedStrings.xml';
        if ($template->hasEntry($sstEntry)) {
            $this->sharedStrings = SharedStringTable::fromXml(
                $template->readEntry($sstEntry),
                $maxUniqueStrings
            );
        }

        return $this;
    }

    /**
     * Choose the template sheet to stream into and where its data begins.
     *
     * Rows above `$dataStartRow` are the template's header block and are
     * copied verbatim. The rows from `$dataStartRow` down are the sample
     * rows: they are read as a style oracle (one per variant) and are NOT
     * written to the output. Calling `sheet()` again finalizes the current
     * sheet and opens the next, so a workbook is written sheet by sheet.
     */
    public function sheet(string $name, int $dataStartRow): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        if (! $this->templateMode) {
            throw XlsxStreamException::templateModeRequired('sheet()');
        }

        $entry = $this->template->entryFor($name);
        if (isset($this->templateStreamedEntries[$entry]) || $entry === $this->templateSheetEntry) {
            throw XlsxStreamException::templateSheetAlreadyStreamed($name);
        }

        $sheet = $this->template->sheet($name, $dataStartRow);
        if (! $sheet->acceptsUnprefixedRows()) {
            throw XlsxStreamException::templateUnsupportedDialect($name);
        }

        // Close the sheet already in flight before opening the next entry —
        // one deflate context is live at a time.
        if ($this->templateSheet !== null) {
            $this->flushRowBuffer();
            $this->finishCurrentSheet();
        }

        $this->templateSheet = $sheet;
        $this->templateSheetName = $name;
        $this->templateSheetEntry = $entry;
        $this->started = true;
        $this->currentSheetIndex++;
        $this->startTemplateSheet();

        return $this;
    }

    /**
     * Open the chosen template sheet's ZIP entry and emit everything down to
     * the last header row, then park the row counter one below the first data
     * row so the next writeRow() lands exactly on dataStartRow.
     */
    protected function startTemplateSheet(): void
    {
        $sheet = $this->templateSheet;
        $entry = $this->templateSheetEntry;

        $this->sheets[] = [
            'index' => $this->currentSheetIndex,
            'name' => $this->templateSheetName,
            'filename' => $entry,
            'rows' => 0,
        ];

        $this->sheetOffset = $this->currentOffset;
        $this->currentSheetRow = 0;
        $this->inSampleMode = false;

        $this->resetPerSheetAccumulators();

        // The template's header rows are real rows in this sheet's first
        // index block. Fold them in before any data arrives so zone-map
        // pruning can never hide a row the un-pruned path would return —
        // the same reason startNewSheet() folds its own header row.
        if ($this->statsColumns !== [] || $this->stringStatsColumns !== []) {
            $this->foldTemplateHeaderRows();
        }

        $this->openSheetStream($entry);

        // Resolve the variants once. Every row reads a style map and a row
        // tag suffix; neither depends on the data, so neither is rebuilt.
        $this->templateStyleMaps = [];
        $this->templateRowSuffixes = [];
        for ($variant = 0, $n = $sheet->variantCount(); $variant < $n; $variant++) {
            $this->templateStyleMaps[$variant] = $sheet->styleMap($variant);
            $this->templateRowSuffixes[$variant] = self::renderRowSuffix($sheet->rowAttributes($variant));
        }

        $this->writeSheetData($sheet->head().$sheet->headerRowsXml());
        $this->currentSheetRow = $sheet->dataStartRow() - 1;
    }

    /**
     * Fold the template's header rows into this sheet's zone maps.
     *
     * currentSheetRow is moved onto each header row for the duration, so an
     * argmin/argmax pointer landing on a header cell records the row that
     * cell actually occupies. Sortedness is deliberately not tracked: the
     * header says nothing about how the data is ordered.
     */
    protected function foldTemplateHeaderRows(): void
    {
        $restore = $this->currentSheetRow;

        foreach ($this->templateSheet->headerRowValues($this->templateSharedStringsReader()) as $rowNumber => $values) {
            $this->currentSheetRow = $rowNumber;
            if ($this->statsColumns !== []) {
                $this->accumulateColumnStats($values, trackOrder: false);
            }
            if ($this->stringStatsColumns !== []) {
                $this->accumulateStringStats($values, trackOrder: false);
            }
        }

        $this->currentSheetRow = $restore;
    }

    /**
     * Read-side view of the template's shared strings, used to resolve the
     * t="s" cells in its header rows. Built once and only when something
     * actually needs to read a template cell back.
     */
    protected function templateSharedStringsReader(): ?SharedStrings
    {
        if ($this->templateSharedStringsReader !== null) {
            return $this->templateSharedStringsReader;
        }

        $entry = 'xl/sharedStrings.xml';
        if (! $this->template->hasEntry($entry)) {
            return null;
        }

        return $this->templateSharedStringsReader = SharedStringsParser::parseInMemory(
            $this->template->readEntry($entry)
        );
    }

    /**
     * Render a variant's row-level attributes into the `<row r="N"` suffix,
     * height and row style included, ending with the closing bracket.
     *
     * @param  array{ht: ?string, customHeight: bool, s: ?int, customFormat: bool}  $attributes
     */
    protected static function renderRowSuffix(array $attributes): string
    {
        $suffix = '';
        if ($attributes['ht'] !== null) {
            $suffix .= ' ht="'.$attributes['ht'].'"';
        }
        if ($attributes['customHeight']) {
            $suffix .= ' customHeight="1"';
        }
        if ($attributes['s'] !== null) {
            $suffix .= ' s="'.$attributes['s'].'"';
        }
        if ($attributes['customFormat']) {
            $suffix .= ' customFormat="1"';
        }

        return $suffix.'>';
    }

    /**
     * cellXfs id for a date in a column the sample rows never styled.
     *
     * Where the template holds an opinion its number format wins outright —
     * that is the whole point of the sample row. Past the sample's last
     * styled cell the template says nothing, so the classic date rule
     * applies; the format is appended to the seeded table on first use, so
     * a template whose data carries no dates keeps styles.xml untouched.
     */
    protected function templateDateStyle(): int
    {
        return $this->templateDateStyleId ??= $this->styles->registerBuiltinNumFmt(self::BUILTIN_NUMFMT_DATETIME);
    }

    /**
     * Refuse a classic operation whose result the template already owns.
     * Called from the setters themselves, never per row.
     */
    protected function refuseInTemplateMode(string $operation, string $because): void
    {
        if ($this->templateMode) {
            throw XlsxStreamException::templateModeForbids($operation, $because);
        }
    }

    /**
     * A template row is styled by its variant, so a registered row style
     * would have two sources fighting for the same s attribute.
     */
    protected function assertTemplateRowUnstyled(?int $rowStyleId): void
    {
        if ($rowStyleId !== null) {
            throw XlsxStreamException::templateModeForbids(
                'writeRow($row, $styleId)',
                'a template row is styled by its variant — pass variant: N instead'
            );
        }
    }

    /**
     * Template row builder, inline-string flavour (the template carries no
     * shared string table).
     *
     * Mirrors buildRowXml's type dispatch exactly; the difference is where
     * the s attribute comes from. The classic writer reads a per-column
     * format registered through the API, this one reads the sample row the
     * template shipped, so the number formats the template chose survive.
     */
    protected function buildRowXmlTemplate(int $rowIndex, array $data, int $variant, ?int $rowStyleId): string
    {
        $this->assertTemplateRowUnstyled($rowStyleId);

        $styleMap = $this->templateStyleMaps[$variant] ?? $this->templateStyleMaps[0];
        $count = count($data);
        if ($count > 0 && ! isset($this->colLetterCache[$count])) {
            for ($c = 1; $c <= $count; $c++) {
                $this->getColumnLetter($c);
            }
        }
        $letters = $this->colLetterCache;

        $xml = '<row r="'.$rowIndex.'"'.($this->templateRowSuffixes[$variant] ?? $this->templateRowSuffixes[0]);
        $col = 0;

        foreach ($data as $value) {
            $index = $col++;
            $cellRef = $letters[$col].$rowIndex;
            $style = $styleMap[$index] ?? null;
            $s = $style !== null ? ' s="'.$style.'"' : '';

            if (is_string($value)) {
                if ($value === '') {
                    $xml .= '<c r="'.$cellRef.'"'.$s.'/>';

                    continue;
                }

                if (is_numeric($value)) {
                    if ($this->shouldPreserveNumericString($value)) {
                        $xml .= '<c r="'.$cellRef.'"'.$s.' t="inlineStr"><is>'
                            .self::renderInlineText($this->fastXmlEscape($value)).'</is></c>';
                    } else {
                        $xml .= '<c r="'.$cellRef.'"'.$s.' t="n"><v>'.(0 + $value).'</v></c>';
                    }

                    continue;
                }

                $xml .= '<c r="'.$cellRef.'"'.$s.' t="inlineStr"><is>'
                    .self::renderInlineText($this->fastXmlEscape($value)).'</is></c>';

                continue;
            }

            if (is_int($value) || is_float($value)) {
                $xml .= '<c r="'.$cellRef.'"'.$s.' t="n"><v>'.$value.'</v></c>';

                continue;
            }

            if ($value === null) {
                $xml .= '<c r="'.$cellRef.'"'.$s.'/>';

                continue;
            }

            if (is_bool($value)) {
                $xml .= '<c r="'.$cellRef.'"'.$s.' t="b"><v>'.($value ? 1 : 0).'</v></c>';

                continue;
            }

            if ($value instanceof \DateTimeInterface) {
                $serial = ($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400;
                $dateStyle = $style ?? $this->templateDateStyle();
                $xml .= '<c r="'.$cellRef.'" s="'.$dateStyle.'" t="n"><v>'.$serial.'</v></c>';

                continue;
            }

            $xml .= '<c r="'.$cellRef.'"'.$s.' t="inlineStr"><is>'
                .self::renderInlineText($this->fastXmlEscape((string) $value)).'</is></c>';
        }

        return $xml.'</row>';
    }

    /**
     * Template row builder, shared-string flavour — the twin used when the
     * template ships an xl/sharedStrings.xml.
     *
     * Text cells become `t="s"` references into the seeded table, which is
     * what the producers that read these files expect: PhpSpreadsheet turns
     * an inline string into a RichText object, so a round trip through it
     * would not compare equal to the value that went in. When the dictionary
     * hits its ceiling intern() returns null and the cell falls back to an
     * inline string, so the file stays valid rather than growing without
     * bound.
     */
    protected function buildRowXmlTemplateShared(int $rowIndex, array $data, int $variant, ?int $rowStyleId): string
    {
        $this->assertTemplateRowUnstyled($rowStyleId);

        $sst = $this->sharedStrings;
        $styleMap = $this->templateStyleMaps[$variant] ?? $this->templateStyleMaps[0];
        $count = count($data);
        if ($count > 0 && ! isset($this->colLetterCache[$count])) {
            for ($c = 1; $c <= $count; $c++) {
                $this->getColumnLetter($c);
            }
        }
        $letters = $this->colLetterCache;

        $xml = '<row r="'.$rowIndex.'"'.($this->templateRowSuffixes[$variant] ?? $this->templateRowSuffixes[0]);
        $col = 0;

        foreach ($data as $value) {
            $index = $col++;
            $cellRef = $letters[$col].$rowIndex;
            $style = $styleMap[$index] ?? null;
            $s = $style !== null ? ' s="'.$style.'"' : '';

            if (is_string($value)) {
                if ($value === '') {
                    $xml .= '<c r="'.$cellRef.'"'.$s.'/>';

                    continue;
                }

                if (is_numeric($value) && ! $this->shouldPreserveNumericString($value)) {
                    $xml .= '<c r="'.$cellRef.'"'.$s.' t="n"><v>'.(0 + $value).'</v></c>';

                    continue;
                }

                $escaped = $this->fastXmlEscape($value);
                $shared = $sst->intern($escaped);
                $xml .= $shared !== null
                    ? '<c r="'.$cellRef.'"'.$s.' t="s"><v>'.$shared.'</v></c>'
                    : '<c r="'.$cellRef.'"'.$s.' t="inlineStr"><is>'.self::renderInlineText($escaped).'</is></c>';

                continue;
            }

            if (is_int($value) || is_float($value)) {
                $xml .= '<c r="'.$cellRef.'"'.$s.' t="n"><v>'.$value.'</v></c>';

                continue;
            }

            if ($value === null) {
                $xml .= '<c r="'.$cellRef.'"'.$s.'/>';

                continue;
            }

            if (is_bool($value)) {
                $xml .= '<c r="'.$cellRef.'"'.$s.' t="b"><v>'.($value ? 1 : 0).'</v></c>';

                continue;
            }

            if ($value instanceof \DateTimeInterface) {
                $serial = ($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400;
                $dateStyle = $style ?? $this->templateDateStyle();
                $xml .= '<c r="'.$cellRef.'" s="'.$dateStyle.'" t="n"><v>'.$serial.'</v></c>';

                continue;
            }

            $escaped = $this->fastXmlEscape((string) $value);
            $shared = $sst->intern($escaped);
            $xml .= $shared !== null
                ? '<c r="'.$cellRef.'"'.$s.' t="s"><v>'.$shared.'</v></c>'
                : '<c r="'.$cellRef.'"'.$s.' t="inlineStr"><is>'.self::renderInlineText($escaped).'</is></c>';
        }

        return $xml.'</row>';
    }

    /**
     * Wrap already-escaped text in `<t>`, adding xml:space="preserve" only
     * when a leading or trailing space would otherwise be collapsed away.
     */
    protected static function renderInlineText(string $escaped): string
    {
        if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
            $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
            return '<t xml:space="preserve">'.$escaped.'</t>';
        }

        return '<t>'.$escaped.'</t>';
    }

    /**
     * Declare the sidecar's content type in a template's [Content_Types].xml.
     *
     * Every package part needs a declared type, and the sidecar's "bin"
     * extension has no Default mapping, so without this override Excel's
     * validator drops into repair mode. This is the only edit template mode
     * makes to a part it did not author: one Override element before the
     * closing tag, added only when the index is on and only when it is not
     * already there, so a template that already declared it is untouched.
     */
    protected static function declareIndexContentType(string $xml): string
    {
        if (str_contains($xml, 'PartName="/'.RandomAccessIndex::ENTRY_PATH.'"')) {
            return $xml;
        }

        if (! preg_match('/<\/((?:[A-Za-z_][\w.\-]*:)?)Types\s*>/', $xml, $match, PREG_OFFSET_CAPTURE)) {
            throw XlsxStreamException::templateContentTypesUnreadable();
        }

        [$prefix, $at] = [$match[1][0], $match[0][1]];
        $override = '<'.$prefix.'Override PartName="/'.RandomAccessIndex::ENTRY_PATH.'" '
            .'ContentType="application/octet-stream"/>';

        return substr($xml, 0, $at).$override.substr($xml, $at);
    }

    /**
     * Move one template entry into the output without touching its bytes.
     *
     * The compressed body is copied through as-is — no inflate, no deflate —
     * so the CRC and the compressed size in the central directory are the
     * template's own and the copy is provably faithful. STORED entries stay
     * stored. Sizes come from the template's central directory, so an entry
     * written with a streaming data descriptor arrives with its sizes in the
     * local header and no descriptor of its own.
     */
    protected function copyEntry(string $name): void
    {
        $zip = $this->template->zip();
        $source = $this->template->source();
        $entry = $zip->entry($name);
        if ($entry === null) {
            throw XlsxStreamException::templateSheetDataMissing($name);
        }

        $this->assertZip32Compatible($this->currentOffset, "cumulative archive offset before copying '{$name}'");

        [$mtime, $mdate] = $this->dosTimeParts(time());

        $header = pack('V', self::LOCAL_FILE_HEADER_SIGNATURE);
        $header .= pack('v', self::VERSION_NEEDED);
        $header .= pack('v', 0x0000);
        $header .= pack('v', $entry['method']);
        $header .= pack('v', $mtime);
        $header .= pack('v', $mdate);
        $header .= pack('V', $entry['crc32']);
        $header .= pack('V', $entry['compressed_size']);
        $header .= pack('V', $entry['uncompressed_size']);
        $header .= pack('v', strlen($name));
        $header .= pack('v', 0);
        $header .= $name;

        $offset = $this->currentOffset;
        $this->writeToDest($header);

        $remaining = $entry['compressed_size'];
        if ($remaining > 0) {
            $stream = $source->streamFrom($zip->dataOffset($source, $name));
            try {
                while ($remaining > 0) {
                    $chunk = fread($stream, (int) min(65536, $remaining));
                    if (! is_string($chunk) || $chunk === '') {
                        throw XlsxStreamException::templateSheetDataMissing($name);
                    }
                    $this->writeToDest($chunk);
                    $remaining -= strlen($chunk);
                }
            } finally {
                if (is_resource($stream)) {
                    fclose($stream);
                }
            }
        }

        $this->centralDirectory[] = [
            'filename' => $name,
            'crc32' => $entry['crc32'],
            'compressed_size' => $entry['compressed_size'],
            'uncompressed_size' => $entry['uncompressed_size'],
            'offset' => $offset,
            'compression' => $entry['method'],
            'flags' => 0x0000,
            'timestamp' => time(),
        ];
    }

    /**
     * Close the streamed sheets, then carry every remaining template entry
     * across. Only the parts this writer actually changed are re-authored:
     * the style table when a format was appended, the shared string table
     * when a string was interned. Everything else — the workbook, the rels,
     * the content types, the sheets nobody streamed into, themes, drawings —
     * is moved byte for byte. The caller owns closing the sink.
     */
    protected function writeTemplateRemainder(): void
    {
        if ($this->templateSheet !== null) {
            $this->flushRowBuffer();
            $this->finishCurrentSheet();
        }

        if ($this->sheets === []) {
            throw XlsxStreamException::templateSheetNotSelected();
        }

        $stylesEntry = 'xl/styles.xml';
        $sstEntry = 'xl/sharedStrings.xml';
        $contentTypes = '[Content_Types].xml';

        if ($this->randomAccessIndexEnabled) {
            $this->writeStaticFile(RandomAccessIndex::ENTRY_PATH, $this->buildRandomAccessIndexPayload());
        }

        foreach ($this->template->zip()->names() as $name) {
            if (isset($this->templateStreamedEntries[$name])) {
                continue;
            }
            // A template produced by this package may already carry a
            // sidecar. It describes rows that are no longer there, and we
            // have just written our own, so the stale one is dropped.
            if ($name === RandomAccessIndex::ENTRY_PATH) {
                continue;
            }
            if ($name === $contentTypes && $this->randomAccessIndexEnabled) {
                $this->writeStaticFile($name, self::declareIndexContentType($this->template->readEntry($name)));

                continue;
            }
            if ($name === $stylesEntry && ! $this->styles->isSeedUnmodified()) {
                $this->writeStaticFile($name, $this->styles->toXml());

                continue;
            }
            if (
                $name === $sstEntry
                && $this->sharedStrings !== null
                && ! $this->sharedStrings->isSeedUnmodified()
            ) {
                $this->writeStaticFile($name, $this->sharedStrings->toXml());

                continue;
            }

            $this->copyEntry($name);
        }

        $this->writeCentralDirectory();

        // Release the template's own handle when this writer opened it; a
        // caller who passed a Template in keeps it and may reuse the layout.
        if ($this->templateOwned) {
            $this->template->close();
        }
    }

    /**
     * Build row XML with optimization.
     *
     * The unstyled path ($rowStyleId === null) is the hot path and is kept
     * byte-for-byte identical to pre-v3.0.2 — a single null check delegates
     * the rare styled case to buildStyledRowXml so the common case pays
     * nothing for the feature.
     */
    protected function buildRowXml(int $rowIndex, array $data, ?int $rowStyleId = null): string
    {
        // Compact dispatch first: one boolean test per row is the whole
        // price on the classic path (the tokenizer date-twin pattern —
        // the two shapes share no hot loop, so neither taxes the other).
        if ($this->compactMode) {
            return $rowStyleId !== null
                ? $this->buildStyledRowXmlCompact($data, $rowStyleId)
                : $this->buildRowXmlCompact($data);
        }

        if ($rowStyleId !== null) {
            return $this->buildStyledRowXml($rowIndex, $data, $rowStyleId);
        }

        // Hot-path shape (measured, output byte-identical to the previous
        // array+implode form):
        // - strings are dispatched first — they are the most common cell
        //   type and previously paid four failed type checks per cell;
        // - the column-letter cache is pre-filled once and read through a
        //   local (COW copy, no per-cell method call + property deref);
        // - cells append straight onto one string; PHP reallocs a
        //   refcount-1 string in place, which beats array insert + implode.
        $count = count($data);
        if ($count > 0 && ! isset($this->colLetterCache[$count])) {
            for ($c = 1; $c <= $count; $c++) {
                $this->getColumnLetter($c);
            }
        }
        $letters = $this->colLetterCache;
        $columnStyleIds = $this->columnStyleIds;
        $hasColumnStyles = ! empty($columnStyleIds);

        $xml = '<row r="' . $rowIndex . '">';
        $col = 0;

        foreach ($data as $value) {
            $cellRef = $letters[++$col] . $rowIndex;

            if (is_string($value)) {
                // Empty string -> empty cell (matches the null cell below)
                if ($value === '') {
                    $xml .= '<c r="' . $cellRef . '"/>';
                    continue;
                }

                // Numeric strings: preserve as string when precision/format would be lost
                if (is_numeric($value)) {
                    if ($this->shouldPreserveNumericString($value)) {
                        $xml .= '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $this->fastXmlEscape($value) . '</t></is></c>';
                    } elseif ($hasColumnStyles && isset($columnStyleIds[$col])) {
                        $xml .= '<c r="' . $cellRef . '" s="' . $columnStyleIds[$col] . '" t="n"><v>' . (0 + $value) . '</v></c>';
                    } else {
                        $xml .= '<c r="' . $cellRef . '" t="n"><v>' . (0 + $value) . '</v></c>';
                    }
                    continue;
                }

                $escaped = $this->fastXmlEscape($value);
                if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                    $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                    $xml .= '<c r="' . $cellRef . '" t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
                } else {
                    $xml .= '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
                }
                continue;
            }

            // Numeric values — split fast/slow so unstyled exports keep v2.0.1 cost
            if (is_int($value) || is_float($value)) {
                if ($hasColumnStyles && isset($columnStyleIds[$col])) {
                    $xml .= '<c r="' . $cellRef . '" s="' . $columnStyleIds[$col] . '" t="n"><v>' . $value . '</v></c>';
                } else {
                    $xml .= '<c r="' . $cellRef . '" t="n"><v>' . $value . '</v></c>';
                }
                continue;
            }

            // Null -> empty cell
            if ($value === null) {
                $xml .= '<c r="' . $cellRef . '"/>';
                continue;
            }

            // Boolean -> native Excel boolean cell
            if (is_bool($value)) {
                $xml .= '<c r="' . $cellRef . '" t="b"><v>' . ($value ? 1 : 0) . '</v></c>';
                continue;
            }

            // DateTimeInterface -> Excel serial date with datetime style
            // (column-specific format wins if set, else fallback to legacy STYLE_DATETIME)
            if ($value instanceof \DateTimeInterface) {
                $serial = ($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400;
                $styleId = $hasColumnStyles && isset($columnStyleIds[$col])
                    ? $columnStyleIds[$col]
                    : self::STYLE_DATETIME;
                $xml .= '<c r="' . $cellRef . '" s="' . $styleId . '" t="n"><v>' . $serial . '</v></c>';
                continue;
            }

            // Stringable / other -> inlineStr
            $escaped = $this->fastXmlEscape((string) $value);

            if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                $xml .= '<c r="' . $cellRef . '" t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
            } else {
                $xml .= '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
            }
        }

        return $xml . '</row>';
    }

    /**
     * Styled-row variant of buildRowXml — stamps $rowStyleId onto every cell
     * so the whole row carries the fill/font, including empty cells (so the
     * highlight doesn't show gaps).
     *
     * Number formats still win their cell: where a column has its own format,
     * the row's fill/font is merged with that column's numFmt (memoized in
     * rowStyleMergeCache) rather than overwriting it. Type detection mirrors
     * buildRowXml exactly; only the s="N" attribute differs.
     */
    protected function buildStyledRowXml(int $rowIndex, array $data, int $rowStyleId): string
    {
        // Mirrors buildRowXml's flattened hot-path shape (string-first
        // dispatch, local letter cache, direct string append) — see the
        // comment there. Only the s="N" handling differs.
        $count = count($data);
        if ($count > 0 && ! isset($this->colLetterCache[$count])) {
            for ($c = 1; $c <= $count; $c++) {
                $this->getColumnLetter($c);
            }
        }
        $letters = $this->colLetterCache;
        $columnStyleIds = $this->columnStyleIds;
        $hasColumnStyles = ! empty($columnStyleIds);
        // Hoist the row's style attribute so it's stringified once per row,
        // not once per cell. Cells that need a merged (column-format) style
        // build their own; everything else reuses this.
        $sAttr = ' s="' . $rowStyleId . '"';

        $xml = '<row r="' . $rowIndex . '">';
        $col = 0;

        foreach ($data as $value) {
            $cellRef = $letters[++$col] . $rowIndex;

            if (is_string($value)) {
                // Empty string -> empty but still filled cell
                if ($value === '') {
                    $xml .= '<c r="' . $cellRef . '"' . $sAttr . '/>';
                    continue;
                }

                // Numeric strings: preserve as string when precision/format would be lost
                if (is_numeric($value)) {
                    if ($this->shouldPreserveNumericString($value)) {
                        $xml .= '<c r="' . $cellRef . '"' . $sAttr . ' t="inlineStr"><is><t>' . $this->fastXmlEscape($value) . '</t></is></c>';
                    } elseif ($hasColumnStyles && isset($columnStyleIds[$col])) {
                        $styleId = $this->mergeRowStyleWithColumnStyle($rowStyleId, $columnStyleIds[$col]);
                        $xml .= '<c r="' . $cellRef . '" s="' . $styleId . '" t="n"><v>' . (0 + $value) . '</v></c>';
                    } else {
                        $xml .= '<c r="' . $cellRef . '"' . $sAttr . ' t="n"><v>' . (0 + $value) . '</v></c>';
                    }
                    continue;
                }

                $escaped = $this->fastXmlEscape($value);
                if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                    $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                    $xml .= '<c r="' . $cellRef . '"' . $sAttr . ' t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
                } else {
                    $xml .= '<c r="' . $cellRef . '"' . $sAttr . ' t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
                }
                continue;
            }

            // Numeric values
            if (is_int($value) || is_float($value)) {
                if ($hasColumnStyles && isset($columnStyleIds[$col])) {
                    $styleId = $this->mergeRowStyleWithColumnStyle($rowStyleId, $columnStyleIds[$col]);
                    $xml .= '<c r="' . $cellRef . '" s="' . $styleId . '" t="n"><v>' . $value . '</v></c>';
                } else {
                    $xml .= '<c r="' . $cellRef . '"' . $sAttr . ' t="n"><v>' . $value . '</v></c>';
                }
                continue;
            }

            // Null -> empty but still filled cell
            if ($value === null) {
                $xml .= '<c r="' . $cellRef . '"' . $sAttr . '/>';
                continue;
            }

            // Boolean -> native Excel boolean cell
            if (is_bool($value)) {
                $xml .= '<c r="' . $cellRef . '"' . $sAttr . ' t="b"><v>' . ($value ? 1 : 0) . '</v></c>';
                continue;
            }

            // DateTimeInterface -> Excel serial date; merge row style with the
            // column's date format (or the legacy datetime style as fallback).
            if ($value instanceof \DateTimeInterface) {
                $serial = ($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400;
                $columnStyleId = $hasColumnStyles && isset($columnStyleIds[$col])
                    ? $columnStyleIds[$col]
                    : self::STYLE_DATETIME;
                $styleId = $this->mergeRowStyleWithColumnStyle($rowStyleId, $columnStyleId);
                $xml .= '<c r="' . $cellRef . '" s="' . $styleId . '" t="n"><v>' . $serial . '</v></c>';
                continue;
            }

            // Stringable / other -> inlineStr
            $escaped = $this->fastXmlEscape((string) $value);

            if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                $xml .= '<c r="' . $cellRef . '"' . $sAttr . ' t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
            } else {
                $xml .= '<c r="' . $cellRef . '"' . $sAttr . ' t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
            }
        }

        return $xml . '</row>';
    }

    /**
     * Compact (r-less) twin of buildRowXml — see compact(). Every branch
     * is the exact projection of its classic counterpart with the
     * ` r="A1"` attribute removed; the transform byte-oracle test pins
     * that equivalence, so the two builders cannot drift apart.
     *
     * No cell references means no column-letter cache fill and no
     * per-cell letters/rowIndex concatenation — the branch bodies are
     * otherwise IDENTICAL to the classic builder. Empty cells still
     * emit `<c/>`: with sequential position assignment the placeholder
     * IS the position carrier, dropping it would shift every later
     * cell left.
     */
    protected function buildRowXmlCompact(array $data): string
    {
        $columnStyleIds = $this->columnStyleIds;
        $hasColumnStyles = ! empty($columnStyleIds);

        $xml = '<row>';
        $col = 0;

        foreach ($data as $value) {
            ++$col;

            if (is_string($value)) {
                if ($value === '') {
                    $xml .= '<c/>';
                    continue;
                }

                if (is_numeric($value)) {
                    if ($this->shouldPreserveNumericString($value)) {
                        $xml .= '<c t="inlineStr"><is><t>' . $this->fastXmlEscape($value) . '</t></is></c>';
                    } elseif ($hasColumnStyles && isset($columnStyleIds[$col])) {
                        $xml .= '<c s="' . $columnStyleIds[$col] . '" t="n"><v>' . (0 + $value) . '</v></c>';
                    } else {
                        $xml .= '<c t="n"><v>' . (0 + $value) . '</v></c>';
                    }
                    continue;
                }

                $escaped = $this->fastXmlEscape($value);
                if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                    $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                    $xml .= '<c t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
                } else {
                    $xml .= '<c t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
                }
                continue;
            }

            if (is_int($value) || is_float($value)) {
                if ($hasColumnStyles && isset($columnStyleIds[$col])) {
                    $xml .= '<c s="' . $columnStyleIds[$col] . '" t="n"><v>' . $value . '</v></c>';
                } else {
                    $xml .= '<c t="n"><v>' . $value . '</v></c>';
                }
                continue;
            }

            if ($value === null) {
                $xml .= '<c/>';
                continue;
            }

            if (is_bool($value)) {
                $xml .= '<c t="b"><v>' . ($value ? 1 : 0) . '</v></c>';
                continue;
            }

            if ($value instanceof \DateTimeInterface) {
                $serial = ($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400;
                $styleId = $hasColumnStyles && isset($columnStyleIds[$col])
                    ? $columnStyleIds[$col]
                    : self::STYLE_DATETIME;
                $xml .= '<c s="' . $styleId . '" t="n"><v>' . $serial . '</v></c>';
                continue;
            }

            $escaped = $this->fastXmlEscape((string) $value);

            if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                $xml .= '<c t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
            } else {
                $xml .= '<c t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
            }
        }

        return $xml . '</row>';
    }

    /**
     * Compact (r-less) twin of buildStyledRowXml — same projection rule
     * and byte-oracle pin as buildRowXmlCompact. Styled empty cells keep
     * their `<c s="N"/>` placeholder for the same position-carrier
     * reason.
     */
    protected function buildStyledRowXmlCompact(array $data, int $rowStyleId): string
    {
        $columnStyleIds = $this->columnStyleIds;
        $hasColumnStyles = ! empty($columnStyleIds);
        $sAttr = ' s="' . $rowStyleId . '"';

        $xml = '<row>';
        $col = 0;

        foreach ($data as $value) {
            ++$col;

            if (is_string($value)) {
                if ($value === '') {
                    $xml .= '<c' . $sAttr . '/>';
                    continue;
                }

                if (is_numeric($value)) {
                    if ($this->shouldPreserveNumericString($value)) {
                        $xml .= '<c' . $sAttr . ' t="inlineStr"><is><t>' . $this->fastXmlEscape($value) . '</t></is></c>';
                    } elseif ($hasColumnStyles && isset($columnStyleIds[$col])) {
                        $styleId = $this->mergeRowStyleWithColumnStyle($rowStyleId, $columnStyleIds[$col]);
                        $xml .= '<c s="' . $styleId . '" t="n"><v>' . (0 + $value) . '</v></c>';
                    } else {
                        $xml .= '<c' . $sAttr . ' t="n"><v>' . (0 + $value) . '</v></c>';
                    }
                    continue;
                }

                $escaped = $this->fastXmlEscape($value);
                if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                    $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                    $xml .= '<c' . $sAttr . ' t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
                } else {
                    $xml .= '<c' . $sAttr . ' t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
                }
                continue;
            }

            if (is_int($value) || is_float($value)) {
                if ($hasColumnStyles && isset($columnStyleIds[$col])) {
                    $styleId = $this->mergeRowStyleWithColumnStyle($rowStyleId, $columnStyleIds[$col]);
                    $xml .= '<c s="' . $styleId . '" t="n"><v>' . $value . '</v></c>';
                } else {
                    $xml .= '<c' . $sAttr . ' t="n"><v>' . $value . '</v></c>';
                }
                continue;
            }

            if ($value === null) {
                $xml .= '<c' . $sAttr . '/>';
                continue;
            }

            if (is_bool($value)) {
                $xml .= '<c' . $sAttr . ' t="b"><v>' . ($value ? 1 : 0) . '</v></c>';
                continue;
            }

            if ($value instanceof \DateTimeInterface) {
                $serial = ($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400;
                $columnStyleId = $hasColumnStyles && isset($columnStyleIds[$col])
                    ? $columnStyleIds[$col]
                    : self::STYLE_DATETIME;
                $styleId = $this->mergeRowStyleWithColumnStyle($rowStyleId, $columnStyleId);
                $xml .= '<c s="' . $styleId . '" t="n"><v>' . $serial . '</v></c>';
                continue;
            }

            $escaped = $this->fastXmlEscape((string) $value);

            if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                $xml .= '<c' . $sAttr . ' t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
            } else {
                $xml .= '<c' . $sAttr . ' t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
            }
        }

        return $xml . '</row>';
    }

    /**
     * Memoized wrapper over StyleRegistry::mergeRowStyleWithColumn so a
     * (row-style, column-style) pair only hits the registry's dedup scan
     * once, not on every styled cell.
     */
    protected function mergeRowStyleWithColumnStyle(int $rowStyleId, int $columnStyleId): int
    {
        return $this->rowStyleMergeCache[$rowStyleId][$columnStyleId]
            ??= $this->styles->mergeRowStyleWithColumn($rowStyleId, $columnStyleId);
    }

    /**
     * Decide whether a numeric string should be preserved as a string cell.
     *
     * Preserve when casting to float would lose precision or strip formatting:
     * - Leading zero (e.g. "0123" — phone numbers, codes)
     * - Plus sign prefix (e.g. "+90123")
     * - Integer string longer than 15 significant digits (PHP float precision limit)
     */
    protected function shouldPreserveNumericString(string $value): bool
    {
        $len = strlen($value);
        if ($len < 2) {
            return false;
        }

        // Leading zero on non-decimal: "0123" yes, "0.5" no
        if ($value[0] === '0' && $value[1] !== '.') {
            return true;
        }

        // Plus sign prefix: "+0123", "+12345"
        if ($value[0] === '+') {
            return true;
        }

        // Big integer: > 15 digits (with optional leading minus) and no decimal/exponent
        $check = $value[0] === '-' ? substr($value, 1) : $value;
        if (strlen($check) > 15 && ctype_digit($check)) {
            return true;
        }

        return false;
    }

    /**
     * Get column letter with caching
     */
    protected function getColumnLetter(int $index): string
    {
        if (!isset($this->colLetterCache[$index])) {
            $n = $index;
            $s = '';
            while ($n > 0) {
                $n--;
                $s = chr(65 + ($n % 26)) . $s;
                $n = intdiv($n, 26);
            }
            $this->colLetterCache[$index] = $s;
        }
        return $this->colLetterCache[$index];
    }

    /**
     * Ultra-fast XML escaping
     */
    protected function fastXmlEscape(string $str): string
    {
        if (! preg_match(self::XML_ESCAPE_NEEDED, $str)) {
            return $str;
        }

        // Only pay the control-byte strip when a control byte is actually
        // present — the usual dirty string just has an '&' or a quote.
        if (preg_match(self::XML_CTRL_BYTES, $str)) {
            $str = preg_replace(self::XML_CTRL_BYTES, '', $str);
        }

        return strtr($str, self::XML_ESCAPE_MAP);
    }

    /**
     * Write static ZIP entry
     */
    /**
     * Reject writes that would push a 32-bit size or offset field past
     * its limit. The ZIP local-file-header and central-directory layouts
     * pack these fields with 'V' (uint32); silently truncating them
     * produces an archive that opens but lies about every byte position
     * after the truncation. Caller gets a clear exception instead.
     */
    private function assertZip32Compatible(int $value, string $context): void
    {
        if ($value > self::ZIP32_MAX_SIZE) {
            $mb = number_format($value / 1024 / 1024, 1);
            throw XlsxStreamException::zip32LimitExceeded("{$context} would be {$mb} MB which exceeds the 4 GB ZIP32 limit");
        }
    }

    protected function writeStaticFile(string $filename, string $content): void
    {
        $uncompressedSize = strlen($content);
        $compressedContent = gzdeflate($content, $this->deflateLevel);
        $compressedSize = strlen($compressedContent);
        $crc32 = crc32($content);

        $this->assertZip32Compatible($uncompressedSize, "static file '{$filename}' uncompressed size");
        $this->assertZip32Compatible($compressedSize, "static file '{$filename}' compressed size");
        $this->assertZip32Compatible($this->currentOffset, 'cumulative archive offset before static file');

        [$mtime, $mdate] = $this->dosTimeParts(time());

        $this->centralDirectory[] = [
            'filename' => $filename,
            'crc32' => $crc32,
            'compressed_size' => $compressedSize,
            'uncompressed_size' => $uncompressedSize,
            'offset' => $this->currentOffset,
            'compression' => self::COMPRESSION_DEFLATED,
            'flags' => 0x0000,
            'timestamp' => time(),
        ];

        $header = pack('V', self::LOCAL_FILE_HEADER_SIGNATURE);
        $header .= pack('v', self::VERSION_NEEDED);
        $header .= pack('v', 0x0000);
        $header .= pack('v', self::COMPRESSION_DEFLATED);
        $header .= pack('v', $mtime);
        $header .= pack('v', $mdate);
        $header .= pack('V', $crc32);
        $header .= pack('V', $compressedSize);
        $header .= pack('V', $uncompressedSize);
        $header .= pack('v', strlen($filename));
        $header .= pack('v', 0);
        $header .= $filename;

        $this->writeToDest($header);
        $this->writeToDest($compressedContent);
    }

    /**
     * Write ZIP central directory
     */
    protected function writeCentralDirectory(): void
    {
        $entryCount = count($this->centralDirectory);
        if ($entryCount > self::ZIP32_MAX_ENTRIES) {
            throw XlsxStreamException::zip32LimitExceeded(
                "central directory would carry {$entryCount} entries, exceeding the 65535 ZIP32 limit"
            );
        }
        $this->assertZip32Compatible($this->currentOffset, 'central directory start offset');

        $centralDirStart = $this->currentOffset;
        $centralDirSize = 0;

        foreach ($this->centralDirectory as $entry) {
            [$mtime, $mdate] = $this->dosTimeParts($entry['timestamp']);

            $header = pack('V', self::CENTRAL_FILE_HEADER_SIGNATURE);
            $header .= pack('v', self::VERSION_MADE_BY);
            $header .= pack('v', self::VERSION_NEEDED);
            $header .= pack('v', $entry['flags']);
            $header .= pack('v', $entry['compression']);
            $header .= pack('v', $mtime);
            $header .= pack('v', $mdate);
            $header .= pack('V', $entry['crc32']);
            $header .= pack('V', $entry['compressed_size']);
            $header .= pack('V', $entry['uncompressed_size']);
            $header .= pack('v', strlen($entry['filename']));
            $header .= pack('v', 0);
            $header .= pack('v', 0);
            $header .= pack('v', 0);
            $header .= pack('v', 0);
            $header .= pack('V', 0x81A40000);
            $header .= pack('V', $entry['offset']);
            $header .= $entry['filename'];

            $this->writeToDest($header);
            $centralDirSize += strlen($header);
        }

        $endRecord = pack('V', self::END_OF_CENTRAL_DIR_SIGNATURE);
        $endRecord .= pack('v', 0);
        $endRecord .= pack('v', 0);
        $endRecord .= pack('v', count($this->centralDirectory));
        $endRecord .= pack('v', count($this->centralDirectory));
        $endRecord .= pack('V', $centralDirSize);
        $endRecord .= pack('V', $centralDirStart);
        $endRecord .= pack('v', 0);

        $this->writeToDest($endRecord);
    }

    /**
     * Finalize the XLSX file
     */
    public function finishFile(): array
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        if (!$this->started) {
            throw $this->templateMode
                ? XlsxStreamException::templateSheetNotSelected()
                : XlsxStreamException::headersNotSet();
        }

        // Sample never reached its target — finalize with whatever we
        // collected so the preamble + header still get emitted.
        if ($this->inSampleMode && ! $this->autoWidthFinalized) {
            $this->finalizeAutoWidthSample();
        }

        if ($this->templateMode) {
            $this->writeTemplateRemainder();
            $this->closed = true;

            return [
                'bytes' => $this->currentOffset,
                'rows' => $this->totalRows,
                'sheets' => count($this->sheets),
                'sheet_details' => $this->sheets,
            ];
        }

        if ($this->currentSheetRow > 0) {
            $this->flushRowBuffer();
            $this->finishCurrentSheet();
        }

        if (empty($this->sheets)) {
            throw XlsxStreamException::emptyWorkbook();
        }

        $this->writeStaticFile('xl/styles.xml', $this->getStylesXml());
        if ($this->randomAccessIndexEnabled) {
            $this->writeStaticFile(RandomAccessIndex::ENTRY_PATH, $this->buildRandomAccessIndexPayload());
        }
        $this->writeStaticFile('xl/_rels/workbook.xml.rels', $this->getWorkbookRelsXml());
        $this->writeStaticFile('xl/workbook.xml', $this->getWorkbookXml());
        $this->writeStaticFile('[Content_Types].xml', $this->getContentTypesXml());

        $this->writeCentralDirectory();

        $this->closed = true;

        return [
            'bytes' => $this->currentOffset,
            'rows' => $this->totalRows,
            'sheets' => count($this->sheets),
            'sheet_details' => $this->sheets,
        ];
    }

    // XLSX structure generators

    protected function getContentTypesXml(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>';

        for ($i = 1; $i <= count($this->sheets); $i++) {
            $xml .= "\n    " . '<Override PartName="/xl/worksheets/sheet' . $i . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }

        $xml .= '
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';

        // Every package part must have a declared content type. The
        // random-access index sidecar uses extension "bin" which has no
        // Default mapping, so an explicit Override is required — without
        // it Excel's strict validator triggers repair mode on open.
        // application/octet-stream signals "opaque binary, do not interpret".
        if ($this->randomAccessIndexEnabled) {
            $xml .= "\n    " . '<Override PartName="/'.RandomAccessIndex::ENTRY_PATH.'" ContentType="application/octet-stream"/>';
        }

        $xml .= '
</Types>';

        return $xml;
    }

    protected function getRelsXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
    }

    protected function getWorkbookRelsXml(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

        for ($i = 1; $i <= count($this->sheets); $i++) {
            $xml .= "\n    " . '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $i . '.xml"/>';
        }

        $styleId = count($this->sheets) + 1;
        $xml .= "\n    " . '<Relationship Id="rId' . $styleId . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';

        $xml .= '
</Relationships>';

        return $xml;
    }

    protected function getWorkbookXml(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>';

        foreach ($this->sheets as $sheet) {
            $escapedName = $this->fastXmlEscape($sheet['name']);
            $xml .= "\n        " . '<sheet name="' . $escapedName . '" sheetId="' . $sheet['index'] . '" r:id="rId' . $sheet['index'] . '"/>';
        }

        $xml .= '
    </sheets>
</workbook>';

        return $xml;
    }

    protected function getStylesXml(): string
    {
        return $this->styles->toXml();
    }
}

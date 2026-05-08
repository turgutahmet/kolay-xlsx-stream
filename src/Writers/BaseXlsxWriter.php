<?php

namespace Kolay\XlsxStream\Writers;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Styles\StyleRegistry;

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

    // Performance optimizations
    protected int $bufferFlushInterval = 1000;
    protected string $rowBuffer = '';
    protected int $rowBufferCount = 0;
    protected int $deflateLevel = 6;

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

    /**
     * Per-sheet sync points keyed by sheet entry path.
     *
     * @var array<string, list<array{row: int, comp_offset: int, uncomp_offset: int}>>
     */
    protected array $indexSyncPoints = [];

    public function __construct()
    {
        $this->styles = new StyleRegistry();
    }

    /**
     * Write data to destination (must be implemented by child classes)
     */
    abstract protected function writeToDest(string $data): void;

    /**
     * Set deflate compression level (1-9)
     * 3 = fast (20-35% speed boost, larger files)
     * 6 = balanced (default)
     * 9 = best compression (slower)
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
        // Allowed before startFile() (for sheet 1) and between newSheet() calls
        // (to give each manually-rotated sheet its own header style).
        $this->headerStyleId = $this->styles->registerHeaderStyle($options);

        return $this;
    }

    /**
     * Freeze the first row so it stays visible while scrolling.
     *
     * Equivalent to Excel's "View > Freeze Top Row".
     */
    public function freezeFirstRow(): self
    {
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
     *
     * Numeric values in that column are wrapped with the chosen format.
     * String values pass through unchanged.
     */
    public function setColumnFormat(int $column, string|int $format): self
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
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

        $this->columnStyleIds[$column] = $this->styles->registerColumnFormat($format);
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
     * Write a single row (handles multi-sheet automatically)
     */
    public function writeRow(array $row): void
    {
        if (!$this->started) {
            throw XlsxStreamException::headersNotSet();
        }
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        if (count($row) > self::MAX_COLUMNS) {
            throw XlsxStreamException::tooManyColumns(count($row), self::MAX_COLUMNS);
        }

        // Check if we need to start a new sheet
        if ($this->currentSheetRow === 0 || $this->currentSheetRow >= self::ROWS_PER_SHEET) {
            if ($this->currentSheetRow > 0) {
                $this->flushRowBuffer();
                $this->finishCurrentSheet();
            }
            $this->currentSheetIndex++;
            $this->startNewSheet();
        }

        $this->currentSheetRow++;
        $this->totalRows++;

        $rowXml = $this->buildRowXml($this->currentSheetRow, $row);

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
    public function writeRows(iterable $rows): void
    {
        foreach ($rows as $row) {
            $this->writeRow($row);
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
    protected function flushRowBuffer(): void
    {
        if ($this->rowBuffer === '') {
            return;
        }

        $rowsInBuffer = $this->rowBufferCount;
        $shouldSync = $this->randomAccessIndexEnabled
            && ($this->rowsSinceSync + $rowsInBuffer) >= $this->indexSyncPeriod;

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

        $this->rowsSinceSync = 0;
        $this->rowBuffer = '';
        $this->rowBufferCount = 0;
    }

    /**
     * Convenience accessor mirroring the path used in startNewSheet().
     */
    protected function currentSheetEntry(): string
    {
        return "xl/worksheets/sheet{$this->currentSheetIndex}.xml";
    }

    /**
     * Serialize the per-sheet sync points into the binary sidecar payload.
     * Sheets are written in workbook order so the sidecar's section order
     * matches the workbook.xml sheet listing.
     */
    protected function buildRandomAccessIndexPayload(): string
    {
        $sheetSections = [];
        foreach ($this->sheets as $sheet) {
            $entry = $sheet['filename'];
            $sheetSections[] = [
                'entry' => $entry,
                'total_rows' => $sheet['rows'],
                'sheet_crc32' => $sheet['crc32'] ?? 0,
                'sync_points' => $this->indexSyncPoints[$entry] ?? [],
            ];
        }

        return RandomAccessIndex::encode($this->indexSyncPeriod, $sheetSections);
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
            $headerRow = '<row r="1">';
            foreach ($this->columns as $i => $header) {
                $cellRef = $this->getColumnLetter($i + 1).'1';
                $escaped = $this->fastXmlEscape((string) $header);
                $headerRow .= '<c r="'.$cellRef.'"'.$headerStyleAttr.' t="inlineStr"><is><t>'.$escaped.'</t></is></c>';
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

        $sheetFooter = '</sheetData>'.$this->buildAutoFilterXml().'</worksheet>';

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
    }

    /**
     * Build row XML with optimization
     */
    protected function buildRowXml(int $rowIndex, array $data): string
    {
        $cells = [];
        $colIndex = 1;
        // Hoist the styles-empty check so the fast (unstyled) path skips the
        // per-cell lookup entirely — matches v2.0.1 hot-path cost.
        $hasColumnStyles = ! empty($this->columnStyleIds);

        foreach ($data as $value) {
            $cellRef = $this->getColumnLetter($colIndex) . $rowIndex;
            $col = $colIndex;
            $colIndex++;

            // Null / empty string -> empty cell
            if ($value === null || $value === '') {
                $cells[] = '<c r="' . $cellRef . '"/>';
                continue;
            }

            // Boolean -> native Excel boolean cell
            if (is_bool($value)) {
                $cells[] = '<c r="' . $cellRef . '" t="b"><v>' . ($value ? 1 : 0) . '</v></c>';
                continue;
            }

            // DateTimeInterface -> Excel serial date with datetime style
            // (column-specific format wins if set, else fallback to legacy STYLE_DATETIME)
            if ($value instanceof \DateTimeInterface) {
                $serial = ($value->getTimestamp() - self::EXCEL_EPOCH_TIMESTAMP) / 86400;
                $styleId = $hasColumnStyles && isset($this->columnStyleIds[$col])
                    ? $this->columnStyleIds[$col]
                    : self::STYLE_DATETIME;
                $cells[] = '<c r="' . $cellRef . '" s="' . $styleId . '" t="n"><v>' . $serial . '</v></c>';
                continue;
            }

            // Numeric values — split fast/slow so unstyled exports keep v2.0.1 cost
            if (is_int($value) || is_float($value)) {
                if ($hasColumnStyles && isset($this->columnStyleIds[$col])) {
                    $cells[] = '<c r="' . $cellRef . '" s="' . $this->columnStyleIds[$col] . '" t="n"><v>' . $value . '</v></c>';
                } else {
                    $cells[] = '<c r="' . $cellRef . '" t="n"><v>' . $value . '</v></c>';
                }
                continue;
            }

            // Numeric strings: preserve as string when precision/format would be lost
            if (is_string($value) && is_numeric($value)) {
                if ($this->shouldPreserveNumericString($value)) {
                    $escaped = $this->fastXmlEscape($value);
                    $cells[] = '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
                } elseif ($hasColumnStyles && isset($this->columnStyleIds[$col])) {
                    $cells[] = '<c r="' . $cellRef . '" s="' . $this->columnStyleIds[$col] . '" t="n"><v>' . (0 + $value) . '</v></c>';
                } else {
                    $cells[] = '<c r="' . $cellRef . '" t="n"><v>' . (0 + $value) . '</v></c>';
                }
                continue;
            }

            // String / Stringable / other -> inlineStr
            $escaped = $this->fastXmlEscape((string)$value);

            if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" ||
                $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                $cells[] = '<c r="' . $cellRef . '" t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
            } else {
                $cells[] = '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
            }
        }

        return '<row r="' . $rowIndex . '">' . implode('', $cells) . '</row>';
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
        static $trans = [
            '&' => '&amp;',
            '<' => '&lt;',
            '>' => '&gt;',
            '"' => '&quot;',
            "'" => '&apos;',
        ];

        // Double-quoted needle so \xNN escape sequences resolve to the
        // actual control bytes (0x00..0x1F minus \t \n \r). The previous
        // single-quoted form embedded the literal characters \, x, 0..9,
        // A..F instead, which let pure-lowercase strings with embedded
        // nulls bypass sanitization entirely — Excel rejects the workbook
        // because XML 1.0 forbids those bytes.
        if (strpbrk($str, "&<>\"'\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0B\x0C\x0E\x0F\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F") === false) {
            return $str;
        }

        $str = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F]/', '', $str);

        return strtr($str, $trans);
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
            throw XlsxStreamException::headersNotSet();
        }

        // Sample never reached its target — finalize with whatever we
        // collected so the preamble + header still get emitted.
        if ($this->inSampleMode && ! $this->autoWidthFinalized) {
            $this->finalizeAutoWidthSample();
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

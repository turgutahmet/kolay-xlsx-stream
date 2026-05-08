# Changelog

All notable changes to `kolay/xlsx-stream` will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.0.0] — 2026-05-06

The package becomes **bidirectional**: a streaming reader is added
alongside the existing writer, plus an opt-in random-access primitive
that turns XLSX into the first PHP-readable spreadsheet format with
O(1) row lookups. No breaking changes — every v2.x call site keeps
working untouched.

### Added — Streaming reader

- **`StreamingXlsxReader`** (`src/Readers/StreamingXlsxReader.php`):
  public facade for reading XLSX files. Factory methods:
  - `StreamingXlsxReader::fromFile($path)` — local filesystem
  - `StreamingXlsxReader::fromS3($s3, $bucket, $key)` — direct S3
    streaming via Range GET, no temp file
  - `StreamingXlsxReader::from(Source $source)` — bring your own source
- API:
  - `sheets()` — list all sheets in workbook order
  - `onSheet($name)` / `onSheetIndex($i)` — switch active sheet
  - `header()` — first row, cached
  - `rows($skip = 0, $limit = null)` — Generator over rows
  - `chunked($size, $skip = 0)` — Generator over fixed-size batches
    (designed for bulk DB inserts — amortises round-trip cost)
  - `rowCount()` — total rows; O(1) when an index sidecar is present,
    O(N) full scan otherwise
  - `rowAt(int $rowNumber)` — single row by 1-based row number; O(1)
    indexed, O(N) plain
  - `rowRange(int $from, int $to)` — Generator over inclusive row range
- **Bounded memory by construction** — peak RAM stays at 22-24 MB for
  every file size from 100 rows to 4.5 million. The reader never loads
  more than the longest in-progress row XML plus a 256 KB working buffer.
- **External XLSX support** — files written by PhpSpreadsheet, openpyxl,
  Apache POI, Excel itself, etc. read correctly. Shared-strings tables
  (`xl/sharedStrings.xml`) are loaded transparently when present.
  Tables above 20 MB compressed surface a clear error rather than
  silently consuming RAM (an on-disk variant is tracked for future).
- **Zero indirection on self-written files** — files written by
  `SinkableXlsxWriter` use inline strings exclusively, so the reader's
  fast path skips the shared-strings load entirely. This is the
  "born-bidirectional" round-trip pitch: write your data, read it back,
  same package, ~24 MB RAM, file-size-independent.

### Added — Random-access primitive

- **Writer opt-in: `withRandomAccessIndex(int $every = 10000)`** on
  `SinkableXlsxWriter` / `BaseXlsxWriter`. When enabled:
  - Periodic `ZLIB_FULL_FLUSH` is injected at row-buffer boundaries so
    the deflate stream gets byte-aligned `0x00 0x00 0xFF 0xFF` resume
    markers a downstream reader can fresh-init inflation from
    (no `inflatePrime()` needed — PHP's zlib doesn't expose it).
  - On `finishFile()` an `xl/_kxs/index.bin` ZIP entry is emitted
    listing each sync point's `(row_number, comp_offset, uncomp_offset)`
    triple, plus per-sheet total row counts. CRC32-validated.
  - The sidecar is OOXML-unreferenced — Excel, PhpSpreadsheet, OpenSpout,
    LibreOffice ignore it and open the file as a normal XLSX. Backward
    compatibility is total.
- **Reader auto-detection** — `StreamingXlsxReader` looks for the
  sidecar at construction; when present, `rowAt()`, `rowRange()` and
  `rowCount()` use it for O(1) seeks. When absent, the same calls fall
  back to a sequential scan with identical results.
- **Cost** (measured against the v3.0 benchmarks at 500K rows, sync
  period 10K):
  - Wall time: **+1.0 %** (within measurement noise)
  - File size: **+0.03 %** — 16× below the 0.5 % design budget
  - RAM: zero detectable overhead
- **Speedup**:
  - `rowAt(N)` up to **74.5× faster** than full scan at row 375K
  - `rowCount()` returns inside measurement noise vs **7.08 s** plain
    scan (effectively unbounded speedup — single CRC-validated header
    read)

### Added — Reader & writer ergonomics

- **`castColumn(int $col, string|callable $cast)` / `castColumns([])`** on
  `StreamingXlsxReader`. Built-in casts: `date`, `datetime`, `int`,
  `float`, `bool`. Pass a callable for custom transformations. Excel's
  1900 leap-year quirk is handled internally. **Datetimes return UTC by
  default** so the same file produces the same output on every server;
  override with `castTimezone($tz)`. Mac-origin 1904 epoch supported via
  `use1904Epoch()`.
- **`PhpStreamSink`** (`src/Sinks/PhpStreamSink.php`) — write a workbook
  directly to any PHP stream resource. Factories: `output()` for
  `php://output` (HTTP response streaming), `temp()` for `php://temp`,
  `memory()` for `php://memory`. Pairs with Laravel's `Response::stream()`
  for zero-temp-file XLSX downloads.
- **`setColumnFormat(int $col, int $builtinNumFmtId)` overload** —
  accept Excel's reserved-range built-in numFmtIds (0-49). Constants
  exposed on `BaseXlsxWriter`: `BUILTIN_NUMFMT_DATE`, `BUILTIN_NUMFMT_DATETIME`,
  `BUILTIN_NUMFMT_CURRENCY`, etc. Built-in ids render with the reader's
  locale (e.g. `dd.mm.yyyy` in tr-TR, `mm/dd/yyyy` in en-US) — useful
  when the same export ships to users in multiple regions.
- **`setAutoColumnWidth(sample: N, strict: false)`** — opt-in pass that
  scans the first N data rows and derives per-column width from the
  widest observed value. Catches columns where data is wider than the
  header. Lenient mode (`strict: false`, default) catches any internal
  failure, logs to `error_log`, and falls back to the heuristic — the
  user's export ships as a valid file with reasonable widths instead of
  HTTP 500. Strict mode propagates the exception (use during testing).
- **Slow-network protection in `StreamingSheetReader`** — empty-read
  retry counter with 10 ms backoff. After 100 consecutive empty reads
  with `feof === false` the reader throws `XlsxReadException::sourceUnreadable`
  rather than spinning indefinitely. Catches stalled S3/HTTP streams.
- **`ZipDirectory::dataOffset()` caching** — repeated random-access
  reads on the same entry now pay the 30-byte LFH range fetch only once.
- **`StreamingXlsxReader::__destruct`** — cheap deterministic cleanup
  for callers who don't explicitly call `close()`.

### Added — Documentation & benchmarks

- README hero refreshed: package described as bidirectional with
  random-access support, "Latest Benchmark" section now leads with v3.0
  read + random-access tables. v2.2.2 write benchmark moved down as
  "Write benchmark — v2.2.2", numbers unchanged. v1.x baseline kept
  for historical context.
- Quick Start sections added for **Reading XLSX Files** and
  **Random-Access Reading**, covering local + S3, header skipping,
  chunked batches, the writer-side opt-in, and `rowAt`/`rowRange`/
  `rowCount` usage.
- Comparison-with-other-libraries table extended with a Read column,
  Memory (Read) column, and Random Access column. Kolay XLSX Stream is
  the only package on the list with native read + zero disk + O(1)
  random access.
- New canonical benchmark scripts:
  - `benchmark-comprehensive.php` — write (existing, unchanged)
  - `benchmark-read.php` — sequential read, local + S3 cold cache
  - `benchmark-random-access.php` — `rowAt(N)` plain vs indexed, plus
    `rowCount()` comparison
- Test layout reorganized — `tests/Writers/` and `tests/Readers/`
  mirror `src/Writers/` and `src/Readers/`. The shared `TestCase` base
  stays at `tests/TestCase.php`. Existing v2.x writer tests moved into
  `tests/Writers/` with no behavioural change.

### Tests

- 95 new tests across the reader suite (foundation, cell tokenizer,
  streaming sheet reader, facade, multi-sheet resolver, shared strings,
  external XLSX round-trip, random-access decoder, random-access read).
- Test suite total: **194 tests, 485 assertions, all green**.
- Memory profile across the entire reader+writer suite stays at 38 MB
  peak — no inflation versus v2.2.2.

### Compatibility

- **No breaking changes.** Every public API from v2.x continues to
  work. Files produced by the v2.x writer are read by the v3.0 reader
  with zero overhead. Files produced by the v3.0 writer (without
  `withRandomAccessIndex()`) are byte-identical to v2.2.2 output.
- PHP 8.1+, Laravel 10/11/12/13 — same matrix as v2.2.2.
- AWS SDK PHP `^3.300` — same as v2.x. Required only when using
  `S3MultipartSink` (writer) or `StreamingXlsxReader::fromS3()` (reader).

## [2.2.2] — 2026-05-03

### Fixed

- **XML control-byte sanitization bypass** — `fastXmlEscape()` used a
  single-quoted needle for `strpbrk()`, so `\x00`, `\x01`, … escapes
  were embedded as the literal characters `\`, `x`, `0`, `1`, …
  instead of as actual control bytes. Strings whose characters didn't
  overlap with `&<>"'\x0..9A..F` (e.g. pure-lowercase Latin or
  multi-byte UTF-8 with embedded nulls) bypassed sanitization entirely
  and Excel rejected the workbook with `Char 0x0 out of allowed
  range`. The needle is now a double-quoted string so the escape
  sequences resolve correctly. Pre-existing bug going back to v1.x.

### Performance (side effect of the fix)

- The buggy needle was a 129-character literal that `strpbrk()` had to
  compare each input character against on the fast path. The corrected
  needle is 36 characters (5 special chars + 31 actual control bytes),
  so the per-cell sanitization check is ~3.5× cheaper. Comprehensive
  benchmark on the same machine and workload as the v2.2 baseline:
  - 1M rows local: **210K rows/s** (was 161K in v2.2 / 169K in v2.2.1)
    — **+30% over v2.2**, **+15% over v1.x baseline**
  - 4.5M rows S3: **130K rows/s** (was 107K in v2.2) — **+22%**
  - Memory unchanged — local stays at 0–2 MB constant, S3 sawtooth
    pattern unchanged.

  The README's "Latest Benchmark" table has been refreshed with the
  full v2.2.2 numbers and the "What changed since v1.x" notes flipped
  from "v2.x is ~5–10% slower locally" to "v2.x is now ~15–25% faster
  locally".

### Documentation

- README's `onProgress` section now notes that `$bytes` only advances
  when zlib emits compressed output. With small datasets (or large
  `setBufferFlushInterval()` relative to `setProgressInterval()`),
  consecutive events may report the same byte count between flushes.
  The row counter is always exact.

### Tests

- New `tests/XmlSanitizationTest.php` — 6 cases: pure-lowercase with
  null byte, full-control-byte input across the C0 set, mixed-case
  with control bytes, headers with control bytes, preservation of
  `\t` `\n` `\r` (valid XML chars), and the slow-path `&<>"'`
  escape sanity check. All assertions include a `libxml`-strict parse
  to match Excel's behavior.

## [2.2.1] — 2026-05-03

### Added — polish

- **Color hex validation** — `setHeaderStyle()` rejects anything that
  isn't a 6-character hex value (with or without a leading `#`) up
  front, instead of producing a styles.xml Excel silently rejects.
- **Custom font name** — `setHeaderStyle(['name' => 'Arial', ...])` now
  applies. The value is XML-escaped on the way out so font names
  containing `&` or `"` don't break the workbook.
- **Empty workbook is now a hard error** — calling `finishFile()` after
  `startFile()` with no rows and no `newSheet()` now throws
  `XlsxStreamException`. The previous behaviour produced `<sheets/>`
  which Excel and most readers refuse to open.
- **Out-of-range column format detection** —
  `setColumnFormat($col, $preset)` with `$col > count($headers)` no
  longer silently registers a format that's never applied. Validation
  is deferred to the next `startNewSheet()` (called from `writeRow()`'s
  auto-trigger or `newSheet()`) so callers can still pre-configure
  formats for an upcoming `newSheet()` whose column count is different
  from the current sheet's.

### Changed

- `BaseXlsxWriter::buildColsXml()` uses a plain `for` loop instead of
  `range()`/`foreach`, so it no longer allocates a 1..N integer array
  per sheet startup (matters at the 16,384 column limit).
- `StyleRegistry::resolveCellXf()` and `resolveFont()` use strict (`===`)
  comparison for dedup — same key order is guaranteed by the
  registration code paths, and strict comparison is both faster and
  semantically more correct.
- Auto-width minimum for `integer` and `decimal` formats raised to 14
  characters (was 10/12). The `#,##0` format adds thousand separators,
  so a 10-digit integer renders as `1,234,567,890` (13 characters) —
  the previous minimum was just below the threshold and Excel would
  render `####`. Same root cause as the v2.2.0 currency-width fix.

### Performance

- No measurable regression vs v2.2.0 baseline (1M rows local: 169K
  rows/s plain, 162K rows/s styled). Memory still O(1).

## [2.2.0] — 2026-05-03

### Added — styling & multi-sheet

- **`StyleRegistry`** — internal helper that emits `xl/styles.xml` on demand,
  deduplicates fonts, fills, number formats, and cellXfs by value so the
  same preset registered twice reuses the same style id.
- **`setHeaderStyle($options)`** — bold, fill color, text color, font size on
  the header row. Re-callable between `newSheet()` calls so each sheet can
  have its own header look.
- **`setColumnFormat($col, $presetOrCode)`** — named presets (`date`,
  `datetime`, `datetime_iso`, `time`, `integer`, `decimal`, `percent`,
  `currency_try`, `currency_usd`, `currency_eur`, `currency_gbp`) or any
  raw Excel format code. String cells are unaffected; numeric and DateTime
  cells in formatted columns are stamped with the column's style id.
- **`clearColumnFormats()`** — resets all column-level formats and widths.
  Useful between `newSheet()` calls when the next sheet has a different
  column layout.
- **`freezeFirstRow()` / `freezeRowsAndColumns(rows, columns)`** — pin the
  header row and/or the first N columns while scrolling.
- **`enableAutoFilter()`** — emit Excel's filter dropdowns on the header row
  with an automatically computed range.
- **`setColumnWidths([col => width])`** — explicit per-column widths.
- **`setAutoColumnWidth()`** — header-based heuristic with format awareness:
  the registered column format dictates a sensible minimum width (currency
  ≥ 14, datetime ≥ 20, date ≥ 12, percent ≥ 10) so values like ₺50,000.00
  no longer render as `####`.
- **`newSheet($name, $headers = null)`** — true multi-sheet workbooks with
  custom sheet names, not just the auto-split fallback at 1,048,576 rows.
  Optional `$headers` swaps the header row for the new sheet.

### Changed

- `xl/styles.xml` is now emitted at `finishFile()` instead of `startFile()`
  so styles registered between `newSheet()` calls are still captured.
- Configuration setters (`setHeaderStyle`, `setColumnFormat`,
  `setColumnWidths`, `setAutoColumnWidth`, `freezeRowsAndColumns`,
  `enableAutoFilter`) only refuse to run after `finishFile()` — they can
  be called any time before close, including between `newSheet()` calls.
- `BaseXlsxWriter::buildRowXml()` branches on a hoisted `$hasColumnStyles`
  flag so the unstyled hot path matches v2.1's per-cell cost (1M rows
  local: 161K rows/s plain vs 168K v2.1 baseline). Styled exports add
  ~2-3% throughput overhead and ~3% file-size overhead from the `s="N"`
  cell attributes.

### Performance vs v1.x baseline (May 2026 re-measurement)

- **S3 throughput**: ~2-3× faster across the board for ≥50K rows
  (1M rows: 95K rows/s vs 43K, 4.5M: 107K rows/s vs 46K). Most of the
  win comes from updated `aws/aws-sdk-php` (3.379+) and a faster network
  on the measurement machine; the multipart-upload code path is unchanged.
- **Local throughput**: ~5-10% slower than the September 2025 v1.x
  numbers, the cost of v2.0+ per-cell type detection (boolean cells,
  `DateTimeInterface` → serial date, big-integer-string preservation).
- **Memory**: unchanged — local stays at 0-2 MB constant, S3 keeps the
  same sawtooth pattern as the buffer fills and flushes per part.

## [2.1.0] — 2026-05-03

### Added
- **`SinkableXlsxWriter::forDisk($disk, $path)`** — Stream directly to any
  Laravel filesystem disk without manually constructing an `S3Client`. Reads
  credentials, region, bucket, and endpoint from
  `config('filesystems.disks.{$disk}')`. Supports `s3` and `local` drivers.
- **`onProgress(callable $cb)` + `setProgressInterval(int $rows)`** — Register
  a callback that fires every N rows with `(int $rowsWritten, int $bytesWritten)`.
  Null short-circuits when no callback is registered, so there is zero overhead
  for users who do not opt in. Useful for queue jobs that update a progress UI.
- **`writeRows(iterable)`** — `writeRows()` now accepts any iterable, not just
  arrays. Generators, `Iterator`, and `IteratorAggregate` all work directly:

  ```php
  $writer->writeRows(User::query()->lazy(1000));
  ```

  Backwards compatible — existing array callers are unaffected.

### Performance
- No measurable regression vs v2.0.1 (1M rows: 169K r/s vs 170K r/s baseline).
  All new features avoid the per-row hot path.

## [2.0.1] — 2026-05-03

### Changed
- CI: lower PHPStan to level 5; restrict Pint to `src/` so legacy
  `examples/`, `config/`, `tests/`, and `benchmark-comprehensive.php` are not
  retroactively flagged. Apply Pint autofix to `src/` (purely cosmetic — no
  runtime change).

## [2.0.0] — 2026-05-03

### Added
- **`\DateTimeInterface` support** — Passing a `DateTime` / `DateTimeImmutable` /
  `Carbon` instance now produces a real Excel date cell (numeric serial date with
  the built-in `yyyy-mm-dd hh:mm:ss` format). Previously this caused a fatal error.
- **Native Excel boolean cells** — `true` / `false` values now produce
  `t="b"` cells with `<v>1</v>` / `<v>0</v>`, matching the OOXML spec.
- **State machine guards** — Calling `writeRow()` before `startFile()`,
  calling `startFile()` twice, or calling `finishFile()` twice now throws
  a controlled `XlsxStreamException` instead of silently producing broken
  output or fatal `TypeError`.
- **Excel column limit enforcement** — `startFile()` and `writeRow()` reject
  more than 16,384 columns (Excel's hard limit) with a clear exception.
- **Orphan multipart upload visibility** — `S3MultipartSink::abort()` now
  emits an `error_log` warning when AWS rejects the abort, so orphan
  multipart uploads (which keep accruing S3 charges) are observable.
- **GitHub Actions CI matrix** — Tests run on PHP 8.1/8.2/8.3/8.4 against
  Laravel 10/11/12/13 with PHPStan and Pint checks.
- **`UPGRADE.md`** — Step-by-step migration guide from v1.x.

### Changed
- **Big numeric strings preserved** — Numeric strings with more than 15 digits
  (e.g. IBAN, IDs, card numbers) are now written as text instead of being
  cast to a float (which lost precision via scientific notation). Strings
  prefixed with `+` are also preserved.
- **PHP version requirement** — Minimum PHP raised from `^8.0` to `^8.1`.
  PHP 8.0 has been EOL since November 2023.
- **Laravel version requirement** — Now requires Laravel 10, 11, 12, or 13.
  Laravel 9 has been EOL since February 2024.
- **AWS SDK constraint tightened** — Now requires `aws/aws-sdk-php: ^3.300`
  for current API surface.
- **Destructors hardened** — `S3MultipartSink::__destruct()` and
  `FileSink::__destruct()` now catch `\Throwable` (not just `\Exception`)
  to prevent process crashes during shutdown.

### Fixed
- Calling `finishFile()` twice no longer crashes with
  `TypeError: hash_update(): Argument #1 ($context) must be a valid,
  non-finalized HashContext`.

### Removed
- Support for PHP 8.0 (EOL November 2023).
- Support for Laravel 9 (EOL February 2024).

## [1.0.2] — 2025-10-17

- Add Laravel 12 support and S3 integration tests.

## [1.0.1] — 2025-09-10

- Update author email in composer.json; add PHP version badge to README.

## [1.0.0] — 2025-09-10

- Initial release: high-performance XLSX streaming writer with FileSink and
  S3MultipartSink, multi-sheet support, and configurable compression.

[2.2.2]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v2.2.1...v2.2.2
[2.2.1]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v2.2.0...v2.2.1
[2.2.0]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v2.1.0...v2.2.0
[2.1.0]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v2.0.1...v2.1.0
[2.0.1]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v2.0.0...v2.0.1
[2.0.0]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v1.0.2...v2.0.0
[1.0.2]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/turgutahmet/kolay-xlsx-stream/releases/tag/v1.0.0

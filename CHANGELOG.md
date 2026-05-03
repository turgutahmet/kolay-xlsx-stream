# Changelog

All notable changes to `kolay/xlsx-stream` will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

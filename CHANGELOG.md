# Changelog

All notable changes to `kolay/xlsx-stream` will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[2.0.0]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v1.0.2...v2.0.0
[1.0.2]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/turgutahmet/kolay-xlsx-stream/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/turgutahmet/kolay-xlsx-stream/releases/tag/v1.0.0

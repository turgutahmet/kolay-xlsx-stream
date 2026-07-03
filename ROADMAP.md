# Roadmap

This document tracks where the package is heading. Items are best-effort —
priorities shift based on real-world feedback and contributions. Anything
listed below `Backlog` is "considered, not committed".

## Shipped

| Version | Date | Highlights |
|---|---|---|
| [3.1.0](CHANGELOG.md#310--2026-07-04) | 2026-07-04 | Queryable XLSX: KXSI TLV sidecar extension with per-block column stats (zone maps) — `withColumnStats()` on the writer; `columnStats()`/`rowsWhere()`/`findRow()` on the reader (Parquet-style block pruning, sidecar-only aggregates, single-block point lookups); `shards()`/`rowsForShard()` for zero-coordination queue fan-out; writer +55 % throughput (PCRE-JIT escape gate, flattened row builder, level-5 default); S3 reader open 7 → 3 round trips; `rowCount()` boundary-count fast path |
| [3.0.0](CHANGELOG.md#300--2026-05-06) | 2026-05-06 | Streaming reader (`StreamingXlsxReader`) with bounded ~24 MB RAM; born-indexed XLSX via opt-in `withRandomAccessIndex()` enabling O(1) `rowAt`/`rowRange`/`rowCount`; external XLSX support via shared-strings; new read + random-access benchmark scripts; tests reorganized under `Writers/` + `Readers/`; no breaking changes |
| [2.2.2](CHANGELOG.md#222--2026-05-03) | 2026-05-03 | XML control-byte sanitization fast-path bug fix (long-standing, pre-v2.x); `onProgress` byte-counter doc note |
| [2.2.1](CHANGELOG.md#221--2026-05-03) | 2026-05-03 | Hex color validation, custom font name, empty-workbook guard, deferred column-format check, wider integer/decimal auto-width minimums, micro-perf tightenings |
| [2.2.0](CHANGELOG.md#220--2026-05-03) | 2026-05-03 | Header styling, column number formats, freeze pane, auto-filter, manual + format-aware auto column widths, manual `newSheet()` |
| [2.1.0](CHANGELOG.md#210--2026-05-03) | 2026-05-03 | `forDisk()` Laravel Storage integration, `onProgress()` callback, `writeRows(iterable)` Generator support |
| [2.0.1](CHANGELOG.md#201--2026-05-03) | 2026-05-03 | CI / lint cleanup |
| [2.0.0](CHANGELOG.md#200--2026-05-03) | 2026-05-03 | DateTime support, native boolean cells, big-int preservation, state machine guards, modernized dependency matrix |

## Next: v3.2 — Reader/writer ergonomics & deferred polish

Additive items that didn't make v3.0/v3.1. **No breaking changes
planned.** (v3.1 pivoted to the queryable-XLSX layer — zone maps,
shards, hot-path perf — so this list carries over.)

### Reader

- **`xl/_kxs/index.bin` migration tool** — CLI command that walks a
  prefix on S3 (or a local directory), re-encodes each XLSX with
  `withRandomAccessIndex()` and uploads with `If-Match` for race
  safety. Lets fleets retroactively make existing data lakes
  random-access without re-running the original export job.
- **Strategy 2 — on-disk shared-strings index** — for archives where
  `xl/sharedStrings.xml` exceeds the 20 MB compressed in-memory
  threshold. Inflate to `php://temp`, build an offset map per `<si>`,
  resolve `t="s"` references via fseek + fread. Lifts the upper bound
  on external-XLSX read support without breaking the bounded-RAM
  contract.
- **`onProgress` callback for the reader** — symmetric to the writer's
  `onProgress`. Useful for long-running ingestion jobs that need to
  report row-count progress to a queue dashboard.

### Writer

- **Parallel S3 multipart upload** — currently parts are uploaded
  sequentially, RTT-bound. AWS SDK exposes async/concurrent upload
  primitives; running 3-4 parts in flight should land **2-3× S3
  throughput** on workloads above 1M rows. Tradeoff: peak memory grows
  to `partSize × concurrency` (≈ 96-128 MB at default 32 MB parts).
- **ZIP64 support for files > 4 GB** — current 32-bit ZIP offsets
  silently overflow. Most real workloads are well under this limit
  (4.5M rows ≈ 178 MB), but enterprise exports can hit it. Reader-side
  already detects ZIP64 sentinel values and refuses with a clear error.
- **Sample-based auto column width** — opt-in pass that scans the first
  N rows of data and computes per-column maximum widths, replacing the
  current header-text + format-min heuristic. Defaults stay heuristic
  to avoid the per-cell `mb_strlen` cost when not requested.
- **`setColumnFormat($col, int $builtinId)` overload** — accept Excel
  built-in `numFmtId` ints (e.g. `BUILTIN_NUMFMT_DATE = 14`) so callers
  with locale-sensitive formatting can lean on Excel's locale-aware
  built-ins instead of getting a hardcoded `yyyy-mm-dd` regardless of
  client locale.
- **`PhpStreamSink`** — write directly to `php://output` so HTTP
  responses can stream the workbook without a temp file. Pairs with
  Laravel's `Response::stream()`.

### Polish

- Skip `<col customWidth="1"/>` when the resolved width matches
  Excel's default (8.43). Saves a few bytes per default-width column;
  matters more at the 16,384-column ceiling than for typical reports.
- Lazy S3 client / multipart init — `S3MultipartSink::__construct`
  currently calls `createMultipartUpload` synchronously. Defer until
  the first `write()` so the sink is cheaper to instantiate in DI
  contexts (and so test fixtures can build one without S3 mocks).

### Queryable-XLSX follow-ups (v3.1 groundwork, PoC-verified)

The KXSI TLV framing reserves room for these; the deflate-level
mechanics (userland `crc32_combine`, full-flush segment concatenation)
were proven with runnable PoCs before v3.1 shipped:

- **Resumable S3 exports** — snapshot writer state at sync points
  (running CRC via `crc32_combine`); `Writer::resume($snapshot)` after
  a crash re-enters at row N+1 instead of restarting a 25-minute job.
- **Appendable XLSX** — end the last sheet at a full-flush boundary,
  reopen and continue with a fresh deflate context; on S3,
  `UploadPartCopy` makes append O(appended data). Wants ZIP64 first.
- **Distributed export + server-side stitch** — N queue workers each
  write a headless full-flush-terminated deflate segment; a finalizer
  assembles one valid born-indexed .xlsx via `UploadPartCopy` without
  downloading the segments (pigz-style concatenation).
- **Compact cell refs** (`c/@r` omitted, spec-legal) — measured 2.1×
  writer throughput and −44 % file size, but fast-excel-reader silently
  returns zero rows on such files, so this stays opt-in and gated on
  the external-reader compat matrix.
- **Retrofit index for foreign XLSX** — zran-style inflate-state
  snapshots make files *other* writers produced random-access after
  one pass. Pure PHP can't resume mid-stream (no `inflatePrime`);
  proven feasible via FFI → candidate for an optional
  `kolay/xlsx-stream-retrofit` package.

## Future: v4.0 — Considerations

Things that probably need a major version bump because they change
public behaviour or require a newer floor. **None of these are
committed yet** — they go in once the cost/benefit case is clear.

- **Cell-level styling API** — current API is column-level. Per-cell
  styling (e.g. red text for negative numbers) needs a different shape
  (`writeRow` would need to accept either a value or a styled cell
  wrapper). Worth designing carefully before promising.
- **Drop PHP 8.1** — will be ~2.5 years past EOL by typical v4 timing.
  Lets us use PHP 8.3+ features like `json_validate`, typed class
  constants, and `\WeakMap` improvements.
- **Drop Laravel 10** — EOL Feb 2025; will be ~3 years past EOL by v4.
- **`SinkableXlsxWriter` config object** — replace the chain of `setX()`
  setters with a strongly-typed configuration value object. Better for
  IDE autocomplete and validation, but breaks every existing call site.
- **Package split — `kolay/xlsx-stream-s3`** — move S3 sink + S3 source
  into a sub-package so callers who only use local files don't pull
  the AWS SDK transitively. Only worth doing if download volume grows
  and the AWS SDK install size becomes a real complaint.

## Backlog (considered, not committed)

Lower priority or feedback-dependent. May land in v2.x patches or v3,
or never.

- **Image embedding** — `$writer->insertImage($cellRef, $path, $sizing)`
  for logos/charts. Significant XML + relationship bookkeeping.
- **Hyperlink cells** — `new Hyperlink('https://...', 'Click')` value
  object. Needs worksheet rels.
- **Excel built-in style table extras** — borders, conditional
  formatting hooks (very limited subset, only what's safe for
  streaming).
- **`onError` callback** — companion to `onProgress` for sink/network
  errors during long S3 uploads (currently the writer aborts and
  rethrows; a callback would let users log + continue if appropriate).
- **CSV / TSV sink** — same `Sink` contract, different output. Trivial
  but adds a separate test surface; unclear it belongs in this package.
- **Sheet name uniqueness check** — `newSheet('Foo')` twice currently
  produces two sheets with the same name. Excel tolerates this; better
  if we'd auto-suffix or throw.
- **Performance: cell XML template caching** — pre-compute the
  per-type cell open/close fragments instead of `'<c r="' . ... . '">'`
  concat per cell. Marginal but measurable at 1M+ rows.

## How to suggest features

Open an issue with the use case (one paragraph is enough). Concrete
"my export does X, currently I have to Y" reports beat abstract
wishlists by a wide margin. Performance suggestions backed by a
benchmark land fastest.

PRs welcome — see [CONTRIBUTING](README.md#contributing). Tests +
benchmark gate are required for anything that touches the hot path.

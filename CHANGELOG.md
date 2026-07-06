# Changelog

All notable changes to `kolay/xlsx-stream` will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.3.0] — 2026-07-06

A backward-compatible minor: a real query engine on top of the born-indexed
sidecar, writer & DX ergonomics, integrity verification, and a critical S3
write-memory fix. No breaking API changes — the base `Source` contract is
unchanged and classic writer output stays byte-identical (compact mode and
group-aligned blocks are opt-in). Full suite 604 tests / 7541 assertions;
local write/read and the whole query surface measured at parity with 3.2.x
(no hot-path regression).

### Added — query engine

- **`rowsWhereAll(array $predicates)`** — multi-predicate AND with zone-map
  INTERSECTION. A row matching every predicate lives in a block surviving
  every predicate, so querying two differently-clustered columns reads far
  fewer blocks than either alone. Untracked columns filter per row.
- **`estimatedRows(col, op, value, value2?)` + `explain(array $predicates)`**
  — zero-I/O selectivity and query plans from the sidecar alone.
  `estimatedRows` returns a hard zone-map `upper` bound plus a t-digest/HLL
  `estimate`; `explain` returns `{strategy, candidateBlocks, runs,
  estimatedRows, estimatedBytes}` — the byte budget a pruned scan would fetch.
- **`topRows(col, k, desc)`** — "ORDER BY column [DESC] LIMIT k" from the
  sidecar. On a sorted column the k extremes are read from one end (one seek,
  early exit — an indexed top-N); on an unsorted column a single scan holds
  an O(k) heap.
- **`sampleRows(k, seed?)`** — a uniform random sample of k rows fetched via
  the index (≈k block reads, not a full scan). Seeded for reproducibility;
  unbiasedness proven by a chi-square goodness-of-fit gate.
- **Column addressing by header name** across the whole query surface —
  `columnStats('amount')`, `rowsWhere('date', …)`, `groupStats('region',
  'total')`, etc. Index addressing keeps its exact fast path.
- **`Readers\Bucket::year()/month()/day()`** — ready-made monotone bucket
  callables for `groupStats()` over a date column ("GROUP BY month" in one
  call), so the group-pure block pushdown still applies.
- **`verify()`** — read-side integrity check against the writer's per-block
  running-CRC (SCRC) pins and whole-sheet CRC. A single inflate pass at O(1)
  memory reports which block (if any) went bad — the "verified read" for
  audit/payroll/HR data. Falls back to the whole-entry ZIP CRC without a
  sidecar.
- **`onFullScan(callable)`** — observability hook fired when a query cannot
  push down and degrades to a full scan (unindexed column, unsorted groupBy),
  with `{query, column, reason, entry}` context.

### Added — writer & S3

- **`queryable(array $columns, every?, withSketches?)`** — one call that
  turns on the random-access index + zone maps + sketches for a set of
  columns; byte-identical to chaining the three explicit calls.
- **`compact()`** — opt-in r-less cell output (ECMA-376 makes `c/@r` and
  `row/@r` optional). Compressed sheets shrink ~52–62 %, and both writing and
  reading get faster (smaller XML to build and tokenize). The compact row
  builders are the exact r-less projection of the classic ones (byte-oracle
  test); classic output is byte-identical when off. Opens in Excel/
  LibreOffice/Numbers; read by PhpSpreadsheet 5.8 & OpenSpout 4.28.
- **`syncAtGroupBoundaries(column)`** — align index block boundaries to a
  column's group changes so every block holds one group; `groupStats()` then
  folds each block straight from the sidecar and reads only the header block.
  Enables the index if not already on; default off → output byte-identical.
- **Opt-in per-part integrity: `new S3MultipartSink(..., verifyParts: true)`**
  — attaches a `Content-MD5` to every part so S3 verifies it and rejects a
  corrupted part; a bad byte never enters the object ("verified export").
- **Bounded ranged reads (`Contracts\SupportsBoundedStream`)** — a pruned
  scan that stops at a sync boundary fetches only that block run's bytes; on
  S3 an open-ended `[offset, EOF]` GET becomes an exact ranged read. Optional
  gap-bridging keeps one ranged read alive across a short gap of non-matching
  blocks when transferring those bytes beats a round-trip (via an optional
  `Contracts\ProvidesCostHints` on the Source).

### Fixed

- **S3 multipart writes now use O(1) memory (was O(file size)).** The
  parallel (async) upload path leaked: the AWS SDK's promise/command graph
  held each dispatched part's body in a reference cycle refcounting cannot
  free, so write memory climbed with row count (~130 MB at 3M rows, ≈ the
  whole file — unbounded beyond), silently breaking the streaming-to-S3
  memory promise. Uploads are now **synchronous by default** and release
  each part as it lands, keeping memory flat at ~part size regardless of
  file size (verified ~30 MB peak on a 3M-row write). Local writes were
  always flat.

### Changed

- **`S3MultipartSink` default `concurrency` is now 1 (was 4).** This is the
  memory fix above: the default is strictly synchronous uploads (O(1) memory,
  steady throughput — no faster on bandwidth-bound links, and free of the
  leak). Parallel part uploads remain available as an opt-in (`concurrency >
  1` / `XLSX_STREAM_S3_CONCURRENCY`), now with per-part GC to bound their
  footprint — prefer them only where you have measured a win on a high-latency
  link. See UPGRADE.md.

## [3.2.2] — 2026-07-05

Correctness patch. No new public API. Row/cell bytes are unchanged
(pinned by test — every `s=` style id and cell body is byte-identical
to 3.2.1); the writer changes are constructor-time config wiring and
two `<cols>`/preamble metadata fixes, all below. The single-sheet read
path measures within ±1 % of 3.2.1 (A/B, five-run medians) — the
auto-split fix costs nothing unless a file actually has a continuation
chain.

### Fixed — auto-split workbooks now answer queries for the WHOLE table

When an export exceeds Excel's 1,048,576-row sheet ceiling the writer
auto-splits it across continuation sheets. Every query API silently
answered from the active sheet alone: on a 2.1M-row file `rowCount()`
reported 1,048,575, `findRow()` returned null for anything past the
first sheet, and `columnStats()` summed less than half the data —
silently. This release makes the whole query surface span the
continuation chain:

- **Spanned APIs:** `rowCount()`, `rows()`/`chunked()` (including the
  `skip` fast path), `rowAt()`, `rowRange()`, `rowsWhere()`,
  `findRow()`, `columnStats()`, `groupStats()`, `quantile()`/
  `median()`, `countDistinct()`, `shards()`.
- **Chain detection is deliberately conservative:** consecutive sheets
  chain only when every non-final member is exactly full (1,048,575
  rows — the writer's split point, a count no hand-built sheet hits by
  accident) AND every member carries an identical header row, AND the
  file is born-indexed (the sidecar provides the O(1) totals the
  detection needs). Intentional multi-sheet workbooks — different
  headers, or same-schema sheets that aren't full — keep their
  per-sheet semantics exactly as before. Single-sheet files take a
  `count($sheets) < 2` early-out; nothing changes for them.
- **Global row numbering:** a chain is ONE logical table. Continuation
  members' repeated header rows do not exist logically; data row *i*
  lives at global row *i + 1*. `rowsWhere()`/`findRow()`/`rowRange()`
  keys, `rowAt()` addressing and `rows(skip:)` all use this numbering.
- **Aggregates compose exactly:** `columnStats()` folds each member's
  block stats; `quantile()`/`countDistinct()` merge the members'
  t-digest/HLL sketches (that merge-associativity is why the sidecar
  stores them per sheet). All-or-nothing rule: if any chain member
  lacks the column's stats or sketch, the answer is `null` rather than
  a silent partial — the exact failure mode this release removes.
- **`shards()`** now fans out over every chain sheet (previously a
  4M-row file dispatched workers for only its first sheet).
  `rowsForShard()` keys remain LOCAL to the shard's sheet, as
  documented — the "skip `$rn === 1`" guidance now applies per member,
  since each member's first shard carries its (repeated) header.
- Behaviour note: selecting a chain member explicitly via
  `onSheet()`/`onSheetIndex()` still answers for the whole chain — a
  physical continuation sheet is not a meaningful query target. Sheets
  outside the chain are unaffected.

### Fixed — sample-mode auto width: config mutations no longer act retroactively

With sample-based auto column width (`setAutoColumnWidth(N)`), the
active sheet's preamble is deferred until the sample finalizes. Config
mutations issued while a sample was pending — the natural
"prepare-the-next-sheet, then `newSheet()`" sequence — leaked backwards
into that deferred preamble: `clearColumnFormats()` erased the pending
sheet's explicit `setColumnWidths()` entries (a user width of 40 lost
to a sampled 131), and `setHeaderStyle()` repainted the pending sheet's
header with the NEXT sheet's style. Reordering didn't help: calling
`clearColumnFormats()` after `newSheet()` instead tripped the
format-vs-header-count validation. The rule is now enforced in one
place: **a config mutation never acts retroactively** — mutators that
feed the preamble (`setHeaderStyle`, `setColumnFormat`,
`clearColumnFormats`, `setColumnWidths`, `setAutoColumnWidth`, freeze
panes) finalize a pending sample first, with the state it was sampled
under. Found by a real-Excel manual test; regression-tested red→green.

### Fixed — builtin numFmt ids now get sensible minimum column widths

Heuristic auto width (`setAutoColumnWidth(true)`) knew format-aware
minimum widths only for named format presets; columns formatted with
the `BUILTIN_NUMFMT_*` id constants (date, datetime, currency, …) fell
back to width 8 and rendered as `#######` in Excel. All thirteen
builtin ids now map to the same minimums their named counterparts use
(date 12, datetime 20, currency 14, …). Cell bytes are untouched —
this only changes the `<cols>` width hints of heuristic-width files.

### Fixed — the published config now actually applies (versioned)

`config/xlsx-stream.php` shipped `writer` / `s3` / `memory` / `logging`
keys (e.g. `XLSX_STREAM_COMPRESSION_LEVEL`) that no code ever read —
settings silently did nothing. As of 3.2.2 the config is wired in, with
two safety properties:

- **Version gate:** the package applies the config only when it carries
  `'version' => 2` (present in the new published file). Copies
  published before 3.2.2 lack the key and stay inert — exactly their
  previous behaviour — because their stale defaults contradict the code
  (old file said compression level 1; the writer's real default has
  been 5 since v3.1) and honouring them would have silently changed
  output for existing installs. Re-publish with
  `php artisan vendor:publish --tag=xlsx-stream-config --force` to opt
  in.
- **Precedence:** code-level setters always override config values,
  which override package defaults. Invalid config values are ignored
  (environment data must not turn every `new Writer` into a 500);
  setters keep throwing on invalid input as before.

Applied keys map 1:1 to real knobs: `writer.compression_level`,
`writer.buffer_flush_interval`, `s3.part_size` and `s3.concurrency`
(the `forDisk()` defaults). The old `memory` / `logging` /
`max_rows_per_sheet` keys had no implementation behind them and are
gone — a config key that silently does nothing is the exact bug this
release fixes. Two of the old keys have real homes elsewhere,
documented in the config file: transient S3 retries belong to the AWS
SDK's retry middleware (`'retries' => N` on your `S3Client`; the sink
adds one last-resort re-dispatch), and "logging" is the `onProgress()`
callback. Outside Laravel nothing changes (`config()` absence is
guarded).

## [3.2.1] — 2026-07-04

Documentation-only patch. The v3.2.0 tag was published with this
changelog section still reading "Unreleased"; Packagist's stable-version
immutability (correctly) blocked amending the published tag, so the
corrections ship as a patch release. No code changes — the 3.2.0 and
3.2.1 trees differ only in documentation.

- CHANGELOG: 3.2.0 release date stamped.
- ROADMAP: v3.2.0 moved to Shipped; long-shipped items removed from the
  "next" list (`PhpStreamSink`, sample-based auto width, built-in
  numFmtId overload, parallel multipart, shared-strings ceiling); next
  release scoped as v3.3 — resumable exports & smarter S3 reads.

## [3.2.0] — 2026-07-04

### Added — KXSI is now an open specification

- **`SPEC.md`**: normative binary layout, TLV contract (unknown tags MUST
  be skipped; the version byte never bumps for additive sections), the
  two hard-won soundness laws ("stats widen, never narrow"; "block stats
  MUST account for every row the scan path can yield, including
  preamble-emitted rows"), must-understand `flags` bitmask semantics
  (now enforced by the reference decoder), security bounds, and ten
  reserved section tags. Conformance suite: 5 committed fixture
  workbooks under `tests/SpecVectors/` with sidecar hexdumps and JSON
  goldens, byte-pinned in CI.
- **`SCRC` TLV section**: running CRC32 of the sheet's uncompressed
  bytes at every sync point, emitted whenever the index is enabled
  (capture cost ≈ 69 ns per sync point). Enables truncation detection
  today and is the shared prerequisite for resumable exports,
  appendable workbooks, and signing. Reader accessor:
  `syncPointCrcs($entry)`.
- **`groupStats(int $groupBy, int $aggregate, ?callable $bucket)`**:
  sorted-group aggregate pushdown — group-pure interior blocks
  contribute their precomputed sidecar sums without being read; only
  boundary blocks are inflated. 20 groups over 1M rows in ~57 ms with
  interior blocks provably never fetched (spy-tested); degrades to an
  honest full scan on unsorted/untracked columns.
- **`withColumnSketches(array $columns)`** (writer) + **`quantile()` /
  `median()` / `countDistinct()`** (reader): file-level approximate
  statistics answered from the sidecar alone. Per sheet, per tracked
  column, the writer embeds a merging t-digest (`TDIG` TLV section,
  δ=100, ~1-4 KB) and a HyperLogLog (`CHLL`, p=11, 2051 bytes) — the
  reader then answers "what is the p99 amount" or "how many distinct
  emails" for a multi-GB file with **zero row reads and zero
  additional range requests** after open (spy-pinned: no `range()`, no
  `streamFrom()`). Measured error on 100K samples: p01/p99 within
  0.04 % rank error and p50 within 0.05 % on continuous distributions
  (pinned at 0.2 %/0.5 % in CI; heavy-duplicate mid-range pinned at
  4.5 % — interpolation between duplicate masses); distinct counts
  within 3 % at 100K cardinality, near-exact below ~1K (linear-counting
  range). Both sketches merge associatively (HLL merge is exactly the
  union sketch) — the property future segment/partition stitching
  builds on. Orthogonal to `withColumnStats()`; implies
  `withRandomAccessIndex()`. Header row deliberately excluded from
  sketches (estimation bias vs `STAT`'s pruning-soundness header fold —
  the asymmetry is documented in SPEC.md §4.3). Warm quantile calls run
  in ~1.4 µs, countDistinct in ~62 µs (1M-row fixture, 4 MB peak).
  Write-side cost is ~0.4 µs per sketched cell after batch-fold
  optimization: ~10 % wall time with 2 tracked columns on a realistic
  12-column 500K-row export, ~45 % on the deliberately minimal 4-column
  hot-path bench whose baseline is 0.48 µs/cell.
- **`autoDetectDates(bool $withTime = true)`** (reader, opt-in):
  parses `xl/styles.xml` once into a styleId→isDate bitmap (builtin ids
  14–22/45–47 plus a quote/bracket/escape-aware format-code scanner) and
  yields `DateTimeImmutable` for date-styled numeric cells in external
  files — the "why is my date 45123?" fix. Explicit `castColumn` wins;
  query predicates keep matching raw serials; the default read path is
  measurably unchanged.

### Performance

- **Tokenizer micro-pass: +30 % read throughput** (isolated A/B vs
  v3.1.0 on a 13K real-row corpus: 106K → 138K rows/s): `strcspn`-based
  tag-end scanning, single-pass `r`/`t`(/`s`) attribute extraction,
  dense-row fast path, plain-`<is><t>` fast path, and namespaced hot
  functions imported via `use function` (the namespace-fallback lookup
  alone was ~5-8 %). Same treatment in `SharedStringsParser`.
- **`rows(skip: N)` fast path**: with an index, the skip seeks through
  the same sync-point machinery as `rowAt` (measured skip=1M:
  3.94 s → 2.5 ms); without one, a `</row>` boundary scan replaces
  tokenized skipping (~29× on skip=400K).
- **Within-block fast-forward** for `rowAt`/`rowRange`/`rowsWhere`/
  `rowsForShard`: rows between a sync point and the target are
  boundary-scanned instead of tokenized — random `rowAt` at 10K sync
  period measured 21.1 → 1.13 ms/call (~19×).
- **Parallel S3 multipart upload**: parts now upload through a bounded
  async window (`concurrency` constructor param, default 4;
  `uploadPartAsync` + FIFO waits, parts re-sorted before completion,
  first-error abort with in-flight settlement). Peak memory is capped
  at ~`partSize × (concurrency + 2)` and the former ~40 MB-per-million-
  rows ratchet is gone (measured 46.5 MB flat at 1M rows). On an
  RTT-bound link the window projects ~2×; on a single-connection-
  saturated uplink the measured win is stability (7.8–9.9 s steady vs
  10.5–30.5 s swings sequentially). `concurrency: 1` reproduces the
  v3.1 request sequence exactly.
- **Packed shared strings + streaming SST parse**: the shared-strings
  table now parses incrementally from the inflate loop (the full XML
  never exists in memory) into a packed offset-indexed buffer — 1M-entry
  / 30 MB table: 83.7 → 24.0 MB peak (3.5×). Support thresholds raised
  accordingly: compressed 20 → 64 MB, uncompressed 100 → 320 MB; the
  forged-metadata guard now trips at the first byte past the declared
  size with O(1) memory.
- `blockRanges()` memoized per sheet on the query path.

### Benchmarks

- New `bench/read_bench.php`, `bench/query_bench.php`,
  `bench/sst_bench.php` (same fresh-process/median hygiene as the write
  bench). v3.2 baselines recorded in `bench/results.json`: full-scan
  126.9K rows/s (6-col) at 6 MB peak; cold open 1.07 ms; rowAt
  1.28 ms; findRow 26.5 ms; groupStats 56.9 ms over 1M rows.

## [3.1.0] — 2026-07-04

### Added — Queryable XLSX (KXSI "STAT" sidecar extension)

- **`withColumnStats(array $columns)`** on the writer: track per-block
  min/max/sum/count statistics (zone maps) for chosen 1-based columns,
  plus a per-sheet sorted-ascending/descending verdict, embedded in the
  random-access sidecar. Implies `withRandomAccessIndex()`. Cost: a few
  comparisons per tracked cell and ~13 KB of sidecar per column at
  4M rows / 10K sync period.
- **TLV extension framing**: optional tagged sections (`STAT` today)
  appended after the v2 core body — the version byte stays 2 on
  purpose. The v3.0.x decoder parses exactly `sheet_count` core
  sections and ignores trailing bytes, so **old readers keep full
  random access on stats-bearing files** (they just can't see the
  stats), and future sections ship without a version bump that would
  strand older readers. Stats-less files remain byte-identical to
  v3.0.x output.
- **`columnStats(int $column)`** on the reader: whole-sheet
  min/max/sum/avg/count + sortedness answered from the sidecar alone —
  on S3, one range request against a multi-GB file.
- **`rowsWhere(int $column, string $op, $value, $value2 = null)`**:
  numeric predicate scan (`=`, `<`, `<=`, `>`, `>=`, `between`) that
  uses zone maps to skip every block whose `[min,max]` provably cannot
  match — Parquet-style row-group pruning inside a spreadsheet Excel
  still opens normally. Degrades to a full-scan filter (same results)
  when stats are absent.
- **`findRow(int $column, $value)`**: point lookup; on a column the
  writer observed to be sorted this reads a single block — two S3
  round trips end to end.
- **`shards(int $count)` / `rowsForShard(array $shard)`**: split a
  born-indexed sheet into independently decompressible, JSON-serializable
  row ranges for queue fan-out — N workers each stream 1/N of the file
  with zero coordination. Non-indexed files degrade to a single
  whole-sheet shard.
- **`SupportsSuffixRange`** optional Source capability (implemented by
  `LocalFileSource` and `S3RangeSource`) — fetch the ZIP tail and learn
  the object size in one request.

### Performance

- **Writer +55 % throughput** (500K-row mixed bench, apples-to-apples
  with generation cost excluded: median 2.68s → 1.73s, ~289K rows/s,
  peak RAM unchanged at 6 MB):
  - `fastXmlEscape` gate switched from a 34-char `strpbrk` needle to a
    PCRE-JIT character class (~5× cheaper on clean strings, which are
    the overwhelming majority);
  - `buildRowXml`/`buildStyledRowXml` flattened — string-first type
    dispatch, column-letter cache hoisted to a local, direct string
    append instead of array + implode. Sheet XML verified
    **byte-identical** against v3.0.2 on an edge-case corpus;
  - default deflate level 6 → 5 — measured knee of the curve for
    XLSX-shaped XML: ~0.2 % larger file, ~20 % less wall time. Call
    `setCompressionLevel(6)` to restore the old bytes.
- **Reader open on S3: 7 → 3 round trips** — single suffix-range GET
  replaces HEAD + tail GET, LFH + body fetches coalesced for small
  entries (≤ 1 MB), `xl/workbook.xml` fetched + inflated once instead
  of twice (it fed both sheet resolution and date1904 detection).
- **`rowCount()` fast path** without an index: counts `</row>`
  boundaries on the inflated stream instead of tokenizing every cell
  (5–10× faster), skips shared-strings resolution entirely (and its S3
  fetch), and no longer refuses files whose SST exceeds the RAM
  thresholds — counting never needed it.
- **`castDate` timezone cache** — one `DateTimeZone` construction per
  configuration instead of per date cell (2–3× faster date casting).
- EOCD scan uses `strrpos` instead of a byte-wise backward loop;
  `S3RangeSource::range()` guards zero/negative lengths.

### Fixed

- **`rowsForShard()` implicit sheet switch left stale per-sheet state.**
  A shard addressing a different worksheet switched the reader's
  current entry without the resets `onSheet()`/`onSheetIndex()`
  perform: `header()` kept returning the previous sheet's cached
  header, and `castColumn()` registrations from the previous sheet
  were applied to the new sheet's rows. The implicit switch now runs
  the same resets. Found by adversarial audit with reproduction.
- **Preset-name guard rejected valid pure-lowercase format codes.**
  Date-token runs like `'dddd'` (weekday) or `'mmss'` are legitimate
  Excel format codes but matched the typo-guard's preset-name shape.
  Strings built only from the `d/m/y/h/s` token letters now pass, and
  `setColumnFormat($col, $code, raw: true)` skips preset resolution
  and the guard entirely — the escape hatch for arbitrary codes from
  external constant tables (e.g. PhpSpreadsheet `NumberFormat`).
- **Zone-map pruning could hide a numeric-looking header row.** The
  header is emitted via the sheet preamble, not `writeRow()`, so it
  wasn't folded into block 0's stats — `rowsWhere()`/`findRow()` for a
  value matching the header but outside the data range returned the
  header on the full-scan path and nothing on the pruned path. The
  writer now folds header cells into block 0 (min/max/sum/count; the
  sortedness verdict deliberately stays data-only), restoring the
  "stats widen, never narrow" invariant. Found by adversarial oracle
  testing; regression-tested against both paths.
- **`rowsWhere()` fallback path yielded 0-based keys.** Without stats
  the full-scan filter iterated a key-stripping wrapper, so the same
  query yielded sequential 0-based keys instead of the documented
  1-based sheet row numbers the pruned path emits. Both paths now
  yield identical keys.
- **`findRow()` docblock overstated its mechanism.** It performs
  zone-map pruning with first-match early exit, not a binary search;
  on sorted columns the prune alone bounds the lookup to typically one
  block, so the cost claim was right and the wording now matches the
  implementation. The sidecar's sorted flag remains informational
  (surfaced via `columnStats()['sorted']`).
- **Unknown format-preset names now throw at write time.**
  `setColumnFormat($col, 'currency')` — a typo'd or guessed preset
  name — used to fall through as a literal `formatCode="currency"`,
  producing a styles.xml that MS Excel refuses to open without repair
  (verified against real Excel). A lowercase-letters string that isn't
  a known preset is rejected with the preset list in the message; raw
  Excel format codes (`"#,##0.000"` etc.) are unaffected.

### Benchmarks

- `bench/row_style_bench.php` now measures row-generation cost in a
  calibration pass and subtracts it (previous published numbers
  understated writer throughput by ~10–13 %); DateTime test data comes
  from a pre-built pool. `bench/results.json` carries re-measured
  v3.0.2 and v3.1.0 numbers side by side.

## [3.0.2] — 2026-06-26

### Added — Per-row styling

- **`registerRowStyle(array $options): int`** on
  `SinkableXlsxWriter` / `BaseXlsxWriter`. Registers a reusable row
  style (background fill + text color, plus `bold`/`size`/`name`) and
  returns a stable style id. Options are identical to
  `setHeaderStyle()`: `fill` (`#RRGGBB`), `color` (`#RRGGBB`), `bold`,
  `size`, `name`.
- **`writeRow(array $row, ?int $styleId = null)`** — optional second
  argument stamps a registered style onto every cell of that row, so
  callers can highlight selected rows (failed red, VIP blue, etc.) with
  both background and text color at once. `null` (the default) keeps the
  row on the unstyled fast path.

  ```php
  $failed = $writer->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
  foreach ($rows as $r) {
      $writer->writeRow($r->toArray(), $r->failed ? $failed : null);
  }
  ```

### Performance & guarantees

- **No regression on the unstyled path.** A single null check delegates
  the styled case to a separate builder, so the hot path stays
  byte-for-byte identical to v3.0.1. Benchmark (500k rows, 10 runs,
  median): unstyled unchanged (~3.10s), ~1% of rows styled is within
  noise, every row styled costs ~1%. Peak memory stays flat at 6 MB in
  all cases.
- **Flat memory via dedup.** Styles are deduplicated in the registry —
  one logical style is a single `styles.xml` entry no matter how many
  rows use it. Painting millions of rows adds no per-row memory.
- **Column number formats preserved.** A styled row still honours a
  column's number format: a currency/date column keeps its format and
  gains the row's fill/color via an on-the-fly merged cell style
  (memoized per row-style/column pair).
- **Empty cells in a styled row are still filled** so the highlight is
  visually contiguous.

### Compatibility

- Fully backward compatible — existing `writeRow(array $row)` calls work
  unchanged. The reader is unaffected (it is value-only and ignores the
  cell style attribute, which the writer already emitted for header
  styles and column formats). Output remains compatible with Excel,
  PhpSpreadsheet, and OpenSpout.

## [3.0.1] — 2026-05-14

### Changed

- **Widened `aws/aws-sdk-php` constraint** from `^3.300` to `^3.180`. The
  package only touches stable S3 surface (`S3Client`, `getObject`,
  `createMultipartUpload`, `uploadPart`, `completeMultipartUpload`,
  `abortMultipartUpload`, `AwsException`) — all of which have been
  present since the early 3.x line. `^3.300` was unnecessarily
  restrictive and blocked projects whose `composer.lock` had pinned
  an older but still-current aws-sdk-php release (e.g. 3.288.x).

## [3.0.0] — 2026-05-13

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
- Plus **74 hardening tests** added during the RC cycle (ZIP32 guards,
  STORED-method worksheets, forged sst metadata, max-column boundary,
  stale-index detection, builtin numFmtId overload, sample byte cap,
  PhpStreamSink, OOXML compliance for indexed XLSX).
- Test suite total: **268 tests, all green**.
- Memory profile across the entire reader+writer suite stays under 40 MB
  peak — no inflation versus v2.2.2.

### Pre-RC hardening

Real bugs caught during three external review cycles before the v3.0-rc.1
tag (2026-05-09) and addressed before the final tag:

- **Stale-index detection** — when an external editor (Excel, OpenSpout)
  rewrites a sheet, the embedded sheet CRC32 cross-check trips and the
  reader silently falls back to a sequential scan. Verified by
  `RandomAccessIndexWriterTest::stale_rewritten_sheet_falls_back`.
- **Excel repair mode fix** — `[Content_Types].xml` now declares an
  Override for the `xl/_kxs/index.bin` sidecar
  (`application/octet-stream`). Without it Excel flagged the file as
  repairable on first open.
- **Forged sst metadata defense** — three-stage guard chain on the
  shared-strings entry: compressed-size limit (20 MB), uncompressed-size
  limit (100 MB), and a post-inflate `strlen()` check that catches
  forged CD metadata.
- **ZIP32 limit guards** — writer rejects archives exceeding 4 GB or
  65,535 entries with a clear error instead of silent truncation.
- **Reader max row XML cap** — 16 MB ceiling on any single `<row>`
  element guards against malicious sparse rows.
- **Auto-width sample byte cap** — 8 MB buffer ceiling prevents pathological
  payloads from forcing unbounded sample collection; switches to
  heuristic when reached.
- **64-bit PHP guard** — `RandomAccessIndex::decode()` checks `PHP_INT_SIZE`
  before `unpack('P')` to fail loudly on 32-bit builds.
- **CellTokenizer max column guard** — column references past Excel's
  `XFD` (16,383) are rejected at parse time instead of overflowing into
  silent bad reads.
- **Random-access index semantic validation** — monotonic sync-point
  offsets, duplicate-sheet detection, path-length cap, trailing-byte
  rejection, comp_offset bounded by compressed size from the central
  directory.
- **STORED-method worksheets** — reader handles entries with compression
  method 0 (uncompressed) alongside DEFLATE for tooling
  (PhpSpreadsheet, certain Java exporters) that doesn't always compress
  worksheet entries.
- **`use1904Epoch()` auto-detect** — `xl/workbook.xml`'s
  `workbookPr/@date1904` attribute is read at workbook open; callers
  on Mac-origin files no longer need to call the setter manually.
- **Reader column casts reset** — switching active sheet via `onSheet()`
  / `onSheetIndex()` clears column casts; the previous behaviour leaked
  casts across sheets.
- **Slow-network protection** — reader breaks out of empty-read loops
  after 100 retries × 10 ms and surfaces `XlsxReadException::sourceUnreadable`
  instead of spinning.

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

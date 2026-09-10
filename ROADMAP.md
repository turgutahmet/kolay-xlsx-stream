# Roadmap

This document tracks where the package is heading. Items are best-effort —
priorities shift based on real-world feedback and contributions. Anything
listed below `Backlog` is "considered, not committed".

## Shipped

| Version | Date | Highlights |
|---|---|---|
| [3.4.0](CHANGELOG.md#340--2026-08-02) | 2026-08-02 | The spreadsheet that profiles itself. **`profile()`** — a full per-column report (numeric/empty counts, min/max/avg, certified percentiles, histogram, distinct, top values, correlations) from the sidecar, reading no data rows. **Exact + certified quantiles:** `exactQuantile()` (Rank-Sandwich — the STAT × TDIG certificate brackets the value, then only the blocks that can hold it are read; a sorted column answers from one row), `explainQuantile()` (zero-I/O plan + rank certificate). **New sketches (five additive registered TLV sections):** `correlation()` exact Pearson (`CORR`, centred Welford/Chan moments — accurate on timestamp columns), `topValues()` frequent items (`TOPK`), `argMin()`/`argMax()` extreme-row pointers (`ARGP`), `quantile(col,q,from,to)` + `groupQuantile()` range/group percentiles (`TDGB`), string `rowsWhere()`/`findRow()` (`STRZ`, normative unsigned-UTF-8 collation — Turkish İ/ı is byte order, not locale). **Scan:** `histogram()` (equi-width / equi-depth), `countEmpty()`, late-materialization (probe-then-tokenize, planner-gated, byte-identical). KXSI format byte stays `2`; classic writer output byte-identical; SPEC doc 1.8.1 |
| [3.3.0](CHANGELOG.md#330--2026-07-06) | 2026-07-06 | The query engine grows up + integrity + O(1) S3 writes. **Query:** `rowsWhereAll()` (multi-predicate AND via zone-map intersection), `estimatedRows()`/`explain()` (zero-I/O plans), `topRows()` (indexed ORDER BY … LIMIT), `sampleRows()` (seeded uniform sample), column addressing by header name, `Bucket::month/day/year` for `groupStats`, bounded ranged reads + gap-bridging. **Writer/DX:** `queryable()` one-call preset, opt-in `compact()` (r-less cells, ~52–62% smaller sheets), `syncAtGroupBoundaries()` (zero-scan `groupStats`), `onFullScan()` hook. **Integrity:** `verify()` (block-granularity CRC report), opt-in per-part `Content-MD5` so S3 rejects a corrupted part. **Fixed:** S3 multipart writes are now O(1) memory (was O(file size)) — default `concurrency` 1, parallel opt-in. No breaking changes (base `Source` contract stable; classic output byte-identical) |
| [3.2.2](CHANGELOG.md#322--2026-07-05) | 2026-07-05 | Correctness patch: on auto-split workbooks (>1,048,575 rows) the entire query surface — `rowCount`/`rows`/`rowAt`/`rowRange`/`rowsWhere`/`findRow`/`columnStats`/`groupStats`/`quantile`/`countDistinct`/`shards` — now spans the continuation chain as one logical table instead of silently answering from the active sheet alone; misleading never-read config keys removed |
| [3.2.0](CHANGELOG.md#320--2026-07-04) | 2026-07-04 | KXSI becomes an **open specification** ([SPEC.md](SPEC.md)) with a byte-pinned conformance suite; `SCRC` per-sync-point integrity CRCs; approximate analytics — `withColumnSketches()` embeds t-digest + HyperLogLog, `quantile()`/`median()`/`countDistinct()` answer with zero row reads; `groupStats()` sorted-group pushdown; +30 % read throughput (tokenizer micro-pass); `rows(skip)` fast path (~1,580× indexed) + within-block fast-forward (~19×); parallel S3 multipart upload (flat memory, steady wall times); packed shared strings (ceiling 20 → 64 MB compressed at 3.5× less peak); `autoDetectDates()` for external files |
| [3.1.0](CHANGELOG.md#310--2026-07-04) | 2026-07-04 | Queryable XLSX: KXSI TLV sidecar extension with per-block column stats (zone maps) — `withColumnStats()` on the writer; `columnStats()`/`rowsWhere()`/`findRow()` on the reader (Parquet-style block pruning, sidecar-only aggregates, single-block point lookups); `shards()`/`rowsForShard()` for zero-coordination queue fan-out; writer +55 % throughput (PCRE-JIT escape gate, flattened row builder, level-5 default); S3 reader open 7 → 3 round trips; `rowCount()` boundary-count fast path |
| [3.0.0](CHANGELOG.md#300--2026-05-06) | 2026-05-06 | Streaming reader (`StreamingXlsxReader`) with bounded ~24 MB RAM; born-indexed XLSX via opt-in `withRandomAccessIndex()` enabling O(1) `rowAt`/`rowRange`/`rowCount`; external XLSX support via shared-strings; new read + random-access benchmark scripts; tests reorganized under `Writers/` + `Readers/`; no breaking changes |
| [2.2.2](CHANGELOG.md#222--2026-05-03) | 2026-05-03 | XML control-byte sanitization fast-path bug fix (long-standing, pre-v2.x); `onProgress` byte-counter doc note |
| [2.2.1](CHANGELOG.md#221--2026-05-03) | 2026-05-03 | Hex color validation, custom font name, empty-workbook guard, deferred column-format check, wider integer/decimal auto-width minimums, micro-perf tightenings |
| [2.2.0](CHANGELOG.md#220--2026-05-03) | 2026-05-03 | Header styling, column number formats, freeze pane, auto-filter, manual + format-aware auto column widths, manual `newSheet()` |
| [2.1.0](CHANGELOG.md#210--2026-05-03) | 2026-05-03 | `forDisk()` Laravel Storage integration, `onProgress()` callback, `writeRows(iterable)` Generator support |
| [2.0.1](CHANGELOG.md#201--2026-05-03) | 2026-05-03 | CI / lint cleanup |
| [2.0.0](CHANGELOG.md#200--2026-05-03) | 2026-05-03 | DateTime support, native boolean cells, big-int preservation, state machine guards, modernized dependency matrix |

## Next: v3.5 — Template mode

Additive, no breaking changes; the KXSI format, [SPEC.md](SPEC.md) and the
`Source` contract are untouched (format byte stays `2`).

The expensive part of an .xlsx is not its layout — it is its rows. Template
mode takes the layout (headers, merges, the style table, column widths, freeze
panes, rich text) from a workbook **another producer made** — PhpSpreadsheet,
Excel, or xlsx-stream itself — and streams the rows into it with the streaming
writer. The template doubles as a **style-id oracle**: a sample row's `s="…"`
map tells the writer how that producer encoded each style, so no styling API
has to be reinvented.

```php
$template = Template::open('/tmp/layout.xlsx');
$writer = SinkableXlsxWriter::fromTemplate($sink, $template);
$writer->sheet('Leaves', dataStartRow: 3);
foreach ($rows as $i => $row) {
    $writer->writeRow($row, variant: $i % 2);   // zebra, styled by the template
}
$writer->finishFile();   // untouched sheets + static parts copied byte-identical
```

Measured on a real report (8,000 × 10 cells): **17× faster, 14× less memory**
than building the same file with PhpSpreadsheet. Ratios move with the
workload — PhpSpreadsheet's per-row style cost is not linear — so every
figure is reported with the load it came from in
[BENCHMARK.md §5](BENCHMARK.md) rather than reduced to one headline number.

Parity is the acceptance criterion, not the speed: the same layout written
entirely by PhpSpreadsheet versus written by PhpSpreadsheet and streamed into,
both read back through PhpSpreadsheet, over 400 rows × 6 columns shows **zero
differences** in value, type, font, fill, border, alignment, number format,
merge, width, height, freeze and gridlines.

Everything below the sample rows is copied verbatim, so a range drawn over
them — an auto filter, a conditional format, a merge — is **refused** at open
time rather than carried across to cover only the rows the template declared.
A template's filter belongs on the header row.

The classic writer is unaffected: template mode is a separate builder family
behind one boolean, and existing golden outputs stay byte-identical. Everything
outside the four-point seam (`<sheetData>` cut, shared-strings append, optional
`styles.xml` append, one `[Content_Types]` override) is copied byte-for-byte —
the package does not interpret the template, it opens and closes one seam.

## Then: v3.6 — Resumable exports, Azure Blob source & writer styling

### Writer styling API (merge / alignment / sheet name)

Template mode covers *static* layout, but not a merge whose row span is only
known while streaming. First-class `mergeColumnsByGroup()` (hooking the group
tracking `syncAtGroupBoundaries()` already does), `halign`/`valign`/`wrap` via
`setColumnStyle()`, and `setSheetName()` for the first sheet.

### Resumable S3 exports

Snapshot writer state at sync points — the `SCRC` running CRCs shipped
in v3.2 exist precisely for this. `Writer::resume($snapshot)` after a
crash re-enters at row N+1 instead of restarting a 25-minute queue job.
The deflate mechanics (userland `crc32_combine`, full-flush segment
concatenation) are already PoC-proven.

### Azure Blob storage source ([Issue #1](https://github.com/turgutahmet/kolay-xlsx-stream/issues/1))

A `Source` adapter for Azure Blob Storage, parallel to the S3 range source,
so the same zero-download random-access reads (`rowAt`, `profile()`, string
`findRow`, …) work against Azure. Kept an isolated adapter (not bundled into
the core), mindful of streaming memory — a ranged read, never a whole-blob
buffer.

### Reader I/O planning

Shipped in v3.3 ✅: **range coalescing** (gap-bridging with a cost model
via `Contracts\ProvidesCostHints` — keep one ranged read alive across a
short gap when it beats a round-trip), **`explain()`/`estimatedRows()`**
(zero-I/O plans + hard upper bound + t-digest point estimate),
**`rowsWhereAll()`** (multi-predicate zone-map intersection). Still ahead:

- **Hedged range requests** — re-issue a slow critical-path range on a
  second connection, first response wins; insurance against S3 tail
  latency on point lookups.
- **Prefetch pipeline** — keep 3-4 blocks in flight during sequential
  S3 scans so inflate/parse overlaps fetch.
- **`onProgress` callback for the reader** — symmetric to the writer's,
  for ingestion dashboards. (The reader already has `onFullScan()` for
  pushdown observability as of v3.3.)

### Tooling

- **`xl/_kxs/index.bin` migration tool** — CLI that walks an S3 prefix
  (or local directory), re-encodes each XLSX with
  `withRandomAccessIndex()` and uploads with `If-Match` for race
  safety — retroactively makes existing data lakes queryable.
- **Reference TS/JS reader** — a small browser/Node KXSI reader over
  HTTP Range requests (fetch + `DecompressionStream`), driven by the
  SPEC.md conformance vectors. Proves the format is a spec, not a PHP
  quirk.

### Polish (carried over)

- Skip `<col customWidth="1"/>` when the resolved width matches
  Excel's default (8.43).
- Lazy S3 client / multipart init — defer `createMultipartUpload` to
  the first `write()` so the sink is cheap to instantiate in DI
  contexts.

### Queryable-XLSX follow-ups (PoC-verified, sequenced after v3.6)

- **Appendable XLSX** — end the last sheet at a full-flush boundary,
  reopen and continue with a fresh deflate context; on S3,
  `UploadPartCopy` makes append O(appended data). Wants ZIP64 first.
- **Distributed export + server-side stitch** — N queue workers each
  write a headless full-flush-terminated deflate segment; a finalizer
  assembles one valid born-indexed .xlsx via `UploadPartCopy` without
  downloading the segments. Parallel per-block deflate is provably
  **byte-identical** to serial output (PoC), so the whole path can be
  CI-verified by hash comparison.
- **Verified partial reads (`MRKL`)** — a Merkle tree over per-block
  hashes (tag reserved in SPEC.md): prove a ranged read belongs to a
  signed file without fetching the rest of it.
- ~~**Compact cell refs** (`c/@r` omitted, spec-legal)~~ — **shipped in
  v3.3 as opt-in `compact()`** (~52–62 % smaller compressed sheets;
  verified in Excel/LibreOffice/Numbers and read by PhpSpreadsheet 5.8 /
  OpenSpout 4.28). Stays opt-in while it accrues field mileage.
- **Retrofit index for foreign XLSX** — zran-style inflate-state
  snapshots make files *other* writers produced random-access after
  one pass. Pure PHP can't resume mid-stream (no `inflatePrime`);
  proven feasible via FFI → candidate for an optional
  `kolay/xlsx-stream-retrofit` package.

## Future: v4.0 — Considerations

Things that probably need a major version bump because they change
public behaviour or require a newer floor. **None of these are
committed yet** — they go in once the cost/benefit case is clear.

- **ZIP64 writer support** (files > 4 GB) — the real unlock for
  appendable workbooks and distributed stitching at scale; the reader
  side already detects ZIP64 sentinels and refuses loudly.
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
- **Schema-specialized row builder (codegen)** — generate one closure
  per column-type signature so each row is a single call with zero
  per-cell branching; must be proven byte-identical like the v3.1/v3.2
  hot-path work before it ships.

## How to suggest features

Open an issue with the use case (one paragraph is enough). Concrete
"my export does X, currently I have to Y" reports beat abstract
wishlists by a wide margin. Performance suggestions backed by a
benchmark land fastest.

PRs welcome — see [CONTRIBUTING](README.md#contributing). Tests +
benchmark gate are required for anything that touches the hot path.

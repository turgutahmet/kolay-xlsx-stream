# Roadmap

This document tracks where the package is heading. Items are best-effort —
priorities shift based on real-world feedback and contributions. Anything
listed below `Backlog` is "considered, not committed".

## Shipped

| Version | Date | Highlights |
|---|---|---|
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

## Next: v3.3 — Durable exports & smarter S3 reads

Additive, no breaking changes planned. The through-line: exports that
survive crashes, and reads that plan their own I/O.

### Resumable S3 exports (the headline)

Snapshot writer state at sync points — the `SCRC` running CRCs shipped
in v3.2 exist precisely for this. `Writer::resume($snapshot)` after a
crash re-enters at row N+1 instead of restarting a 25-minute queue job.
The deflate mechanics (userland `crc32_combine`, full-flush segment
concatenation) are already PoC-proven.

### Reader I/O planning

- **Range coalescing with a cost model** — when a pruned query wants
  blocks {17, 19, 20}, decide between three requests or one padded
  request from measured RTT/bandwidth (S3 `GetObject` accepts only a
  single range per request).
- **`explain()` / `estimatedRows()`** — return the query plan without
  executing: strategy, candidate blocks, estimated requests/bytes, and
  a zero-I/O result-size estimate (zone-map counts give a hard upper
  bound; the t-digest CDF gives a point estimate). Doubles as a CI
  oracle (plan vs. actual request count).
- **`rowsWhereAll()`** — multi-predicate queries whose per-column
  zone-map survivor sets are intersected before any block is read
  (`WHERE date BETWEEN … AND amount > …` scans the intersection, never
  more blocks than the most selective predicate alone).
- **Hedged range requests** — re-issue a slow critical-path range on a
  second connection, first response wins; insurance against S3 tail
  latency on point lookups.
- **Prefetch pipeline** — keep 3-4 blocks in flight during sequential
  S3 scans so inflate/parse overlaps fetch.
- **`onProgress` callback for the reader** — symmetric to the writer's,
  for ingestion dashboards.

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

### Queryable-XLSX follow-ups (PoC-verified, sequenced after v3.3)

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
- **Compact cell refs** (`c/@r` omitted, spec-legal) — measured 2.1×
  writer throughput and −44 % file size, but at least one popular
  reader silently returns zero rows on such files; stays opt-in and
  gated on the external-reader compat matrix.
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

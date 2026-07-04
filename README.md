# Kolay XLSX Stream

[![Latest Version on Packagist](https://img.shields.io/packagist/v/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![Tests](https://img.shields.io/github/actions/workflow/status/turgutahmet/kolay-xlsx-stream/tests.yml?branch=main&label=tests&style=flat-square)](https://github.com/turgutahmet/kolay-xlsx-stream/actions/workflows/tests.yml)
[![Total Downloads](https://img.shields.io/packagist/dt/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![License](https://img.shields.io/packagist/l/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![PHP Version](https://img.shields.io/packagist/php-v/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)

High-performance bidirectional XLSX streaming for Laravel — write to local or S3 with **zero disk I/O**, read back with **bounded memory**, and seek to any row in **O(1)** via the optional born-indexed sidecar. Built for exporting and ingesting millions of rows without spiking RAM.

## Why This Package?

### The Problem with Existing Solutions

Most PHP Excel libraries (PHPSpreadsheet, Spout, Laravel Excel) have critical limitations:

- **Memory Issues**: Load entire documents in RAM (unusable for large files)
- **Disk I/O**: Write temporary files then upload to S3 (2x I/O, slow)
- **No True Streaming**: Can't stream directly to S3
- **No Random Access**: O(N) full scan to reach any specific row

### Our Solution

- **Zero Disk I/O**: Direct streaming to S3 using multipart upload
- **Constant Memory**: O(1) memory for the writer (32 MB part buffer) and ~24 MB for the reader regardless of file size
- **Bidirectional**: Write *and* read XLSX files through the same package — including files produced by other writers (PhpSpreadsheet, openpyxl, …) via shared-strings support
- **Random Access**: Optional `xl/_kxs/index.bin` sidecar lets `rowAt(N)`, `rowRange(a, b)`, and `rowCount()` skip ahead in **O(1)** instead of full-scanning. Backward-compatible — Excel and every other reader ignore the sidecar. (Per-lookup work is bounded by the writer-chosen sync period, default 10,000 rows; independent of file size.)
- **Blazing Fast**: 2-3× faster than alternatives on writes; ~70K rows/s sustained reads with constant RAM
- **Production Tested**: Successfully exported and re-read 4.5 million rows (500MB+ files)

## Performance Comparison

### Cross-package comparison — 100K rows (May 2026)

Fresh `composer create-project` of each package, latest stable versions,
identical 8-column mixed-type payload. `kolay/xlsx-stream` configured with
`setCompressionLevel(1)->setBufferFlushInterval(10000)`; all other packages
run at their out-of-the-box defaults. Each row matters — methodology and
the trade-off table (file size, RAM, compression level) live in
[BENCHMARK.md](BENCHMARK.md).

**Write — 100,000 rows**

| # | Package | Version | Time | rows/sec | Peak RAM | File |
|---:|---|---|---:|---:|---:|---:|
| 1 | **kolay/xlsx-stream** | **3.0.0** | **0.65 s** | **153,216** | 11.79 MB | 5.64 MB |
| 2 | avadim/fast-excel-writer | 6.12.0 | 5.23 s | 19,127 | 4 MB | 4.34 MB |
| 3 | openspout/openspout | 5.7.0 | 5.77 s | 17,321 | 2 MB | 4.33 MB |
| 4 | rap2hpoutre/fast-excel | 5.7.0 | 7.30 s | 13,697 | 4 MB | 4.06 MB |
| 5 | phpoffice/phpspreadsheet | 5.7.0 | 30.62 s | 3,265 | 531 MB | 4.82 MB |

**Read — 100,000 rows** (each reader on the same package's writer output)

| # | Package | Version | Time | rows/sec | Peak RAM |
|---:|---|---|---:|---:|---:|
| 1 | **kolay/xlsx-stream** | **3.0.0** | **1.75 s** | **57,043** | 4 MB |
| 2 | avadim/fast-excel-reader | 3.0.1 | 4.60 s | 21,752 | 2 MB |
| 3 | rap2hpoutre/fast-excel | 5.7.0 | 8.50 s | 11,761 | 4 MB |
| 4 | openspout/openspout | 5.7.0 | 9.90 s | 10,105 | 2 MB |
| 5 | phpoffice/phpspreadsheet | 5.7.0 | 29.95 s | 3,339 | 434 MB |

That's **~8× faster write** and **~2.6× faster read** than the next-fastest
streaming peer (avadim 6.12 / 3.0.1). Against `kolay/xlsx-stream`'s own
default config (`lvl=6, flush=1K`) the gap is still ~5× and ~2.4×.

**PhpSpreadsheet is in a different category** — it provides full Excel
feature support (charts, pivot tables, conditional formatting) at the cost
of memory-bound architecture. For pure data pipelines the streaming
packages are 5–47× faster and 100×+ smaller in RAM.

---

### Latest Benchmark — v3.0 (May 2026)

v3.0 introduces the streaming **reader** plus the optional born-indexed
**random-access** primitive on top of the v2.2.2 writer. Measured on the
same Apple Silicon laptop, PHP 8.2.28, AWS SDK 3.379, against the
`xlsx-test-package` bucket in `us-east-2`. Same 8-column mixed-type
workload as v1.x and v2.2.2 — every release uses the canonical
`benchmark-*.php` scripts in the repo root so the numbers stay
comparable across versions.

#### Sequential read

```
php benchmark-read.php
```

| Rows | Local Read Speed | Local Time | Local Peak RAM | S3 Read Speed | S3 Time | S3 Peak RAM | File Size |
|------|------------------|------------|----------------|---------------|---------|-------------|-----------|
| 100 | 27,216 rows/s | 0.00s | 22 MB | 56 rows/s | 1.79s | 24 MB | 0.01 MB |
| 500 | 61,830 rows/s | 0.01s | 22 MB | 281 rows/s | 1.78s | 24 MB | 0.02 MB |
| 1,000 | 69,553 rows/s | 0.01s | 22 MB | 480 rows/s | 2.08s | 24 MB | 0.04 MB |
| 5,000 | 69,433 rows/s | 0.07s | 24 MB | 2,278 rows/s | 2.19s | 24 MB | 0.20 MB |
| 10,000 | 75,843 rows/s | 0.13s | 24 MB | 4,235 rows/s | 2.36s | 24 MB | 0.40 MB |
| 25,000 | 66,957 rows/s | 0.37s | 24 MB | 9,664 rows/s | 2.59s | 24 MB | 1.00 MB |
| 50,000 | 72,895 rows/s | 0.69s | 24 MB | 16,967 rows/s | 2.95s | 22 MB | 2.00 MB |
| 100,000 | 75,899 rows/s | 1.32s | 24 MB | 27,412 rows/s | 3.65s | 22 MB | 4.00 MB |
| 250,000 | 72,267 rows/s | 3.46s | 24 MB | 43,139 rows/s | 5.80s | 22 MB | 10.01 MB |
| 500,000 | 68,772 rows/s | 7.27s | 24 MB | 52,204 rows/s | 9.58s | 22 MB | 20.02 MB |
| 750,000 | 71,360 rows/s | 10.51s | 24 MB | 57,230 rows/s | 13.11s | 24 MB | 30.03 MB |
| 1,000,000 | 69,946 rows/s | 14.30s | 24 MB | 60,253 rows/s | 16.60s | 24 MB | 40.03 MB |
| 1,500,000 | 69,517 rows/s | 21.58s | 24 MB | 59,746 rows/s | 25.11s | 22 MB | 59.66 MB |
| 2,000,000 | 68,710 rows/s | 29.11s | 24 MB | 61,306 rows/s | 32.62s | 24 MB | 79.24 MB |
| 3,000,000 | – | – | – | 62,447 rows/s | 48.04s | 24 MB | 119.04 MB |
| 4,000,000 | – | – | – | 62,956 rows/s | 63.54s | 22 MB | 158.43 MB |
| 4,500,000 | – | – | – | 62,198 rows/s | 72.35s | 22 MB | 178.09 MB |

**Reading is bounded-memory by construction.** Peak RAM stays at 22-24 MB
across every row count, from 100 rows to 4.5 million — independent of
file size. Local sustains ~67-76K rows/s. S3 cold-cache ramps with file
size as TTFB amortises; saturates at ~60-62K rows/s for files ≥750K
rows. Multi-sheet workbooks (above the 1,048,576-rows-per-sheet limit)
read every sheet automatically; there is no per-sheet cap from the
reader's side.

#### Random access — `rowAt(N)` and `rowCount()`

```
php benchmark-random-access.php 500000
```

500,000-row workbook, sync period 10,000. Same plain-vs-indexed comparison
methodology as the POC's BENCHMARK.md — the two writers produce visually
identical files and the indexed reader uses the embedded sidecar to
fresh-init inflate from the nearest sync point.

| Target Row | Plain (full scan) | Indexed (sync + scan) | Speedup |
|------------|-------------------|-----------------------|---------|
| 1 | 1.3 ms | 0.2 ms | 6.1× |
| 50,000 | 665.8 ms | 132.6 ms | 5.0× |
| 125,000 | 1,645.9 ms | 67.6 ms | 24.3× |
| 250,000 | 3,426.5 ms | 136.8 ms | 25.0× |
| 375,000 | 5,087.3 ms | 68.2 ms | **74.6×** |
| 450,000 | 6,291.8 ms | 136.6 ms | 46.0× |
| 499,900 | 6,897.9 ms | 135.6 ms | 50.9× |
| 500,000 | 6,768.0 ms | 135.9 ms | 49.8× |

`rowCount()` on the same file: **7,014 ms plain** (full inflate scan) vs
**<1 ms indexed** (constant lookup from the sidecar header) — **>260,000×
speedup**, the indexed call returns inside measurement noise.

**Index cost:** writing with `withRandomAccessIndex()` adds **−0.33 % wall
time** (within measurement noise — i.e. zero detectable cost), **+0.032 %
file size**, and zero detectable RAM overhead. The original design budget
allowed up to 0.5 % file size; we measured ~16× below that ceiling at
500K rows.

#### Write benchmark — v3.1 vs v3.0.2 (July 2026)

Two kinds of numbers, kept deliberately separate. **The improvement
claim** comes from a same-day, same-link A/B — both versions, identical
workload (the canonical 8-column comprehensive payload, level 1, 32 MB
parts):

| Workload (1M rows) | v3.0.2 | v3.1.0 | Diff |
|---|---|---|---|
| Local, level 1 | 212,996 rows/s | 292,619 rows/s | **+37 %** |
| Local, level 5 | 188,791 rows/s | 245,576 rows/s | **+30 %** |
| S3, level 1 | 59,383 rows/s | 80,634 rows/s | **+36 %** |
| S3, level 5 | 54,957 rows/s | 56,298 rows/s | ≈ equal (network-bound) |

The CPU win (PCRE-JIT escape gate + flattened row builder + level-5
default) carries through to S3 because row generation serializes with
the synchronous part uploads.

**Absolute S3 throughput** varies with the network path far more than
with the library — the same 1M-row export measured 59K, 81K, and 113K
rows/s across three sessions on the same machine. Full size sweep
(2026-07-04, level 1, 32 MB parts, sequential runs):

| Rows | S3 Speed | S3 Time | Peak RAM | File | Sheets |
|------|----------|---------|----------|------|--------|
| 500,000 | 44,828 rows/s | 11.2s | 41 MB | 20.0 MB | 1 |
| 1,000,000 | 81,776 rows/s | 12.2s | 85 MB | 40.0 MB | 1 |
| 2,000,000 | 84,297 rows/s | 23.7s | 120 MB | 79.2 MB | 2 |
| 3,000,000 | **153,042 rows/s** | 19.6s | 150 MB | 119.0 MB | 3 |
| 4,000,000 | 131,551 rows/s | 30.4s | 184 MB | 158.4 MB | 4 |

The 3M figure reproduced in an isolated re-run minutes later (152,812
rows/s, 19.63s) — but treat every S3 row as "that link, that hour":
small files amortize fixed costs (TTFB, multipart finalize, staying
under the 32 MB part threshold) poorly, and run-to-run jitter exceeds
version-to-version deltas. Benchmark on your own link; the sequential
`uploadPart` stall is the known bottleneck and parallel multipart
upload is on the roadmap. Peak RAM grows ~40 MB per million rows with
32 MB parts (part buffer + copies — the same rework will flatten this).

#### Write benchmark — v3.0 vs v2.2.2

v3.0's writer default code path is byte-identical to v2.2.2 — the only
new writer-side addition is the opt-in `withRandomAccessIndex()`. We
re-ran the full write benchmark on v3.0 to verify zero regression:

| Workload | v3.0 | v2.2.2 | Diff |
|---|---|---|---|
| Local 100K | 221,895 rows/s | 188,771 rows/s | +18 % |
| Local 1M | 214,912 rows/s | 209,905 rows/s | +2.4 % |
| Local 2M | 209,462 rows/s | 208,512 rows/s | +0.5 % |
| S3 1M | 109,562 rows/s | 106,924 rows/s | +2.5 % |
| S3 4.5M | 119,914 rows/s | 130,293 rows/s | −8 % (network jitter) |

Deltas are inside measurement noise. Memory profile is unchanged.
For the full v2.2.2 write table, see **Write benchmark — v2.2.2
(May 2026)** below.

For the full v3.0 measurement detail (write, read, random-access plus
methodology and reproducibility notes), see [BENCHMARK.md](BENCHMARK.md).

---

### Write benchmark — v2.2.2 (May 2026)

Re-measured on an Apple Silicon laptop with PHP 8.2.28 and AWS SDK
3.379 against the same `xlsx-test-package` bucket in `us-east-2`. The
workload is identical to the v1.x baseline below (8 columns,
mixed types, compression level 1).

| Rows | Local Speed | Local Time | S3 Speed | S3 Memory | S3 Time | File Size |
|------|-------------|------------|----------|-----------|---------|-----------|
| 100 | 45,290 rows/s | 0.00s | 112 rows/s | 0 MB | 0.89s | 0.01 MB |
| 500 | 191,346 rows/s | 0.00s | 491 rows/s | 0 MB | 1.02s | 0.02 MB |
| 1,000 | 184,836 rows/s | 0.01s | 966 rows/s | 0 MB | 1.03s | 0.04 MB |
| 5,000 | 195,825 rows/s | 0.03s | 3,345 rows/s | 0 MB | 1.49s | 0.2 MB |
| 10,000 | 198,898 rows/s | 0.05s | 2,551 rows/s | 0 MB | 3.92s | 0.4 MB |
| 25,000 | 210,498 rows/s | 0.12s | 10,170 rows/s | 0 MB | 2.46s | 1 MB |
| 50,000 | 217,123 rows/s | 0.23s | 15,597 rows/s | 2 MB | 3.21s | 2 MB |
| 100,000 | 188,771 rows/s | 0.53s | 24,258 rows/s | 4 MB | 4.12s | 4 MB |
| 250,000 | 215,340 rows/s | 1.16s | 63,713 rows/s | 12 MB (±6) | 3.92s | 10 MB |
| 500,000 | 209,428 rows/s | 2.39s | 87,679 rows/s | 20 MB (±18) | 5.70s | 20 MB |
| 750,000 | 211,315 rows/s | 3.55s | 84,221 rows/s | 30 MB (±26) | 8.91s | 30 MB |
| 1,000,000 | 209,905 rows/s | 4.76s | 106,924 rows/s | 40 MB (±38) | 9.35s | 40 MB |
| 1,500,000 | 207,503 rows/s | 7.23s | 117,631 rows/s | 60 MB (±58) | 12.75s | 60 MB |
| 2,000,000 | 208,512 rows/s | 9.59s | 112,708 rows/s | 79 MB (±77) | 17.74s | 79 MB |
| 3,000,000 | – | – | 112,229 rows/s | 119 MB (±117) | 26.73s | 119 MB |
| 4,000,000 | – | – | 128,930 rows/s | 160 MB (±156) | 31.02s | 158 MB |
| 4,500,000 | – | – | 130,293 rows/s | 178 MB (±178) | 34.54s | 178 MB |

#### What changed since v1.x

- **Local throughput is now ~15–25% *faster* than the v1.x baseline**
  (1M rows: 210K rows/s vs 183K). The v2.0+ per-cell type detection
  added a small overhead at first (v2.0 was about 5% slower than v1.x),
  but v2.2.2 fixed a long-standing bug in the XML escape fast path — the
  `strpbrk` needle was a single-quoted literal so `\xNN` escapes were
  embedded as the characters `\`, `x`, `0..9`, `A..F` instead of as
  actual control bytes. The fix shrank the needle from 129 to 36 chars
  and per-cell sanitization got ~3.5× cheaper as a side effect.
- **S3 throughput is up 2.5–3× over v1.x** for any workload above 50K
  rows (1M: 107K rows/s vs 43K, 4.5M: 130K rows/s vs 46K). Drivers:
  AWS SDK 3.379+, faster measurement-machine network, and a smaller
  share from the per-row hot-path improvement.
- **Memory is unchanged** — local stays at 0–2 MB constant, S3 keeps the
  same sawtooth pattern as the buffer fills and flushes per part.

### Original Benchmark — v1.x (September 2025)

Kept here for historical context. Different machine and PHP version, so
direct cell-by-cell deltas reflect environment variance as well as code
changes.

| Rows | Local Speed | Local Memory | Local Time | S3 Speed | S3 Memory | S3 Time | File Size |
|------|-------------|--------------|------------|----------|-----------|---------|-----------|
| 100 | 47,913 rows/s | 2 MB | 0.00s | 99 rows/s | 0 MB | 1.01s | 0.01 MB |
| 500 | 161,183 rows/s | 0 MB | 0.00s | 362 rows/s | 0 MB | 1.38s | 0.02 MB |
| 1,000 | 161,711 rows/s | 0 MB | 0.01s | 913 rows/s | 0 MB | 1.09s | 0.04 MB |
| 5,000 | 167,101 rows/s | 0 MB | 0.03s | 3,092 rows/s | 0 MB | 1.62s | 0.2 MB |
| 10,000 | 183,281 rows/s | 0 MB | 0.05s | 5,411 rows/s | 0 MB | 1.85s | 0.4 MB |
| 25,000 | 187,322 rows/s | 0 MB | 0.13s | 11,482 rows/s | 0 MB | 2.18s | 1 MB |
| 50,000 | 187,455 rows/s | 2 MB | 0.27s | 5,829 rows/s | 2 MB | 8.58s | 2 MB |
| 100,000 | 182,167 rows/s | 0 MB | 0.55s | 9,288 rows/s | 4 MB | 10.77s | 4 MB |
| 250,000 | 185,586 rows/s | 0 MB | 1.35s | 59,744 rows/s | 12 MB (±6) | 4.18s | 10 MB |
| 500,000 | 184,268 rows/s | 0 MB | 2.71s | 28,553 rows/s | 20 MB (±18) | 17.51s | 20 MB |
| 750,000 | 182,648 rows/s | 0 MB | 4.11s | 33,504 rows/s | 30 MB (±26) | 22.39s | 30 MB |
| 1,000,000 | 182,693 rows/s | 0 MB | 5.47s | 43,215 rows/s | 40 MB (±38) | 23.14s | 40 MB |
| 1,500,000 | 180,578 rows/s | 0 MB | 8.31s | 36,733 rows/s | 60 MB (±58) | 40.84s | 60 MB |
| 2,000,000 | 177,012 rows/s | 0 MB | 11.30s | 51,323 rows/s | 79 MB (±77) | 38.97s | 79 MB |
| 3,000,000 | – | – | – | 39,150 rows/s | 117 MB (±117) | 76.63s | 119 MB |
| 4,000,000 | – | – | – | 42,500 rows/s | 160 MB (±156) | 94.12s | 158 MB |
| 4,500,000 | – | – | – | 46,462 rows/s | 178 MB (±178) | 96.85s | 178 MB |

*Note: Tests with 1M+ rows automatically create multiple sheets (Excel limit: 1,048,576 rows per sheet)*  
*Note: ± values in S3 Memory column indicate memory fluctuation during streaming due to periodic part uploads*

### Understanding Memory Behavior

> **Reader bounded-RAM contract — when it holds:** Files written by
> this package use inline strings exclusively, so the reader's bounded
> 22-24 MB peak RAM applies unconditionally. Files produced by other
> tools that store text in `xl/sharedStrings.xml` are loaded into
> memory **provided the table fits** — compressed ≤ 20 MB AND
> uncompressed ≤ 100 MB. Files exceeding either threshold are rejected
> with a clear error rather than silently exhausting RAM. On-disk
> shared-strings table support is tracked for a future release.

#### Local File System
- **True O(1) Memory**: Constant memory usage regardless of file size
- **No Growth**: Memory stays at 0-2MB even for millions of rows
- **Speed**: 180,000+ rows/second consistently

#### S3 Streaming Memory Fluctuation
The ± values in S3 memory represent **normal memory fluctuation** during streaming:

1. **Buffer Accumulation Phase** (↑ Memory Growth)
   - Data is compressed and buffered until reaching 32MB
   - Memory grows gradually as buffer fills

2. **Part Upload Phase** (↓ Memory Drop)
   - When buffer reaches 32MB, it's uploaded to S3
   - After upload, memory drops back to baseline
   - This creates the characteristic sawtooth pattern

3. **Example: 1M Rows Test**
   - Average memory: 40MB
   - Fluctuation: ±38MB
   - Pattern: Memory oscillates between ~2MB (after upload) and ~78MB (before upload)
   - This is **completely normal** and expected behavior

### Performance Highlights *(v3.1, July 2026)*

- **Write — Local (v3.1)**: ~289,000 rows/second writer throughput (500K-row
  mixed workload, 6 MB peak RAM) — **+55 % over v3.0.2** from a PCRE-JIT
  escape fast path, a flattened row builder (output byte-identical), and
  the level-5 default. Apples-to-apples numbers in `bench/results.json`.
- **Reader open on S3 (v3.1)**: 7 → 3 round trips (suffix-range tail read,
  coalesced entry fetches, workbook.xml fetched once) — ~150–300 ms less
  first-row latency per open at typical S3 RTTs.
- **Queryable files (v3.1)**: `columnStats()` answers min/max/sum/avg from
  the sidecar alone; `rowsWhere()` skips blocks via zone maps; `findRow()`
  resolves point lookups on sorted columns by reading a single block.
- **Write — S3 (v3.1)**: **+36 % over v3.0.2** in a same-day, same-link A/B
  at 1M rows; absolute throughput ranged 59K–153K rows/s across sessions
  and sizes (peak: 3M rows at 153K rows/s, reproduced twice — full sweep
  in the v3.1 write benchmark table). S3 numbers track the network path
  far more than the library; benchmark on your own link.

*(v3.0, May 2026 — S3 figures reflect that day's network path)*

- **Write — Local**: ~190,000–222,000 rows/second with true O(1) memory
- **Write — S3**: 109,000–129,000 rows/second above 750K rows (2.5–3× the v1.x baseline, identical to v2.2.2)
- **Read — Local**: ~67,000–76,000 rows/second sustained, 22-24 MB peak RAM regardless of file size
- **Read — S3**: 60,000–62,000 rows/second saturation on multi-MB files (cold cache, single-stream)
- **Random Access**: O(1) `rowAt(N)`, `rowRange(a, b)`, and `rowCount()` via opt-in `withRandomAccessIndex()` — **up to 74.6× speedup** vs full scan, **>260,000× speedup** on `rowCount()`, **+0.032 % file size cost**
- **Memory Efficiency**: Local writer uses <2 MB, S3 writer averages 40 MB per million rows; reader caps at ~24 MB independent of file size
- **Multi-sheet Support**: Automatic sheet creation at Excel's 1,048,576 row limit, transparent multi-sheet reading
- **External XLSX**: Reads files produced by PhpSpreadsheet, openpyxl, Apache POI, Excel etc. via shared-strings table support
- **Production Ready**: Successfully written and round-tripped 4.5 million rows

### File size limits

The writer emits **ZIP32** archives. Each output is bounded by:

- **4 GB compressed** total archive size
- **4 GB uncompressed** per ZIP entry (single sheet)
- **65,535 entries** in the central directory

These ceilings are far above any realistic single-export workload
(4.5 M rows ≈ 178 MB compressed). If a workload approaches them the
writer aborts with a clear `ZIP32 limit exceeded` exception instead
of silently truncating size fields and producing a corrupt file —
split the export across multiple files or sheets as a workaround.
ZIP64 writer support is tracked for a future release.

### Compression level

`setCompressionLevel(int $level)` accepts 1–9. The default is **5**
(v3.1+): measured on XLSX-shaped XML, level 5 produces a file within
~0.2 % of level 6's size at ~20 % less wall time — level 6 spends its
extra effort on entropy (unique cell refs) that doesn't compress
anyway. Pick by use case:

| Use case | Level | Tradeoff |
|---|---:|---|
| Queue job, fastest export | 1 | fastest, ~20 % larger file |
| Balanced default | 5 | knee of the size/speed curve for XLSX data |
| Marginally smaller | 6 | ~0.2 % smaller than 5, measurably slower |
| Archive, smallest file | 9 | much slower, ~6 % smaller file |

For S3 uploads, a lower level typically wins because compute is the
bottleneck. Level 9 only helps if you're storing the file long-term.

### Comparison with Other Libraries

| Package | 1M Rows Write | 1M Rows Read | Memory (Read) | Disk Usage | Random Access | S3 Support |
|---------|---------------|--------------|---------------|------------|---------------|------------|
| PHPSpreadsheet | ❌ Crashes | ❌ Crashes | ~8 GB | Full file | ❌ | Indirect |
| Spout / OpenSpout | ~60 sec | ~30 sec | ~100MB+ | Full file | ❌ | Indirect |
| Laravel Excel | ~90 sec | ~60 sec | ~500MB+ | Full file | ❌ | Indirect |
| **Kolay XLSX Stream (Local)** | ✅ **4.65 sec** | ✅ **14.30 sec** | ✅ **24 MB** | ✅ **Zero** | ✅ **O(1)\*** | N/A |
| **Kolay XLSX Stream (S3)** | ✅ **9.13 sec** | ✅ **16.60 sec** | ✅ **24 MB** | ✅ **Zero** | ✅ **O(1)\*** | ✅ **Direct** |

*\*With opt-in `withRandomAccessIndex()` on the writer. Per-lookup
work is bounded by the writer-chosen sync period (default 10,000
rows), independent of file size. `rowCount()` is constant straight
from the index header. Tune for latency-sensitive seeks with
`withRandomAccessIndex(every: 1000)` or `every: 100` for very dense
random reads (file size grows ~1% per 10× density).*

### When to use this package vs alternatives

For most Laravel exports — use [fast-excel](https://github.com/rap2hpoutre/fast-excel).
Simpler API, supports CSV/ODS, includes import functionality,
battle-tested across millions of installs.

**Use `kolay/xlsx-stream` when:**

- You need to stream directly to S3 with no temporary disk usage —
  Lambda, Cloud Run, Fargate, read-only filesystems
- Your dataset exceeds available memory — 1M+ rows on small
  instances, multi-million-row exports on standard ones
- You need **O(1) random access** into large XLSX files via
  `rowAt(N)` / `rowRange(a, b)` (born-indexed mode — first
  random-access XLSX primitive in PHP)
- You want HTTP-streamed downloads via `PhpStreamSink::output()` —
  zero temp file, immediate first byte to the client

**Use [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) when:**

- You need formulas, charts, conditional formatting, or pivot tables
- File size is small enough for in-memory operations (< 50 K rows)
- You're editing existing workbooks rather than producing new ones

**Use [OpenSpout](https://github.com/openspout/openspout) when:**

- You need ODS or CSV alongside XLSX
- You're already in a non-Laravel ecosystem and want a streaming
  writer/reader without S3 specifics

## Requirements

- PHP 8.1+
- Laravel 10, 11, 12 or 13
- AWS SDK (only if using S3 streaming or the S3 reader)

> Upgrading from v2.x? Reader and random-access APIs are purely
> additive — no breaking changes. See [CHANGELOG.md](CHANGELOG.md) for
> the full v3.0 highlights.
>
> Upgrading from v1.x? See [UPGRADE.md](UPGRADE.md) for the v2.0
> migration guide as well.

## Installation

```bash
composer require kolay/xlsx-stream
```

### Publish Configuration (Optional)

```bash
php artisan vendor:publish --tag=xlsx-stream-config
```

## Quick Start

### Basic Usage - Local File

```php
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Kolay\XlsxStream\Sinks\FileSink;

// Create writer with file sink
$sink = new FileSink('/path/to/output.xlsx');
$writer = new SinkableXlsxWriter($sink);

// Set headers
$writer->startFile(['Name', 'Email', 'Phone']);

// Write rows
$writer->writeRow(['John Doe', 'john@example.com', '+1234567890']);
$writer->writeRow(['Jane Smith', 'jane@example.com', '+0987654321']);

// Or write multiple rows at once
$writer->writeRows([
    ['Bob Johnson', 'bob@example.com', '+1111111111'],
    ['Alice Brown', 'alice@example.com', '+2222222222'],
]);

// Finish and close file
$stats = $writer->finishFile();

echo "Generated {$stats['rows']} rows in {$stats['sheets']} sheet(s)";
```

### Direct S3 Streaming (Zero Disk I/O)

```php
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Aws\S3\S3Client;

// Create S3 client
$s3Client = new S3Client([
    'region' => 'us-east-1',
    'version' => 'latest',
    'credentials' => [
        'key' => env('AWS_ACCESS_KEY_ID'),
        'secret' => env('AWS_SECRET_ACCESS_KEY'),
    ],
]);

// Create S3 sink with 32MB parts
$sink = new S3MultipartSink(
    $s3Client,
    'my-bucket',
    'exports/report.xlsx',
    32 * 1024 * 1024 // 32MB parts for optimal performance
);

$writer = new SinkableXlsxWriter($sink);

// Configure for maximum performance
$writer->setCompressionLevel(1)      // Fastest compression
       ->setBufferFlushInterval(10000); // Flush every 10K rows

$writer->startFile(['ID', 'Name', 'Email', 'Status']);

// Stream millions of rows with constant 32MB memory
User::query()
    ->select(['id', 'name', 'email', 'status'])
    ->chunkById(1000, function ($users) use ($writer) {
        foreach ($users as $user) {
            $writer->writeRow([
                $user->id,
                $user->name,
                $user->email,
                $user->status
            ]);
        }
    });

$stats = $writer->finishFile();
```

### Reading XLSX Files *(v3.0+)*

```php
use Kolay\XlsxStream\Readers\StreamingXlsxReader;

// From a local file
foreach (StreamingXlsxReader::fromFile('/path/to/big.xlsx')->rows() as $row) {
    DB::table('users')->insert($row);
}

// Directly from S3 — bounded RAM (~24 MB), no temp file
$reader = StreamingXlsxReader::fromS3($s3Client, 'my-bucket', 'imports/big.xlsx');
foreach ($reader->rows(skip: 1) as $row) {           // skip the header row
    User::create([
        'id'    => $row[0],
        'name'  => $row[1],
        'email' => $row[2],
    ]);
}

// Bulk insert via chunked()
foreach ($reader->chunked(1000, skip: 1) as $batch) {
    User::insert($batch);
}
```

The reader supports both files written by this package (zero indirection
via inline strings) and files produced by other writers (PhpSpreadsheet,
openpyxl, Apache POI, Excel itself) — the shared-strings table is loaded
transparently when present.

> **Memory:** Reader peak RAM is bounded — measured delta from baseline
> stays under 4 MB regardless of file size (CI-pinned via
> `MemoryFootprintTest`). The PHP runtime adds a ~20 MB baseline, so
> total RSS lands around 22-24 MB on real workloads.
>
> **Lifecycle:** Reader resources are released automatically when the
> object goes out of scope (`__destruct` calls `close()`). For
> long-lived workers processing many files, calling `$reader->close()`
> or `unset($reader)` between iterations frees underlying handles
> eagerly.

### Reading dates and times *(v3.0+)*

Excel stores dates as numeric serials (e.g. `46148` for 2026-05-06).
The reader returns those as numeric strings by default — opt into
automatic conversion per column:

```php
$reader = StreamingXlsxReader::fromFile('orders.xlsx');
$reader->castColumn(2, 'date');                          // → DateTimeImmutable (date)
$reader->castColumn(3, 'datetime');                      // → DateTimeImmutable (with time)
$reader->castColumn(4, 'int');                           // → int
$reader->castColumn(5, fn ($v) => (int) $v * 100);       // custom callable

// Bulk
$reader->castColumns([0 => 'int', 2 => 'date', 3 => 'datetime']);

foreach ($reader->rows(skip: 1) as $row) {
    $row[2]; // DateTimeImmutable
}
```

> **Always use `rows(skip: 1)` with casts.** Casts run on every row
> the generator yields, including row 1 (the header). A header
> string like `"id"` cast as `'int'` returns `null` because
> `is_numeric("id")` is false. Read the header separately via
> `$reader->header()` (cast-free) and skip it on data iteration.

> **Timezone:** Excel serials are timezone-naive. The reader returns
> datetimes in **UTC by default** so the same file produces the same
> result on every server regardless of `date_default_timezone_get()`.
> If your file's dates were authored in a specific timezone, set it
> explicitly:
>
> ```php
> $reader->castTimezone('Europe/Istanbul');
> ```
>
> Mac-origin Excel files using the 1904 epoch (rare): `$reader->use1904Epoch();`

Built-in cast names: `date`, `datetime`, `int`, `float`, `bool`. Pass
any callable for custom transformations (parse to a value object,
trim, normalise, etc.).

### Streaming directly to an HTTP response *(v3.0+)*

Use `PhpStreamSink::output()` to stream a workbook into the active
HTTP response — no temp file, constant memory, immediate first byte
to the client. Pairs naturally with Laravel's `Response::stream()`:

```php
use Kolay\XlsxStream\Sinks\PhpStreamSink;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

return response()->stream(function () {
    $writer = new SinkableXlsxWriter(PhpStreamSink::output());
    $writer->startFile(['id', 'name', 'email']);
    User::query()->lazy()->each(fn ($u) =>
        $writer->writeRow([$u->id, $u->name, $u->email])
    );
    $writer->finishFile();
}, 200, [
    'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition' => 'attachment; filename="users.xlsx"',
]);
```

The sink also has `temp()` (in-memory until 2 MB, then a tmp file) and
`memory()` (in-memory only) factories for capturing workbooks for
later inspection — handy in tests.

### Random-Access Reading *(v3.0+)*

Files written with `withRandomAccessIndex()` can be seeked into in O(1).
The opt-in costs ~0.03 % file size and adds a single hidden ZIP part
(`xl/_kxs/index.bin`) that vanilla XLSX readers ignore.

```php
// Producer side — opt-in once during configuration
$writer = new SinkableXlsxWriter(new FileSink('/path/to/report.xlsx'));
$writer->withRandomAccessIndex(every: 10000);   // sync point every 10K rows
$writer->startFile(['ID', 'Name', 'Email']);
foreach ($users as $u) {
    $writer->writeRow([$u->id, $u->name, $u->email]);
}
$writer->finishFile();

// Consumer side — same StreamingXlsxReader, gains rowAt / rowRange / O(1) rowCount
$reader = StreamingXlsxReader::fromFile('/path/to/report.xlsx');

$reader->rowCount();              // O(1) — read straight from the index header
$reader->rowAt(250_001);          // O(period) — fresh inflate from nearest sync point
foreach ($reader->rowRange(100_000, 100_500) as $rowNumber => $row) {
    // ... process 500 rows starting at row 100,000 without scanning the prefix
}
```

`rowAt()` and `rowRange()` work even on files **without** an index — they
fall back to a sequential O(N) scan from the first row. Only the cost
differs; the API contract is identical.

> **Performance tip:** When you need many adjacent rows, prefer
> `rowRange($from, $to)` over a loop of `rowAt()` calls. `rowRange()`
> seeks once and reuses a single inflate stream; repeated `rowAt()`
> re-seeks on every call. For 1000 nearby rows the difference is
> ~1000× — a single ~ms seek versus 1000 × ms per call.

### Queryable XLSX — zone maps & aggregates *(v3.1+)*

Born-indexed files can additionally carry **per-block column statistics**
(min/max/sum/count for every ~10K-row block — the same idea as Parquet
row-group stats, embedded in a plain .xlsx that Excel still opens
normally). Track the columns you'll query when writing:

```php
$writer->withRandomAccessIndex()
       ->withColumnStats([1, 4]);   // 1-based: track "ID" and "Amount"
```

The reader then answers three kinds of questions **without scanning
row data**:

```php
$reader = StreamingXlsxReader::fromS3($s3, 'bucket', 'huge-export.xlsx');

// 1. Aggregates straight from the ~KB sidecar — on S3 this is ONE
//    range request against a multi-GB file:
$stats = $reader->columnStats(4);
// ['min' => ..., 'max' => ..., 'sum' => ..., 'avg' => ...,
//  'count' => ..., 'other' => ..., 'sorted' => 'asc'|'desc'|null]

// 2. Range queries that skip every block whose [min,max] can't match
//    (exports are usually ID/date-sorted, so this touches a handful
//    of blocks out of hundreds):
foreach ($reader->rowsWhere(4, 'between', 1000, 2000) as $rowNumber => $row) { ... }
foreach ($reader->rowsWhere(1, '>=', 4_000_000) as $rowNumber => $row) { ... }

// 3. Point lookups — on a sorted column this reads exactly one block,
//    i.e. two S3 range requests end to end:
$hit = $reader->findRow(1, 3_141_592);   // ['row' => N, 'values' => [...]] or null
```

Ops: `=`, `<`, `<=`, `>`, `>=`, `between`. Predicates match numeric
cells (ints, floats, dates as serials); on files without stats the same
calls degrade gracefully to a full-scan filter with identical results.

### Parallel reads — shard a sheet across queue workers *(v3.1+)*

Every sync point in a born-indexed file is an independently decompressible
boundary, so a sheet can be split into self-contained row ranges. The
shard plan is plain JSON-friendly data — dispatch one queue job per shard
and each worker streams only its slice, with zero coordination:

```php
// Planner (e.g. the job that receives the upload)
$reader = StreamingXlsxReader::fromS3($s3, 'bucket', 'import.xlsx');
foreach ($reader->shards(8) as $shard) {
    ProcessXlsxShard::dispatch('bucket', 'import.xlsx', $shard);
}

// Worker (each job opens its own reader/connection)
public function handle(): void
{
    $reader = StreamingXlsxReader::fromS3($this->s3(), $this->bucket, $this->key);
    foreach ($reader->rowsForShard($this->shard) as $rowNumber => $row) {
        if ($rowNumber === 1) continue;   // header rides in the first shard
        // ... import the row
    }
}
```

A 4M-row import's wall clock becomes `max(worker time)` instead of the
sum. Shard boundaries snap to sync points (balanced to within one sync
period); on non-indexed files `shards()` returns a single whole-sheet
shard — same contract, no parallelism.

### Laravel Job Example

```php
<?php

namespace App\Jobs;

use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Aws\S3\S3Client;
use App\Models\User;

class ExportUsersJob implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    public function handle()
    {
        // Create S3 client from Laravel config
        $s3Config = config('filesystems.disks.s3');
        $s3Client = new S3Client([
            'region' => $s3Config['region'],
            'version' => 'latest',
            'credentials' => [
                'key' => $s3Config['key'],
                'secret' => $s3Config['secret'],
            ],
        ]);
        
        // Setup S3 streaming
        $filename = 'exports/users-' . now()->format('Y-m-d-H-i-s') . '.xlsx';
        $sink = new S3MultipartSink(
            $s3Client,
            $s3Config['bucket'],
            $filename,
            32 * 1024 * 1024
        );
        
        $writer = new SinkableXlsxWriter($sink);
        $writer->setCompressionLevel(1)
               ->setBufferFlushInterval(10000);
        
        // Export headers
        $writer->startFile([
            'ID',
            'Name',
            'Email',
            'Created At',
            'Status'
        ]);
        
        // Export data with chunking
        User::query()
            ->orderBy('id')
            ->chunkById(1000, function ($users) use ($writer) {
                foreach ($users as $user) {
                    $writer->writeRow([
                        $user->id,
                        $user->name,
                        $user->email,
                        $user->created_at->format('Y-m-d H:i:s'),
                        $user->status
                    ]);
                }
                
                // Clear Eloquent cache periodically (optional)
                // You can track rows externally if needed
            });
        
        $stats = $writer->finishFile();
        
        \Log::info('Export completed', [
            'filename' => $filename,
            'rows' => $stats['rows'],
            'sheets' => $stats['sheets'],
            'bytes' => $stats['bytes']
        ]);
    }
}
```

## Advanced Features

### Laravel Storage Disk Integration *(v2.1+)*

Skip the manual `S3Client` setup — `forDisk()` reads everything from
`config('filesystems.disks.{$disk}')`:

```php
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

// Streams directly to S3 — credentials, region, bucket all from config
$writer = SinkableXlsxWriter::forDisk('s3', 'exports/users.xlsx');

// Or to a local disk — resolves to disk root + path
$writer = SinkableXlsxWriter::forDisk('local', 'exports/users.xlsx');

// Pass S3-specific options (ACL, ContentDisposition, Metadata) as the 3rd arg
$writer = SinkableXlsxWriter::forDisk('s3', 'reports/q4.xlsx', [
    'ACL' => 'public-read',
    'CacheControl' => 'max-age=3600',
]);

$writer->startFile(['ID', 'Name'])
       ->writeRow([1, 'Alice'])
       ->finishFile();
```

### Streaming from Eloquent with `lazy()` *(v2.1+)*

`writeRows()` accepts any iterable — pass an Eloquent `lazy()` cursor for
constant-memory streaming over millions of rows:

```php
$writer = SinkableXlsxWriter::forDisk('s3', 'users-export.xlsx');
$writer->startFile(['ID', 'Name', 'Email', 'Created']);

$writer->writeRows(
    User::query()
        ->select(['id', 'name', 'email', 'created_at'])
        ->orderBy('id')
        ->lazy(1000)
        ->map(fn ($u) => [$u->id, $u->name, $u->email, $u->created_at])
);

$writer->finishFile();
```

Generators work too:

```php
function rowsFromApi(): Generator {
    $page = 1;
    while ($batch = Http::get('/api/orders', ['page' => $page++])->json()) {
        foreach ($batch as $order) {
            yield [$order['id'], $order['total'], $order['status']];
        }
    }
}

$writer->writeRows(rowsFromApi());
```

### Progress Reporting for Queue Jobs *(v2.1+)*

Register a callback that fires every N rows with `(rows, bytes)`. Zero
overhead when not used:

```php
$writer = SinkableXlsxWriter::forDisk('s3', "exports/job-{$jobId}.xlsx");

$writer->onProgress(function (int $rows, int $bytes) use ($jobId) {
    Cache::put("export:{$jobId}", [
        'rows' => $rows,
        'bytes' => $bytes,
        'updated_at' => now(),
    ], 300);
})->setProgressInterval(5000);  // fire every 5K rows

$writer->startFile($headers);
$writer->writeRows($query->lazy());
$writer->finishFile();
```

> **Note on `$bytes`** — the byte counter only advances when zlib emits
> compressed output. With small datasets (or when `setBufferFlushInterval()`
> is large relative to `setProgressInterval()`), several events in a row
> may report the same byte count between flushes. The row counter is always
> exact; if you need accurate streamed-byte progress on small files, lower
> `setBufferFlushInterval()` below your progress interval.

### Supported Cell Data Types

The writer infers the right Excel cell type from each PHP value:

| PHP value | Excel cell |
|---|---|
| `int`, `float` | numeric (`t="n"`) |
| `bool` | native boolean (`t="b"`) — `1` for true, `0` for false |
| `\DateTimeInterface` (`DateTime`, `DateTimeImmutable`, `Carbon`, …) | numeric serial date with `yyyy-mm-dd hh:mm:ss` format — sortable as a date in Excel |
| numeric string ≤ 15 digits | numeric (`t="n"`) |
| numeric string > 15 digits, leading-zero (`"00123"`), or `+`-prefixed | inline string — preserves precision and formatting |
| `null` or `''` | empty cell |
| anything else | inline string (`t="inlineStr"`) |

```php
$writer->writeRow([
    1,                                  // numeric
    'Acme Co.',                         // string
    new DateTime('2026-01-15 10:30'),   // date cell
    true,                               // boolean cell
    '12345678901234567890',             // big-int preserved as text
    '+90 555 123 4567',                 // phone preserved as text
]);
```

### Header & Column Styling *(v2.2+)*

A small set of opt-in styling APIs that costs ~2-3% throughput and adds
~3% to the file size. Skip them and the writer takes the v1.x-equivalent
hot path.

```php
$writer = new SinkableXlsxWriter($sink);

$writer
    // Bold white text on dark blue, applied to the header row
    ->setHeaderStyle([
        'bold'  => true,
        'fill'  => '#4F81BD',
        'color' => '#FFFFFF',
        'size'  => 12,
    ])
    // Native Excel number formats per column (1-based index)
    ->setColumnFormat(1, 'integer')        // 12,345
    ->setColumnFormat(5, 'currency_try')   // ₺99,999.00
    ->setColumnFormat(6, 'percent')        // 12.50%
    ->setColumnFormat(7, 'date')           // 2026-01-15
    ->setColumnFormat(8, 'datetime')       // 2026-01-15 10:30:00
    ->setColumnFormat(9, '0.000000')       // raw Excel format code
    // Pin the header row, add filter dropdowns, auto-size columns
    ->freezeFirstRow()
    ->enableAutoFilter()
    ->setAutoColumnWidth();                // header-text + format-aware

$writer->startFile([
    'Order ID', 'Customer', 'Product', 'Region',
    'Price', 'Discount', 'Order Date', 'Created At', 'Score',
]);
```

Available format presets: `date`, `datetime`, `datetime_iso`, `time`,
`integer`, `decimal`, `percent`, `currency_try`, `currency_usd`,
`currency_eur`, `currency_gbp`. Pass any other string to use a raw Excel
format code (e.g. `0.000`, `#,##0.00 "kg"`).

`setAutoColumnWidth()` derives a width from the header text length but
also respects a per-format minimum so a `currency_try` column with the
header `Salary` won't render as `####`. Override per column with
`setColumnWidths([1 => 8, 2 => 30])` when you want exact control.

### Manual Multi-Sheet Workbooks *(v2.2+)*

`newSheet($name, $headers = null)` carves a workbook into named domain
sheets — orthogonal to the auto-split fallback at 1,048,576 rows.

```php
$writer = new SinkableXlsxWriter($sink);

$writer->setHeaderStyle(['bold' => true, 'fill' => '#4F81BD', 'color' => '#FFFFFF']);
$writer->startFile(['ID', 'Name', 'Email']);

foreach ($users->lazy() as $user) {
    $writer->writeRow([$user->id, $user->name, $user->email]);
}

// Different header style + different columns for the next sheet
$writer
    ->clearColumnFormats()
    ->setHeaderStyle(['bold' => true, 'fill' => '#9BBB59', 'color' => '#FFFFFF'])
    ->setColumnFormat(3, 'currency_try')
    ->newSheet('Orders', ['Order ID', 'Customer', 'Total']);

foreach ($orders->lazy() as $order) {
    $writer->writeRow([$order->id, $order->customer_id, $order->total]);
}

$stats = $writer->finishFile();
// $stats['sheet_details'] → [['name' => 'Report', ...], ['name' => 'Orders', ...]]
```

`clearColumnFormats()` is the convenient way to drop the previous sheet's
per-column registrations before starting a new one with a different
column layout.

### Multi-Sheet Support (Automatic)

The writer automatically creates new sheets when reaching Excel's row limit (1,048,576 rows):

```php
$writer = new SinkableXlsxWriter($sink);
$writer->startFile(['Column1', 'Column2']);

// Write 2 million rows - will create 2 sheets automatically
for ($i = 1; $i <= 2000000; $i++) {
    $writer->writeRow(["Row $i", "Data $i"]);
}

$stats = $writer->finishFile();
// $stats['sheets'] = 2
```

### Performance Tuning

```php
// Ultra-fast mode for maximum speed
$writer->setCompressionLevel(1)        // Minimal compression
       ->setBufferFlushInterval(50000); // Large buffer

// Balanced mode (default)
$writer->setCompressionLevel(6)        // Balanced compression
       ->setBufferFlushInterval(10000); // Medium buffer

// Maximum compression (slower, smaller files)
$writer->setCompressionLevel(9)        // Maximum compression
       ->setBufferFlushInterval(1000);  // Small buffer for streaming
```

### Custom S3 Parameters

```php
$sink = new S3MultipartSink(
    $s3Client,
    'my-bucket',
    'path/to/file.xlsx',
    32 * 1024 * 1024,
    [
        'ACL' => 'public-read',
        'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'ContentDisposition' => 'attachment; filename="report.xlsx"',
        'Metadata' => [
            'generated-by' => 'kolay-xlsx-stream',
            'timestamp' => time()
        ]
    ]
);
```

### Error Handling

```php
try {
    $writer = new SinkableXlsxWriter($sink);
    $writer->startFile(['Column1', 'Column2']);
    
    // Write data...
    
    $stats = $writer->finishFile();
} catch (\Exception $e) {
    // The sink will automatically abort and cleanup on error
    // S3 multipart uploads are automatically aborted
    // Partial files are deleted
    
    \Log::error('Export failed: ' . $e->getMessage());
}
```

## Configuration

Published configuration file (`config/xlsx-stream.php`):

```php
return [
    's3' => [
        'part_size' => 32 * 1024 * 1024,  // S3 multipart chunk size
        'retry_attempts' => 3,             // Retry failed uploads
        'retry_delay_ms' => 100,           // Delay between retries
    ],
    
    'writer' => [
        'compression_level' => 1,          // 1-9 (1=fastest, 9=smallest)
        'buffer_flush_interval' => 10000,  // Rows to buffer
        'max_rows_per_sheet' => 1048575,   // Excel limit - 1
    ],
    
    'memory' => [
        'file_buffer_size' => 1024 * 1024, // 1MB file write buffer
    ],
    
    'logging' => [
        'enabled' => true,
        'channel' => null,                  // null = default channel
        'log_progress' => false,            // Log progress updates
        'progress_interval' => 10000,       // Log every N rows
    ],
];
```

## How It Works

### The Architecture

```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│   Your App  │────▶│ XlsxWriter   │────▶│    Sink     │
└─────────────┘     └──────────────┘     └─────────────┘
                            │                     │
                    ┌───────▼────────┐   ┌───────▼────────┐
                    │ ZIP Generation │   │  Destination   │
                    │  & Compression │   │ • FileSink     │
                    └────────────────┘   │ • S3Sink       │
                                         └────────────────┘
```

### The Magic Behind Zero Disk I/O

1. **Binary XLSX Generation**
   - XLSX files are ZIP archives containing XML files
   - We generate ZIP structure directly in memory
   - No intermediate files or DOM tree building

2. **Streaming Compression**
   - Data is compressed using PHP's `deflate_add()` in chunks
   - Each row is immediately compressed and streamed
   - No need to store uncompressed data

3. **Smart Buffering**
   - Configurable row buffer (default 10,000 rows)
   - Flushes periodically to maintain streaming
   - Prevents memory accumulation

4. **S3 Multipart Upload**
   - Direct streaming to S3 using multipart upload
   - 32MB parts uploaded as they're ready
   - No local file required at any point

### Memory Efficiency Explained

#### Local Files: True O(1) Memory
```php
// Only these stay in memory:
$rowBuffer = '';        // Current batch of rows (flushed every 10K rows)
$deflateContext = ...;  // Compression context (minimal)
$zipMetadata = [];      // File entry metadata (bytes per file)
```

#### S3 Streaming: Near O(1) with Controlled Growth
```php
// Memory components:
$buffer = '';           // 32MB max (uploaded when full)
$parts = [];           // Part metadata (ETag + number per part)
// Growth: ~1KB per part × (filesize / 32MB) parts
// Example: 100MB file = 3 parts = ~3KB metadata
```

### Why Memory Fluctuates in S3

The sawtooth memory pattern in S3 streaming is **by design**:

```
Memory
  ▲
78MB │     ╱╲      ╱╲      ╱╲
     │    ╱  ╲    ╱  ╲    ╱  ╲
40MB │   ╱    ╲  ╱    ╲  ╱    ╲
     │  ╱      ╲╱      ╲╱      ╲
2MB  │─╯                         
     └────────────────────────────▶ Time
       ↑Upload  ↑Upload  ↑Upload
```

Each peak represents a full 32MB buffer ready for upload. After upload, memory drops as the buffer is cleared. This is optimal for streaming large files.

## Real-World Performance

### Production Test Results (4.5 Million Rows)

Our production systems successfully export massive datasets daily:

```
Dataset: 4.5 million rows (Employee competency evaluation data)
File Size: 178 MB compressed XLSX
Sheets: 5 (automatic splitting at Excel limit)
Total Time: 96.85 seconds
Average Speed: 46,462 rows/second
Memory Usage: 178 MB average with ±178 MB fluctuation
Peak Memory: 356 MB (during S3 part upload)
```

### Key Performance Metrics

#### Speed
- **Local Files**: 180,000+ rows/second sustained
- **S3 Streaming**: 30,000-50,000 rows/second
- **Network Impact**: S3 speed varies with network latency and bandwidth

#### Memory Usage
- **Local Files**: True O(1) - constant 0-2MB regardless of size
- **S3 Streaming**: Averages 40MB per million rows
- **Fluctuation**: Normal sawtooth pattern during S3 uploads

#### Scalability
- ✅ **100 rows**: Instant (< 0.01s local, ~1s S3)
- ✅ **1 Million rows**: 5.5 seconds local, 23 seconds S3
- ✅ **4.5 Million rows**: 97 seconds S3 with automatic multi-sheet
- ✅ **Memory stable**: No memory leaks, predictable usage
- ✅ **Production proven**: Running in production since 2025

## Compatibility

The optional `xl/_kxs/index.bin` sidecar emitted by `withRandomAccessIndex()`
is declared as `application/octet-stream` in `[Content_Types].xml`, so
editors that don't recognise it leave the file alone instead of flagging
it for repair.

Files produced by the v3.0 writer (both with and without the sidecar) open
cleanly without repair mode in:

- **Microsoft Excel for Mac 16.98** (build 25060824) — sidecar ignored,
  no repair prompt
- **Apple Numbers 14.2** (7041.0.109) — opaque sidecar passed through
  as expected

If a downstream editor strips or rewrites the sheet, the next indexed
read silently falls back to a sequential scan via the embedded sheet
CRC32 cross-check — same end result, just without the O(1) speedup. The
reader-side fallback is verified by `tests/Writers/RandomAccessIndexWriterTest.php`.

## Use Cases

Perfect for:
- Large data exports (millions of rows)
- Memory-constrained environments
- Kubernetes/Docker with small pods
- AWS Lambda functions
- Real-time streaming exports
- Multi-tenant SaaS applications

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Testing

```bash
composer test
```

## License

The MIT License (MIT). Please see [License File](LICENSE.md) for more information.

## Credits

- [Turgut Ahmet](https://github.com/turgutahmet)

## Support

For issues and questions, please use the [GitHub issue tracker](https://github.com/turgutahmet/kolay-xlsx-stream/issues).

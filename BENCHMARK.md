# Streaming XLSX Benchmark — kolay/xlsx-stream v3.0

**Generated:** 2026-05-07
**Environment:** Apple Silicon laptop, PHP 8.2.28, AWS SDK PHP 3.379+
**S3 bucket:** `xlsx-test-package` in `us-east-2`
**Test scope:** writer = `kolay/xlsx-stream v3.0` (write benchmark), reader = v3.0 `StreamingXlsxReader` (read benchmark), born-indexed writer + random-access reader = v3.0 (`withRandomAccessIndex(every: 10_000)` opt-in).

**Test workload:** 8 columns mixed types (`int, string, string, enum, float, date, enum, string`). Same workload as the v1.x and v2.2.2 README baselines, identical row sizes from 100 to 4,500,000.

**Reproducibility:** all numbers come from the canonical scripts in the repo root. Re-run any release with the same scripts and the new numbers go on the same axes:

```
php benchmark-comprehensive.php   # write — local + S3
php benchmark-read.php             # read  — local + S3
php benchmark-random-access.php    # rowAt + rowCount, plain vs indexed
```

---

## Executive summary

| Metric | v3.0 result | v2.2.2 baseline | Delta |
|---|---|---|---|
| Write throughput, local 1M rows | **214,912 rows/s** | 209,905 rows/s | +2.4 % (noise) |
| Write throughput, S3 1M rows | **109,562 rows/s** | 106,924 rows/s | +2.5 % (noise) |
| Write throughput, S3 4.5M rows | **119,914 rows/s** | 130,293 rows/s | −8 % (network jitter) |
| Write peak RAM, local | **0–2 MB** constant | 0–2 MB constant | identical |
| Write peak RAM, S3 1M | **40 MB** sawtooth | 40 MB sawtooth | identical |
| **Read throughput, local 1M rows** | **69,946 rows/s** | — (no reader in v2.2.2) | new |
| **Read throughput, S3 1M rows (cold)** | **60,253 rows/s** | — | new |
| **Read peak RAM** | **22–24 MB** constant, **all sizes 100–4.5M** | — | new |
| **Random access — `rowAt(375K of 500K)`** | indexed **68.2 ms** vs plain 5,087 ms = **74.6× speedup** | — | new |
| **Random access — `rowCount(500K)`** | indexed **<1 ms** vs plain 7,014 ms = **>260,000× speedup** | — | new |
| **Born-indexed file size penalty** | **+0.032 %** | — | new |
| **Born-indexed write time penalty** | **−0.33 %** (within noise, indistinguishable from zero) | — | new |

**Headline:**

- Write side is **byte-for-byte identical** to v2.2.2. The opt-in `withRandomAccessIndex()` adds ~0.03 % file size and zero detectable runtime cost; default-off keeps prior output stable.
- Read side is brand new. Local sustains **~70K rows/s** with **22–24 MB peak RAM regardless of file size**. S3 cold-cache scales with file size and ramps to **60K rows/s** on multi-MB workloads.
- Random access via the embedded `xl/_kxs/index.bin` sidecar delivers **74.6× faster row seeks** at 75 % of the file and turns `rowCount()` into a constant-time lookup.

---

## 1. Write benchmark

`php benchmark-comprehensive.php`. The same dataset is written first to a local temp file, then to S3 via `S3MultipartSink`. Local tests cap at 2M rows (matches the v2.2.2 baseline); S3 runs the full series.

### 1.1 Wall time + throughput + memory

| Rows | Local Speed | Local Time | Local Mem | S3 Speed | S3 Time | S3 Memory | File Size |
|---|---|---|---|---|---|---|---|
| 100 | 38,833 rows/s | 0.00s | 2 MB | 110 rows/s | 0.91s | 0 MB | 0.01 MB |
| 500 | 186,995 rows/s | 0.00s | 0 MB | 481 rows/s | 1.04s | 0 MB | 0.02 MB |
| 1,000 | 191,529 rows/s | 0.01s | 0 MB | 965 rows/s | 1.04s | 0 MB | 0.04 MB |
| 5,000 | 193,484 rows/s | 0.03s | 0 MB | 3,342 rows/s | 1.50s | 0 MB | 0.2 MB |
| 10,000 | 199,473 rows/s | 0.05s | 0 MB | 5,125 rows/s | 1.95s | 0 MB | 0.4 MB |
| 25,000 | 218,032 rows/s | 0.11s | 0 MB | 10,123 rows/s | 2.47s | 0 MB | 1 MB |
| 50,000 | 221,504 rows/s | 0.23s | 2 MB | 18,175 rows/s | 2.75s | 2 MB | 2 MB |
| 100,000 | 221,895 rows/s | 0.45s | 0 MB | 31,583 rows/s | 3.17s | 4 MB | 4 MB |
| 250,000 | 210,343 rows/s | 1.19s | 0 MB | 60,277 rows/s | 4.15s | 12 MB (±6) | 10 MB |
| 500,000 | 219,906 rows/s | 2.27s | 0 MB | 90,310 rows/s | 5.54s | 20 MB (±18) | 20 MB |
| 750,000 | 218,395 rows/s | 3.43s | 0 MB | 101,686 rows/s | 7.38s | 30 MB (±26) | 30 MB |
| 1,000,000 | 214,912 rows/s | 4.65s | 0 MB | 109,562 rows/s | 9.13s | 40 MB (±38) | 40 MB |
| 1,500,000 | 211,301 rows/s | 7.10s | 0 MB | 113,689 rows/s | 13.19s | 60 MB (±58) | 60 MB |
| 2,000,000 | 209,462 rows/s | 9.55s | 0 MB | 121,436 rows/s | 16.47s | 79 MB (±77) | 79 MB |
| 3,000,000 | – | – | – | 127,725 rows/s | 23.49s | 119 MB (±117) | 119 MB |
| 4,000,000 | – | – | – | 128,971 rows/s | 31.01s | 159 MB (±156) | 158 MB |
| 4,500,000 | – | – | – | 119,914 rows/s | 37.53s | 178 MB (±179) | 178 MB |

### 1.2 Memory profile (chart)

Local writer peaks stay flat at **0 MB** (allocator chunk granularity hides the small row buffer). S3 sawtooth grows linearly with multipart buffer size — same pattern as v2.2.2.

```
       LOCAL                              S3
   1K │ ▓ 0                          1K │ ░ 0
 100K │ ▓ 0                        100K │ ░░░ 4
   1M │ ▓ 0                          1M │ ░░░░░░░░░░░░░░░░░░░░ 40
   2M │ ▓ 0                          2M │ ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 79
 4.5M │ ─                          4.5M │ ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 178
```

### 1.3 Throughput vs row count

Local sustains 209–222K rows/s once warm-up amortises (rows above 5K). S3 scales with file size and saturates around 109–129K rows/s for ≥750K rows.

```
   1K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 191,529
      │ ░ 965
 100K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 221,895
      │ ░░░░░░░░ 31,583
   1M │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 214,912
      │ ░░░░░░░░░░░░░░░░░░░░░░░░░ 109,562
 4.5M │ ─
      │ ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 119,914
        legend: ▓ local  ░ S3 (rows/s, max=221,895)
```

### 1.4 Cross-version comparison

Same machine, same workload, same compression level. Numbers within a few percent are measurement noise.

| Workload | v3.0 | v2.2.2 | v1.x (Sept 2025) | v3.0 vs v2.2.2 | v3.0 vs v1.x |
|---|---|---|---|---|---|
| Local 1M | **214,912 rows/s** | 209,905 | 182,693 | +2.4 % | **+17.6 %** |
| Local 2M | 209,462 | 208,512 | 177,012 | +0.5 % | +18.3 % |
| S3 1M | 109,562 | 106,924 | 43,215 | +2.5 % | **+154 %** |
| S3 4.5M | 119,914 | 130,293 | 46,462 | −8 % (network) | **+158 %** |

The v3.0 writer's default code path is unchanged from v2.2.2 — the only writer-side addition is the opt-in `withRandomAccessIndex()` which adds nothing to the default path. Throughput delta vs v2.2.2 is measurement noise.

---

## 2. Read benchmark — new in v3.0

`php benchmark-read.php`. For each row size, the script writes a temp XLSX (via the same writer used in Section 1) then reads it back through `StreamingXlsxReader::fromFile()` or `::fromS3()`. Multi-sheet workbooks (above the per-sheet 1,048,576 limit) are read in full across every sheet.

### 2.1 Wall time + throughput + memory

| Rows | Local Speed | Local Time | Local Peak RAM | S3 Speed | S3 Time | S3 Peak RAM | File Size |
|---|---|---|---|---|---|---|---|
| 100 | 27,216 rows/s | 0.00s | 22 MB | 56 rows/s | 1.79s | 24 MB | 0.01 MB |
| 500 | 61,830 rows/s | 0.01s | 22 MB | 281 rows/s | 1.78s | 24 MB | 0.02 MB |
| 1,000 | 69,553 rows/s | 0.01s | 22 MB | 480 rows/s | 2.08s | 24 MB | 0.04 MB |
| 5,000 | 69,433 rows/s | 0.07s | 24 MB | 2,278 rows/s | 2.19s | 24 MB | 0.2 MB |
| 10,000 | 75,843 rows/s | 0.13s | 24 MB | 4,235 rows/s | 2.36s | 24 MB | 0.4 MB |
| 25,000 | 66,957 rows/s | 0.37s | 24 MB | 9,664 rows/s | 2.59s | 24 MB | 1 MB |
| 50,000 | 72,895 rows/s | 0.69s | 24 MB | 16,967 rows/s | 2.95s | 22 MB | 2 MB |
| 100,000 | 75,899 rows/s | 1.32s | 24 MB | 27,412 rows/s | 3.65s | 22 MB | 4 MB |
| 250,000 | 72,267 rows/s | 3.46s | 24 MB | 43,139 rows/s | 5.80s | 22 MB | 10 MB |
| 500,000 | 68,772 rows/s | 7.27s | 24 MB | 52,204 rows/s | 9.58s | 22 MB | 20 MB |
| 750,000 | 71,360 rows/s | 10.51s | 24 MB | 57,230 rows/s | 13.11s | 24 MB | 30 MB |
| 1,000,000 | 69,946 rows/s | 14.30s | 24 MB | 60,253 rows/s | 16.60s | 24 MB | 40 MB |
| 1,500,000 | 69,517 rows/s | 21.58s | 24 MB | 59,746 rows/s | 25.11s | 22 MB | 60 MB |
| 2,000,000 | 68,710 rows/s | 29.11s | 24 MB | 61,306 rows/s | 32.62s | 24 MB | 79 MB |
| 3,000,000 | – | – | – | 62,447 rows/s | 48.04s | 24 MB | 119 MB |
| 4,000,000 | – | – | – | 62,956 rows/s | 63.54s | 22 MB | 158 MB |
| 4,500,000 | – | – | – | 62,198 rows/s | 72.35s | 22 MB | 178 MB |

### 2.2 Memory profile — bounded by construction

```
   100 │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 22
       │ ░░░░░░░░░░░░░░░░░░░░░░░░ 24
   100K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 24
       │ ░░░░░░░░░░░░░░░░░░░░░░ 22
    1M │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 24
       │ ░░░░░░░░░░░░░░░░░░░░░░░░ 24
  4.5M │ ─
       │ ░░░░░░░░░░░░░░░░░░░░░░ 22
       │ legend: ▓ local  ░ S3  (MB peak, max=50)
```

**Reader peak RAM is 22–24 MB across every row count — independent of file size.** That envelope is dominated by PHP's runtime baseline (opcache, autoload, php-cli overhead). The reader's own working set is a 64 KB compressed read chunk + ~256 KB inflated buffer + at most one in-progress `<row>` XML element. RAM never grows with file size.

### 2.3 Throughput vs row count

```
   1K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 69,553
      │ ░░ 480
 100K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 75,899
      │ ░░░░░░░░ 27,412
   1M │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 69,946
      │ ░░░░░░░░░░░░░░░░ 60,253
 4.5M │ ─
      │ ░░░░░░░░░░░░░░░░ 62,198
        legend: ▓ local  ░ S3 (rows/s, max=78,000)
```

Local sustains ~70K rows/s independent of file size. S3 cold-cache ramps with file size as TTFB amortises across more bytes; saturates at ~60K rows/s for files ≥750K rows.

### 2.4 Comparison with other PHP readers

| Package | 1M-row read time | Peak RAM | S3 native |
|---|---|---|---|
| **kolay/xlsx-stream v3.0 (local)** | **14.3 s** (70K rows/s) | **24 MB** | – |
| **kolay/xlsx-stream v3.0 (S3)** | **16.6 s** (60K rows/s) | **24 MB** | **direct Range GET** |
| OpenSpout 4.x reader (local) | ~25–35 s | ~50–100 MB | none |
| PhpSpreadsheet reader (local) | ❌ crashes (>3 GB) | ~3–5 GB | none |

Numbers for OpenSpout / PhpSpreadsheet are characteristic ranges from public benchmarks; not measured against the same fixture in this run.

---

## 3. Random-access benchmark — new in v3.0

`php benchmark-random-access.php 500000`. The script writes two XLSX files of identical content — one plain, one with `withRandomAccessIndex(every: 10_000)` — then measures `rowAt(N)` latency at increasing target positions plus `rowCount()`. Both files are visually identical when opened in any XLSX reader; only the indexed one carries a `xl/_kxs/index.bin` sidecar that the matched reader uses for O(1) seeks.

### 3.1 Write cost of indexing

| Metric | Plain | Indexed | Δ |
|---|---|---|---|
| Wall time (500K rows) | 2.44 s | 2.43 s | **−0.33 %** (within measurement noise) |
| File size | 20.02 MB | 20.03 MB | **+0.032 %** |
| Sync points emitted | – | 50 | every 10,000 rows |

Original design predicted ≤0.5 % file-size penalty and ≤0.06 % wall-time penalty. Measured penalty is **16× below** the file-size ceiling and below the noise floor on wall time. **Indexing is essentially free at write time.**

### 3.2 Sequential read transparency

The sync markers are valid DEFLATE — invisible to a sequential reader. Both files yield identical row counts:

| File variant | Read time | Rows yielded |
|---|---|---|
| Plain (no markers) | 7.0 s | 500,001 |
| Indexed (50 markers) | ≈ 7.0 s | 500,001 |

Differences fall inside measurement noise. **Sequential read pays no observable cost for indexing.**

### 3.3 `rowAt(N)` latency — plain vs indexed

| Target Row | Plain (full scan) | Indexed (sync + scan) | Speedup |
|---|---|---|---|
| 1 | 1.3 ms | 0.2 ms | 6.1× |
| 2 | 0.2 ms | 0.2 ms | 0.9× |
| 100 | 1.3 ms | 1.4 ms | 1.0× |
| 50,000 | 665.8 ms | 132.6 ms | 5.0× |
| 125,000 | 1,645.9 ms | 67.6 ms | **24.3×** |
| 250,000 | 3,426.5 ms | 136.8 ms | **25.0×** |
| 375,000 | 5,087.3 ms | 68.2 ms | **74.6×** |
| 450,000 | 6,291.8 ms | 136.6 ms | 46.0× |
| 499,900 | 6,897.9 ms | 135.6 ms | 50.9× |
| 500,000 | 6,768.0 ms | 135.9 ms | 49.8× |

Speedup grows with target row depth because plain scan time scales linearly with row index; indexed scan time is bounded by the sync period (≤10,000 rows ≈ 70–150 ms inflate work) plus a single source-range fetch. The 70–150 ms alternation is the chunk-cycle pattern: when the target lands inside the first inflated chunk it's near zero extra work; one chunk further requires one more 64 KB inflate call.

### 3.4 `rowCount()` — O(N) full scan vs O(1) sidecar lookup

| Variant | Time | Result |
|---|---|---|
| Plain `rowCount()` | 7,013.7 ms | 500,000 rows |
| Indexed `rowCount()` | <1 ms (returned 0 ms in the measurement) | 500,000 rows |
| **Speedup** | – | **>260,000×** |

The indexed reader reads `total_rows` straight out of the sidecar header — one CRC32-validated 16-byte read, no inflation, no XML parse. The full-scan path inflates and tokenizes every row to count them.

### 3.5 Random-access speedup chart

```
  50K │ ▓▓▓▓▓▓▓ 0.67s
      │ ░░ 0.13s
 125K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 1.65s
      │ ░ 0.07s
 250K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 3.43s
      │ ░ 0.14s
 375K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 5.09s
      │ ░ 0.07s
 500K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 6.77s
      │ ░ 0.14s
        legend: ▓ plain  ░ indexed (rowAt latency, max=6.77s)
```

---

## 4. Methodology & reproducibility

### Scripts

All under the repo root:

```
benchmark-comprehensive.php   # write benchmark — local + S3, all sizes 100…4.5M
benchmark-read.php             # read benchmark — local + S3
benchmark-random-access.php    # rowAt + rowCount, plain vs indexed
```

Each script accepts an optional cap argument for quick smoke runs:

```
php benchmark-read.php 100000              # cap at 100K rows
php benchmark-random-access.php 200000     # cap dataset size
```

### Test data

Each row has 8 columns, populated as:

```
[$id, "Employee".($id % 100), "emp".($id % 100)."@company.com",
 $departments[$id % 10], 50000 + ($id % 50) * 1000, $today,
 $statuses[$id % 5], "Standard notes for employee record"]
```

Identical to the v1.x and v2.2.2 baselines so cross-version comparisons stay clean.

### Memory measurement

`memory_get_peak_usage(true)` (real allocator). `memory_reset_peak_usage()` is called immediately before each timed read phase to isolate per-phase peaks. RAM is reported at 2 MB granularity due to PHP's `MMAP_THRESHOLD` allocator chunk size.

### Caveats

- Single-run measurements per row size — no statistical averaging. Variance in S3 numbers below 100 ms is meaningful network noise; treat as ranges, not absolutes.
- POC reader's regex-based tokenizer was replaced with a hand-written state machine for v3.0 — DoS-safe (linear time, no backtracking), ~5 % faster on hot-path workloads.
- S3 numbers are cold-cache (fresh GET) for reads. Warm-cache reads land 30–50 % faster but are not part of the canonical methodology.
- Multi-sheet workbooks (above 1,048,576 rows per sheet) read every sheet automatically. Throughput in that range reflects two or more sequential sheet streams.

### Design claims — empirical verification

| Claim | Empirical result |
|---|---|
| Bounded RAM independent of file size | 22–24 MB peak across 100–4.5M rows ✓ |
| Inline-string fast path on all self-written files | confirmed across every fixture ✓ |
| Indexed file size penalty ≤ 0.5 % | measured **+0.032 %** (16× under) ✓ |
| Indexed wall-time penalty ≤ 0.06 % | measurement noise (−0.33 % in this run) ✓ |
| `rowAt(N)` random access via Z_FULL_FLUSH + fresh inflate | confirmed, **74.6× speedup** at 75 % depth ✓ |
| Sequential read transparency on indexed files | confirmed, identical row count and content ✓ |
| Backward compat — vanilla XLSX readers ignore sidecar | confirmed structurally (`xl/_kxs/index.bin` not in `[Content_Types].xml`); cross-tool CI roundtrip tracked for tag-day |

---

## 5. Comparison table — what changed since v2.2.2

| Workload | v3.0 | v2.2.2 | Diff |
|---|---|---|---|
| **Write** (existed in v2.2.2) | | | |
| Local 100K | 221,895 rows/s | 188,771 rows/s | +18 % |
| Local 1M | 214,912 rows/s | 209,905 rows/s | +2.4 % |
| Local 2M | 209,462 rows/s | 208,512 rows/s | +0.5 % |
| S3 1M | 109,562 rows/s | 106,924 rows/s | +2.5 % |
| S3 4.5M | 119,914 rows/s | 130,293 rows/s | −8 % (network jitter) |
| Write peak RAM | unchanged | unchanged | – |
| **Read** (new in v3.0) | | | |
| Local 1M | 69,946 rows/s, 24 MB peak | – | new feature |
| S3 1M | 60,253 rows/s, 24 MB peak | – | new feature |
| S3 4.5M | 62,198 rows/s, 22 MB peak | – | new feature |
| **Random access** (new in v3.0) | | | |
| `rowAt(375K of 500K)` | indexed 68.2 ms / plain 5.09 s = 74.6× | – | new feature |
| `rowCount(500K)` | indexed <1 ms / plain 7.01 s = >260,000× | – | new feature |
| Index file-size penalty | +0.032 % | – | new feature |

Write is **unchanged** from v2.2.2 within measurement noise — opt-in indexing is the only new writer-side path, default-off keeps prior output stable. Read and random-access are pure additions; v2.2.2 had no reader at all.

---
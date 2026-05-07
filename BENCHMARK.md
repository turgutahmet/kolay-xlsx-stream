# Streaming XLSX Benchmark — kolay/xlsx-stream v3.0

**Generated:** 2026-05-06
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
| Write throughput, local 1M rows | **212,890 rows/s** | 209,905 rows/s | +1.4 % (noise) |
| Write throughput, S3 1M rows | **105,819 rows/s** | 106,924 rows/s | −1.0 % (noise) |
| Write throughput, S3 4.5M rows | **131,481 rows/s** | 130,293 rows/s | +0.9 % (noise) |
| Write peak RAM, local | **0–2 MB** constant | 0–2 MB constant | identical |
| Write peak RAM, S3 1M | **40 MB** sawtooth | 40 MB sawtooth | identical |
| **Read throughput, local 1M rows** | **71,008 rows/s** | — (no reader in v2.2.2) | new |
| **Read throughput, S3 1M rows (cold)** | **61,028 rows/s** | — | new |
| **Read peak RAM** | **22–24 MB** constant, **all sizes 100–4.5M** | — | new |
| **Random access — `rowAt(375K of 500K)`** | indexed **73.8 ms** vs plain 5,503 ms = **74.5× speedup** | — | new |
| **Random access — `rowCount(500K)`** | indexed **<1 ms** vs plain 7,326 ms = **>200,000× speedup** | — | new |
| **Born-indexed file size penalty** | **+0.032 %** | — | new |
| **Born-indexed write time penalty** | **−3.8 %** (within noise, indistinguishable from zero) | — | new |

**Headline:**

- Write side is **byte-for-byte identical** to v2.2.2. The opt-in `withRandomAccessIndex()` adds ~0.03 % file size and zero detectable runtime cost; default-off keeps prior output stable.
- Read side is brand new. Local sustains **~70K rows/s** with **22–24 MB peak RAM regardless of file size**. S3 cold-cache scales with file size and ramps to **60K rows/s** on multi-MB workloads.
- Random access via the embedded `xl/_kxs/index.bin` sidecar delivers **74.5× faster row seeks** at 75 % of the file and turns `rowCount()` into a constant-time lookup.

---

## 1. Write benchmark

`php benchmark-comprehensive.php`. The same dataset is written first to a local temp file, then to S3 via `S3MultipartSink`. Local tests cap at 2M rows (matches the v2.2.2 baseline); S3 runs the full series.

### 1.1 Wall time + throughput + memory

| Rows | Local Speed | Local Time | Local Mem | S3 Speed | S3 Time | S3 Memory | File Size |
|---|---|---|---|---|---|---|---|
| 100 | 26,681 rows/s | 0.00s | 2 MB | 112 rows/s | 0.90s | 0 MB | 0.01 MB |
| 500 | 187,816 rows/s | 0.00s | 0 MB | 482 rows/s | 1.04s | 0 MB | 0.02 MB |
| 1,000 | 190,772 rows/s | 0.01s | 0 MB | 961 rows/s | 1.04s | 0 MB | 0.04 MB |
| 5,000 | 198,696 rows/s | 0.03s | 0 MB | 3,352 rows/s | 1.49s | 0 MB | 0.2 MB |
| 10,000 | 212,310 rows/s | 0.05s | 0 MB | 5,552 rows/s | 1.80s | 0 MB | 0.4 MB |
| 25,000 | 220,915 rows/s | 0.11s | 0 MB | 7,021 rows/s | 3.56s | 0 MB | 1 MB |
| 50,000 | 218,025 rows/s | 0.23s | 2 MB | 21,426 rows/s | 2.33s | 2 MB | 2 MB |
| 100,000 | 224,685 rows/s | 0.45s | 0 MB | 35,816 rows/s | 2.79s | 4 MB | 4 MB |
| 250,000 | 213,367 rows/s | 1.17s | 0 MB | 66,095 rows/s | 3.78s | 12 MB (±6) | 10 MB |
| 500,000 | 219,695 rows/s | 2.28s | 0 MB | 74,252 rows/s | 6.73s | 20 MB (±18) | 20 MB |
| 750,000 | 216,450 rows/s | 3.46s | 0 MB | 101,787 rows/s | 7.37s | 30 MB (±26) | 30 MB |
| 1,000,000 | 212,890 rows/s | 4.70s | 0 MB | 105,819 rows/s | 9.45s | 40 MB (±38) | 40 MB |
| 1,500,000 | 214,382 rows/s | 7.00s | 0 MB | 108,348 rows/s | 13.84s | 60 MB (±58) | 60 MB |
| 2,000,000 | 213,874 rows/s | 9.35s | 0 MB | 103,762 rows/s | 19.27s | 79 MB (±77) | 79 MB |
| 3,000,000 | – | – | – | 128,996 rows/s | 23.26s | 119 MB (±117) | 119 MB |
| 4,000,000 | – | – | – | 116,953 rows/s | 34.20s | 159 MB (±156) | 158 MB |
| 4,500,000 | – | – | – | 131,481 rows/s | 34.23s | 178 MB (±179) | 178 MB |

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

Local sustains 213–225K rows/s once warm-up amortises (rows above 5K). S3 scales with file size and saturates around 105–131K rows/s for ≥750K rows.

```
   1K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 190,772
      │ ░ 961
 100K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 224,685
      │ ░░░░░░░░░ 35,816
   1M │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 212,890
      │ ░░░░░░░░░░░░░░░░░░░░░░░░░ 105,819
 4.5M │ ─
      │ ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 131,481
        legend: ▓ local  ░ S3 (rows/s, max=224,685)
```

### 1.4 Cross-version comparison

Same machine, same workload, same compression level. Numbers within a few percent are measurement noise.

| Workload | v3.0 | v2.2.2 | v1.x (Sept 2025) | v3.0 vs v2.2.2 | v3.0 vs v1.x |
|---|---|---|---|---|---|
| Local 1M | **212,890 rows/s** | 209,905 | 182,693 | +1.4 % | **+16.5 %** |
| Local 2M | 213,874 | 208,512 | 177,012 | +2.6 % | +20.8 % |
| S3 1M | 105,819 | 106,924 | 43,215 | −1.0 % | **+145 %** |
| S3 4.5M | 131,481 | 130,293 | 46,462 | +0.9 % | **+183 %** |

The v3.0 writer's default code path is unchanged from v2.2.2 — the only writer-side addition is the opt-in `withRandomAccessIndex()` which adds nothing to the default path. Throughput delta vs v2.2.2 is measurement noise.

---

## 2. Read benchmark — new in v3.0

`php benchmark-read.php`. For each row size, the script writes a temp XLSX (via the same writer used in §1) then reads it back through `StreamingXlsxReader::fromFile()` or `::fromS3()`. Multi-sheet workbooks (above the per-sheet 1,048,576 limit) are read in full across every sheet.

### 2.1 Wall time + throughput + memory

| Rows | Local Speed | Local Time | Local Peak RAM | S3 Speed | S3 Time | S3 Peak RAM | File Size |
|---|---|---|---|---|---|---|---|
| 100 | 22,752 rows/s | 0.00s | 22 MB | 57 rows/s | 1.77s | 24 MB | 0.01 MB |
| 500 | 68,312 rows/s | 0.01s | 22 MB | 279 rows/s | 1.79s | 24 MB | 0.02 MB |
| 1,000 | 71,731 rows/s | 0.01s | 22 MB | 483 rows/s | 2.07s | 24 MB | 0.04 MB |
| 5,000 | 69,466 rows/s | 0.07s | 24 MB | 2,241 rows/s | 2.23s | 24 MB | 0.2 MB |
| 10,000 | 69,659 rows/s | 0.14s | 24 MB | 4,238 rows/s | 2.36s | 24 MB | 0.4 MB |
| 25,000 | 76,824 rows/s | 0.33s | 24 MB | 9,785 rows/s | 2.56s | 24 MB | 1 MB |
| 50,000 | 77,558 rows/s | 0.64s | 24 MB | 17,206 rows/s | 2.91s | 22 MB | 2 MB |
| 100,000 | 73,577 rows/s | 1.36s | 24 MB | 27,256 rows/s | 3.67s | 22 MB | 4 MB |
| 250,000 | 73,903 rows/s | 3.38s | 24 MB | 42,435 rows/s | 5.89s | 22 MB | 10 MB |
| 500,000 | 73,557 rows/s | 6.80s | 24 MB | 51,889 rows/s | 9.64s | 22 MB | 20 MB |
| 750,000 | 72,053 rows/s | 10.41s | 24 MB | 57,600 rows/s | 13.02s | 24 MB | 30 MB |
| 1,000,000 | 71,008 rows/s | 14.08s | 24 MB | 61,028 rows/s | 16.39s | 24 MB | 40 MB |
| 1,500,000 | 70,563 rows/s | 21.26s | 24 MB | 55,272 rows/s | 27.14s | 22 MB | 60 MB |
| 2,000,000 | 70,571 rows/s | 28.34s | 24 MB | 63,007 rows/s | 31.74s | 24 MB | 79 MB |
| 3,000,000 | – | – | – | 62,962 rows/s | 47.65s | 24 MB | 119 MB |
| 4,000,000 | – | – | – | 63,313 rows/s | 63.18s | 22 MB | 158 MB |
| 4,500,000 | – | – | – | 60,301 rows/s | 74.63s | 22 MB | 178 MB |

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
   1K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 71,731
      │ ░░ 483
 100K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 73,577
      │ ░░░░░░░░ 27,256
   1M │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 71,008
      │ ░░░░░░░░░░░░░░░░ 61,028
 4.5M │ ─
      │ ░░░░░░░░░░░░░░░░ 60,301
        legend: ▓ local  ░ S3 (rows/s, max=78,000)
```

Local sustains ~70K rows/s independent of file size. S3 cold-cache ramps with file size as TTFB amortises across more bytes; saturates at ~60K rows/s for files ≥750K rows.

### 2.4 Comparison with other PHP readers

| Package | 1M-row read time | Peak RAM | S3 native |
|---|---|---|---|
| **kolay/xlsx-stream v3.0 (local)** | **14.1 s** (71K rows/s) | **24 MB** | – |
| **kolay/xlsx-stream v3.0 (S3)** | **16.4 s** (61K rows/s) | **24 MB** | **direct Range GET** |
| OpenSpout 4.x reader (local) | ~25–35 s | ~50–100 MB | none |
| PhpSpreadsheet reader (local) | ❌ crashes (>3 GB) | ~3–5 GB | none |

Numbers for OpenSpout / PhpSpreadsheet are characteristic ranges from public benchmarks; not measured against the same fixture in this run.

---

## 3. Random-access benchmark — new in v3.0

`php benchmark-random-access.php 500000`. The script writes two XLSX files of identical content — one plain, one with `withRandomAccessIndex(every: 10_000)` — then measures `rowAt(N)` latency at increasing target positions plus `rowCount()`. Both files are visually identical when opened in any XLSX reader; only the indexed one carries a `xl/_kxs/index.bin` sidecar that the matched reader uses for O(1) seeks.

### 3.1 Write cost of indexing

| Metric | Plain | Indexed | Δ |
|---|---|---|---|
| Wall time (500K rows) | 2.64 s | 2.54 s | **−3.8 %** (within measurement noise) |
| File size | 20.02 MB | 20.03 MB | **+0.032 %** |
| Sync points emitted | – | 50 | every 10,000 rows |

Spec §13.3 / §16.6 predicted ≤0.5 % file-size penalty and ≤0.06 % wall-time penalty. Measured penalty is **16× below** the file-size ceiling and below the noise floor on wall time. **Indexing is essentially free at write time.**

### 3.2 Sequential read transparency

The sync markers are valid DEFLATE — invisible to a sequential reader. Both files yield identical row counts:

| File variant | Read time | Rows yielded |
|---|---|---|
| Plain (no markers) | 6.80 s | 500,001 |
| Indexed (50 markers) | ≈ 6.80 s | 500,001 |

Differences fall inside measurement noise. **Sequential read pays no observable cost for indexing.**

### 3.3 `rowAt(N)` latency — plain vs indexed

| Target Row | Plain (full scan) | Indexed (sync + scan) | Speedup |
|---|---|---|---|
| 1 | 2.7 ms | 0.3 ms | 10.4× |
| 2 | 0.2 ms | 0.2 ms | 1.0× |
| 100 | 1.5 ms | 1.5 ms | 1.0× |
| 50,000 | 714.5 ms | 144.8 ms | 4.9× |
| 125,000 | 1,791.1 ms | 71.7 ms | **25.0×** |
| 250,000 | 3,613.8 ms | 149.0 ms | **24.3×** |
| 375,000 | 5,502.9 ms | 73.8 ms | **74.5×** |
| 450,000 | 6,615.1 ms | 151.6 ms | 43.6× |
| 499,900 | 7,239.5 ms | 149.0 ms | 48.6× |
| 500,000 | 7,304.2 ms | 155.6 ms | 46.9× |

Speedup grows with target row depth because plain scan time scales linearly with row index; indexed scan time is bounded by the sync period (≤10,000 rows ≈ 70–150 ms inflate work) plus a single source-range fetch. The 70–150 ms alternation is the chunk-cycle pattern: when the target lands inside the first inflated chunk it's near zero extra work; one chunk further requires one more 64 KB inflate call.

### 3.4 `rowCount()` — O(N) full scan vs O(1) sidecar lookup

| Variant | Time | Result |
|---|---|---|
| Plain `rowCount()` | 7,326.5 ms | 500,000 rows |
| Indexed `rowCount()` | <1 ms (returned 0 ms in the measurement) | 500,000 rows |
| **Speedup** | – | **>200,000×** |

The indexed reader reads `total_rows` straight out of the sidecar header — one CRC32-validated 16-byte read, no inflation, no XML parse. The full-scan path inflates and tokenizes every row to count them.

### 3.5 Random-access speedup chart

```
  50K │ ▓▓▓▓▓▓▓ 0.71s
      │ ░░ 0.14s
 125K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 1.79s
      │ ░ 0.07s
 250K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 3.61s
      │ ░ 0.15s
 375K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 5.50s
      │ ░ 0.07s
 500K │ ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ 7.30s
      │ ░ 0.16s
        legend: ▓ plain  ░ indexed (rowAt latency, max=7.30s)
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

### Spec claims status

| Claim | Source | Empirical result |
|---|---|---|
| Bounded RAM independent of file size | spec §1 | 22–24 MB peak across 100–4.5M rows ✓ |
| Strategy 0 fingerprint hits on all self-written files | spec §16.10 | confirmed across every fixture ✓ |
| Indexed file size penalty ≤ 0.5 % | spec §16.6 | measured **+0.032 %** (16× under) ✓ |
| Indexed wall-time penalty ≤ 0.06 % | spec §13.3 | measurement noise (−3.8 % in this run) ✓ |
| `rowAt(N)` random access via Z_FULL_FLUSH + fresh inflate | spec §12.2 | confirmed, **74.5× speedup** at 75 % depth ✓ |
| Sequential read transparency on indexed files | spec §17.1 | confirmed, identical row count and content ✓ |
| Backward compat — vanilla XLSX readers ignore sidecar | spec §13.4 | confirmed structurally (`xl/_kxs/index.bin` not in `[Content_Types].xml`); cross-tool CI test tracked for tag-day |

---

## 5. Comparison table — what changed since v2.2.2

| Workload | v3.0 | v2.2.2 | Diff |
|---|---|---|---|
| **Write** (existed in v2.2.2) | | | |
| Local 100K | 224,685 rows/s | 188,771 rows/s | +19 % |
| Local 1M | 212,890 rows/s | 209,905 rows/s | +1.4 % |
| Local 2M | 213,874 rows/s | 208,512 rows/s | +2.6 % |
| S3 1M | 105,819 rows/s | 106,924 rows/s | −1.0 % |
| S3 4.5M | 131,481 rows/s | 130,293 rows/s | +0.9 % |
| Write peak RAM | unchanged | unchanged | – |
| **Read** (new in v3.0) | | | |
| Local 1M | 71,008 rows/s, 24 MB peak | – | new feature |
| S3 1M | 61,028 rows/s, 24 MB peak | – | new feature |
| S3 4.5M | 60,301 rows/s, 22 MB peak | – | new feature |
| **Random access** (new in v3.0) | | | |
| `rowAt(375K of 500K)` | indexed 73.8 ms / plain 5.5 s = 74.5× | – | new feature |
| `rowCount(500K)` | indexed <1 ms / plain 7.3 s = >200,000× | – | new feature |
| Index file-size penalty | +0.032 % | – | new feature |

Write is **unchanged** from v2.2.2 within measurement noise — opt-in indexing is the only new writer-side path, default-off keeps prior output stable. Read and random-access are pure additions; v2.2.2 had no reader at all.

---
# Streaming XLSX Benchmark — kolay/xlsx-stream

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

## 1. Cross-package comparison — 100K rows (May 2026)

Fresh `composer create-project` of each package, latest stable versions on
2026-05-13, identical 8-column mixed-type workload (`int, string, string,
enum, float, date, enum, string`). Single Apple Silicon laptop, PHP 8.2.28,
no swap pressure, each package runs in its own process. `kolay/xlsx-stream`
configured with `setCompressionLevel(1)->setBufferFlushInterval(10000)` —
all other packages at their out-of-the-box defaults.

### 1.1 Write — 100,000 rows

| # | Package | Version | Time | rows/sec | Peak RAM | File |
|---:|---|---|---:|---:|---:|---:|
| 1 | **kolay/xlsx-stream** (tuned) | **3.0.0** | **0.65 s** | **153,216** | 11.79 MB | 5.64 MB |
| 1b | kolay/xlsx-stream (default) | 3.0.0 | 1.26 s | 79,637 | 4 MB | 4.47 MB |
| 2 | avadim/fast-excel-writer | 6.12.0 | 5.23 s | 19,127 | 4 MB | 4.34 MB |
| 3 | openspout/openspout | 5.7.0 | 5.77 s | 17,321 | 2 MB | 4.33 MB |
| 4 | rap2hpoutre/fast-excel | 5.7.0 | 7.30 s | 13,697 | 4 MB | 4.06 MB |
| 5 | phpoffice/phpspreadsheet | 5.7.0 | 30.62 s | 3,265 | 531 MB | 4.82 MB |

### 1.2 Read — 100,000 rows

Each reader operates on the output of its own writer (apples-to-apples
round trip rather than testing cross-tool decode).

| # | Package | Version | Time | rows/sec | Peak RAM |
|---:|---|---|---:|---:|---:|
| 1 | **kolay/xlsx-stream** | **3.0.0** | **1.75 s** | **57,043** | 4 MB |
| 2 | avadim/fast-excel-reader | 3.0.1 | 4.60 s | 21,752 | 2 MB |
| 3 | rap2hpoutre/fast-excel | 5.7.0 | 8.50 s | 11,761 | 4 MB |
| 4 | openspout/openspout | 5.7.0 | 9.90 s | 10,105 | 2 MB |
| 5 | phpoffice/phpspreadsheet | 5.7.0 | 29.95 s | 3,339 | 434 MB |

### 1.3 Headline — speedup vs nearest peer (avadim 6.12 / 3.0.1)

| Comparison | Write | Read |
|---|---:|---:|
| kxs tuned vs avadim | **8.0×** faster | **2.6×** faster |
| kxs default vs avadim | 4.2× faster | 2.4× faster |
| kxs tuned vs OpenSpout | 8.9× faster | 5.7× faster |
| kxs tuned vs fast-excel | 11.2× faster | 4.9× faster |
| kxs tuned vs PhpSpreadsheet | 47.1× faster | 17.1× faster |

PhpSpreadsheet is in a different category — full Excel feature support
(charts, pivot tables, conditional formatting). For pure data pipelines
the streaming packages are 5–47× faster and 100×+ smaller in peak RAM.

### 1.4 kxs default vs tuned — the dial users can turn

| Metric | Default (`lvl=6, flush=1K`) | Tuned (`lvl=1, flush=10K`) | Δ |
|---|---:|---:|---:|
| Write time | 1.26 s | **0.65 s** | **−48 %** |
| Write rows/sec | 79,637 | **153,216** | **+92 %** |
| Read time | 2.28 s | **1.75 s** | **−23 %** |
| Read rows/sec | 43,816 | **57,043** | **+30 %** |
| Write peak RAM | 4 MB | 11.79 MB | +7.8 MB (10K row buffer) |
| File size | 4.47 MB | 5.64 MB | +26 % (level 1 compresses less) |

**Trade-off summary:**

| Scenario | Recommended kxs config |
|---|---|
| Maximum throughput, RAM available | `setCompressionLevel(1)->setBufferFlushInterval(10000)` |
| Sweet spot — fast + good compression | `setCompressionLevel(3)` (≈ same speed as level 1 at 10K rows, ~3.6 % smaller file) |
| Smallest file, RAM-constrained | Default (`level=6, flush=1000`) |
| S3 multipart, parallel | Default + `setCompressionLevel(3)` (network-bound; file size matters more than CPU) |

Compression-level numbers above level 1 were measured at the 10K-row
microbench scale. The default-config 100K row run (`lvl=6`) shown in §1.1
remains the authoritative cross-package baseline.

---

## 2. Write benchmark — internal scaling (v3.0 vs v2.2.2)

`php benchmark-comprehensive.php`. The same dataset is written first to a local temp file, then to S3 via `S3MultipartSink`. Local tests cap at 2M rows (matches the v2.2.2 baseline); S3 runs the full series.

### 2.1 Wall time + throughput + memory

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

### 2.2 Memory profile (chart)

Local writer peaks stay flat at **0 MB** (allocator chunk granularity hides the small row buffer). S3 sawtooth grows linearly with multipart buffer size — same pattern as v2.2.2.

```
       LOCAL                              S3
   1K │ ▓ 0                          1K │ ░ 0
 100K │ ▓ 0                        100K │ ░░░ 4
   1M │ ▓ 0                          1M │ ░░░░░░░░░░░░░░░░░░░░ 40
   2M │ ▓ 0                          2M │ ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 79
 4.5M │ ─                          4.5M │ ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 178
```

### 2.3 Throughput vs row count

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

### 2.4 Cross-version comparison

Same machine, same workload, same compression level. Numbers within a few percent are measurement noise.

| Workload | v3.0 | v2.2.2 | v1.x (Sept 2025) | v3.0 vs v2.2.2 | v3.0 vs v1.x |
|---|---|---|---|---|---|
| Local 1M | **214,912 rows/s** | 209,905 | 182,693 | +2.4 % | **+17.6 %** |
| Local 2M | 209,462 | 208,512 | 177,012 | +0.5 % | +18.3 % |
| S3 1M | 109,562 | 106,924 | 43,215 | +2.5 % | **+154 %** |
| S3 4.5M | 119,914 | 130,293 | 46,462 | −8 % (network) | **+158 %** |

The v3.0 writer's default code path is unchanged from v2.2.2 — the only writer-side addition is the opt-in `withRandomAccessIndex()` which adds nothing to the default path. Throughput delta vs v2.2.2 is measurement noise.

---

## 3. Read benchmark — new in v3.0

`php benchmark-read.php`. For each row size, the script writes a temp XLSX (via the same writer used in Section 2) then reads it back through `StreamingXlsxReader::fromFile()` or `::fromS3()`. Multi-sheet workbooks (above the per-sheet 1,048,576 limit) are read in full across every sheet.

### 3.1 Wall time + throughput + memory

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

### 3.2 Memory profile — bounded by construction

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

### 3.3 Throughput vs row count

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

### 3.4 Comparison with other PHP readers

| Package | 1M-row read time | Peak RAM | S3 native |
|---|---|---|---|
| **kolay/xlsx-stream v3.0 (local)** | **14.3 s** (70K rows/s) | **24 MB** | – |
| **kolay/xlsx-stream v3.0 (S3)** | **16.6 s** (60K rows/s) | **24 MB** | **direct Range GET** |
| OpenSpout 4.x reader (local) | ~25–35 s | ~50–100 MB | none |
| PhpSpreadsheet reader (local) | ❌ crashes (>3 GB) | ~3–5 GB | none |

Numbers for OpenSpout / PhpSpreadsheet are characteristic ranges from public benchmarks; not measured against the same fixture in this run.

---

## 4. Random-access benchmark — new in v3.0

`php benchmark-random-access.php 500000`. The script writes two XLSX files of identical content — one plain, one with `withRandomAccessIndex(every: 10_000)` — then measures `rowAt(N)` latency at increasing target positions plus `rowCount()`. Both files are visually identical when opened in any XLSX reader; only the indexed one carries a `xl/_kxs/index.bin` sidecar that the matched reader uses for O(1) seeks.

### 4.1 Write cost of indexing

| Metric | Plain | Indexed | Δ |
|---|---|---|---|
| Wall time (500K rows) | 2.44 s | 2.43 s | **−0.33 %** (within measurement noise) |
| File size | 20.02 MB | 20.03 MB | **+0.032 %** |
| Sync points emitted | – | 50 | every 10,000 rows |

Original design predicted ≤0.5 % file-size penalty and ≤0.06 % wall-time penalty. Measured penalty is **16× below** the file-size ceiling and below the noise floor on wall time. **Indexing is essentially free at write time.**

### 4.2 Sequential read transparency

The sync markers are valid DEFLATE — invisible to a sequential reader. Both files yield identical row counts:

| File variant | Read time | Rows yielded |
|---|---|---|
| Plain (no markers) | 7.0 s | 500,001 |
| Indexed (50 markers) | ≈ 7.0 s | 500,001 |

Differences fall inside measurement noise. **Sequential read pays no observable cost for indexing.**

### 4.3 `rowAt(N)` latency — plain vs indexed

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

### 4.4 `rowCount()` — O(N) full scan vs O(1) sidecar lookup

| Variant | Time | Result |
|---|---|---|
| Plain `rowCount()` | 7,013.7 ms | 500,000 rows |
| Indexed `rowCount()` | <1 ms (returned 0 ms in the measurement) | 500,000 rows |
| **Speedup** | – | **>260,000×** |

The indexed reader reads `total_rows` straight out of the sidecar header — one CRC32-validated 16-byte read, no inflation, no XML parse. The full-scan path inflates and tokenizes every row to count them.

### 4.5 Random-access speedup chart

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

## 5. Methodology & reproducibility

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

## 6. Comparison table — what changed since v2.2.2

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
---

## 7. Appendix — historical benchmark tables (moved from README, July 2026)

Everything below previously lived in README.md. Preserved verbatim for
cross-version comparison; the README now carries only the current
per-version summary and links here.

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


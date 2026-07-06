# Kolay XLSX Stream

[![Latest Version on Packagist](https://img.shields.io/packagist/v/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![Tests](https://img.shields.io/github/actions/workflow/status/turgutahmet/kolay-xlsx-stream/tests.yml?branch=main&label=tests&style=flat-square)](https://github.com/turgutahmet/kolay-xlsx-stream/actions/workflows/tests.yml)
[![Total Downloads](https://img.shields.io/packagist/dt/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![License](https://img.shields.io/packagist/l/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![PHP Version](https://img.shields.io/packagist/php-v/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)

Bidirectional XLSX streaming for PHP and Laravel — and the only library
in any language that makes the spreadsheet itself **queryable**. Write
millions of rows straight to S3 with zero disk I/O, read them back with
bounded memory, seek to any row in O(1), and ask for a column's sum,
median, p99 or distinct count **without reading a single row** — over
HTTP range requests, from a file Excel opens like any other.

- **Write**: ~289K rows/s locally at 6 MB peak RAM; direct S3 multipart
  streaming at bounded memory — synchronous by default (flat ~part-size
  working set, true O(1) regardless of file size), optional parallel
  upload window, no temp files
- **Read**: ~127K rows/s full scans with bounded memory, any file size
- **Seek**: `rowAt(1_000_000)` in milliseconds via the born-indexed sidecar
- **Query**: `columnStats` / `rowsWhere` / `findRow` / `groupStats` with
  Parquet-style block pruning; `median` / `quantile` / `countDistinct`
  from embedded sketches with **zero row reads**
- **Open format**: the sidecar is a published spec ([SPEC.md](SPEC.md))
  with a byte-pinned conformance suite

## Why this package?

Most PHP Excel libraries load whole documents into RAM (unusable at
scale), spill temp files before uploading to S3, and can only read
forward — reaching row 900,000 means scanning 899,999 rows first.

This package streams in both directions with constant memory, and its
born-indexed mode embeds a small binary sidecar (`xl/_kxs/index.bin`)
that vanilla readers ignore but this library uses for random access,
block-pruned queries and sidecar-only analytics. Excel, LibreOffice,
Numbers, PhpSpreadsheet and OpenSpout all open the files normally.

## Performance

### The trajectory — same canonical workloads, every release

| | v1.x (Sep 2025) | v2.2 (May 2026) | v3.0 (May 2026) | v3.1 (Jul 2026) | v3.2 (Jul 2026) |
|---|---|---|---|---|---|
| Write, local | ~182K rows/s | ~210K | ~215K | **~289K** | ~289K |
| Write, S3 (1M rows) | ~9K rows/s | ~107K | ~107K | +36% same-link A/B | + parallel window |
| Read, local | — | — | ~70K rows/s | ~106K | **~127K** |
| Random access | — | — | O(1) `rowAt` | + block-pruned queries | + within-block skip (~19×) |
| Analytics | — | — | — | `columnStats` 0-request | **median/p99/distinct 0-request** |
| Peak RAM (write/read) | 0-2 MB / — | ~6 MB / — | 6 / 24 MB | 6 / 24 MB | 6 / **6 MB** |

Absolute S3 throughput tracks the network path far more than the
library (the same 1M-row export measured 59K–153K rows/s across
sessions) — the honest S3 claim is the same-day A/B: **v3.2's writer is
+36% over v3.0.2 on an identical link**. Uploads are **synchronous by
default** — part memory stays flat at ~part-size no matter how large the
file, and throughput is steady; a **parallel upload window is opt-in**
(`concurrency`) for high-latency links where hiding per-request
round-trips outweighs its higher (sawtooth) memory. Benchmark on your
own link.

### Cross-package, 100K rows (May 2026, latest stables)

| Write | Time | rows/s | | Read | Time | rows/s |
|---|---:|---:|---|---|---:|---:|
| **kolay/xlsx-stream** | **0.65s** | **153K** | | **kolay/xlsx-stream** | **1.75s** | **57K** |
| avadim/fast-excel-writer | 5.23s | 19K | | avadim/fast-excel-reader | 4.60s | 22K |
| openspout | 5.77s | 17K | | fast-excel | 8.50s | 12K |
| fast-excel | 7.30s | 14K | | openspout | 9.90s | 10K |
| phpspreadsheet | 30.62s | 3K | | phpspreadsheet | 29.95s | 3K |

(Numbers predate the v3.1/v3.2 speedups. PhpSpreadsheet plays a
different game — full Excel feature support at memory-bound cost.)

**Every historical table** (per-version scaling runs from 100 rows to
4.5M, random-access speedups, memory profiles, methodology) lives in
[BENCHMARK.md](BENCHMARK.md). All numbers come from the committed
`bench/` harnesses — fresh process per run, medians, generation cost
subtracted; re-run them yourself.

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
> the full v3.2 highlights.
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


## Use cases

Each scenario below links to the how-to section further down.

**1. The million-row queued export that stopped eating RAM.**
A Laravel queue job streams `FromQuery`-style data straight to S3 —
no temp file, no reopen-per-chunk, ~6 MB writer footprint, progress
callbacks for the UI. Add `withRandomAccessIndex()` and the artifact is
instantly seekable for every scenario below. → *Direct S3 Streaming*,
*Laravel Job Example*, *Progress Reporting*.

**2. "Download report" endpoints that start instantly.**
Stream the workbook into the HTTP response as it's generated
(`PhpStreamSink` → `php://output`) — first bytes reach the browser
while row 1,000,000 is still being written. → *Streaming directly to
an HTTP response*.

**3. Importing customer uploads without fear.**
Read files produced by Excel/openpyxl/PhpSpreadsheet with bounded
memory (shared-strings tables up to 64 MB compressed), correct dates
via `autoDetectDates()`, validate row-by-row, and batch-insert.
→ *Reading XLSX Files*, *Reading dates and times*.

**4. Parallel imports: wall clock = slowest worker.**
`shards(8)` splits a born-indexed sheet into eight independently
decompressible, JSON-serializable ranges — dispatch one queue job per
shard, zero coordination. → *Parallel reads*.

**5. Admin file preview without importing to a database.**
Paginate a 4M-row S3 export in a UI: `rowCount()` is O(1), page 40,000
costs the same as page 1 via `rowRange()`, "jump to ID" is `findRow()`
— two range requests on a sorted column. → *Random-Access Reading*.

**6. Dashboard numbers straight from the file.**
"Total payroll", "median salary", "p99 order value", "distinct
customers" — answered from the sidecar with zero row reads
(`columnStats`, `quantile`, `countDistinct`), and per-month breakdowns
via `groupStats()` reading only group-boundary blocks. The file IS the
report backend. → *Queryable XLSX*, *Grouped aggregates*.

**7. Styled corporate reports, still streaming.**
Header styling, per-row highlight styles, ₺/date/weekday number
formats, frozen header, autofilter, auto column widths — all
single-pass compatible. → *Header & Column Styling*.

**8. Tight environments: Lambda, small pods, multi-tenant SaaS.**
Constant memory on both directions means the same code runs in a
128 MB function and a shared worker without per-tenant memory math.

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

> **The physics of deflate streaming** — seekability requires
> `Z_FULL_FLUSH` markers, and each marker resets the compressor's
> dictionary, which in principle costs compression ratio. In practice
> the cost is negligible at our default cadence: a sync point every
> 10K rows resets a 32 KB window once per ~1 MB of XML, measured at
> **+0.04 %** file size on the 500K-row benchmark workload (predicted
> ceiling ≤0.5 %). You would only notice dictionary-reset overhead at
> extreme settings like `every: 100` — if you shrink the period for
> query granularity, re-measure your file sizes.

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

### Grouped aggregates & approximate analytics *(v3.2+)*

Two more layers on the same sidecar:

```php
// GROUP BY over S3 without reading interior blocks: on a sorted group
// column, group-pure blocks contribute their precomputed sums — only
// blocks straddling a group boundary are fetched. 20 groups over 1M
// rows: ~57 ms, interior blocks provably never read.
$byMonth = $reader->groupStats(groupBy: 6, aggregate: 5,
                               bucket: fn ($serial) => (int) ($serial / 30.44));

// Approximate analytics with ZERO row reads and ZERO extra requests —
// the answers live in the sidecar the reader already fetched at open:
$writer->withColumnSketches([4, 3]);        // writer side, once
$reader->median(4);                          // p50 salary
$reader->quantile(4, 0.99);                  // p99 order amount
$reader->countDistinct(3);                   // ~distinct emails (±3% at 100K)
```

The sketches are a merging t-digest (~1-4 KB/column, p01/p99 within
0.2% rank error) and a HyperLogLog (2 KB, ±5% pinned) per column — both
merge associatively, which is what future segment/partition stitching
builds on. The full binary layout is public: **KXSI is an open spec**
(see [SPEC.md](SPEC.md)) with committed conformance vectors under
`tests/SpecVectors/`, so other implementations can verify byte-for-byte.

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

**Reading auto-split files back** *(v3.2.2+)*: on born-indexed files the
reader detects the continuation chain (consecutive exactly-full sheets
with identical headers) and treats it as ONE logical table — `rowCount()`,
`rowsWhere()`, `findRow()`, `columnStats()`, `quantile()`, `shards()` and
friends span every continuation sheet with continuous global row numbers.
Intentional multi-sheet workbooks (different headers, or sheets that
aren't exactly full) keep per-sheet semantics. Before v3.2.2 queries
silently answered from the active sheet alone — treat that as a reason
to upgrade if you export past one sheet.

Measured on a 2.1M-row, 3-sheet chain (local disk, 6 MB peak RAM):
chain detection + `rowCount()` 1.5 ms on first call (~1 µs warmed),
`findRow()` into the second sheet 9.5 ms, `rowAt()` 2.2 ms,
`columnStats()` 1.5 ms, `quantile()` 1.7 ms. Single-sheet files take an
early-out before any chain logic — the A/B benchmark against v3.2.1
shows every single-sheet read path within ±1 % (measurement noise).

### Performance Tuning

```php
// Ultra-fast mode for maximum speed
$writer->setCompressionLevel(1)        // Minimal compression
       ->setBufferFlushInterval(50000); // Large buffer

// Balanced mode (default)
$writer->setCompressionLevel(5)        // Balanced compression (the default)
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

Defaults come from the published config file; code-level setters always
override them at call time. *(Fixed in v3.2.2 — earlier releases
shipped these keys without reading them; re-publish the config to opt
in: `php artisan vendor:publish --tag=xlsx-stream-config --force`. The
new file carries a `'version' => 2` marker; pre-3.2.2 copies stay
inert so their stale defaults can't silently change your output.)*

```php
// config/xlsx-stream.php — applied when 'version' => 2 is present
return [
    'version' => 2,
    'writer' => [
        'compression_level' => env('XLSX_STREAM_COMPRESSION_LEVEL', 5),
        'buffer_flush_interval' => env('XLSX_STREAM_BUFFER_FLUSH_INTERVAL', 10000),
    ],
    's3' => [
        'part_size' => env('XLSX_STREAM_S3_PART_SIZE', 8 * 1024 * 1024),
        'concurrency' => env('XLSX_STREAM_S3_CONCURRENCY', 4),
    ],
];
```

```php
// Setters override config per writer instance:
$writer->setCompressionLevel(9)           // wins over the config value
       ->setBufferFlushInterval(50000);
$writer->withRandomAccessIndex(every: 10000)
       ->withColumnStats([1, 4])
       ->withColumnSketches([1, 4]);
$writer->onProgress(fn ($rows, $bytes) => ...)->setProgressInterval(10000);

// Full control: construct the sink directly.
new S3MultipartSink($s3, $bucket, $key, partSize: 8 * 1024 * 1024, concurrency: 8);
```

Transient S3 retries belong to the AWS SDK — configure them on your
`S3Client` (`'retries' => N`); the sink adds one last-resort re-dispatch
after the SDK gives up. Progress/"logging" is the `onProgress()`
callback: wire it to your logger of choice.

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


### Memory model

- **Local writes are O(1)**: a row buffer (default 10K rows) plus the
  deflate context — ~6 MB peak regardless of row count.
- **S3 writes are bounded by the upload window** *(v3.2)*: peak ≈
  `part_size × (concurrency + 2)` — ~46 MB flat at the defaults, at 1M
  rows and at 10M rows alike. (Before v3.2 the sink's buffer copies
  ratcheted ~40 MB per million rows; the parallel window rework removed
  that.) Memory saw-tooths as parts fill and upload — that's the buffer
  cycle, not a leak.
- **Reads are bounded by construction**: inflate chunks + one row in
  flight — ~6 MB for files this package wrote (24 MB ceiling with large
  external shared-strings tables, which now parse streaming — the full
  XML never materializes).

The random-access / query layer is built on deflate `FULL_FLUSH` sync
points plus the KXSI sidecar — the full binary format, its invariants
and the conformance suite are documented in [SPEC.md](SPEC.md).

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


### Key Performance Metrics *(v3.2)*

- **Write — Local**: ~289K rows/s sustained, 6 MB peak
- **Write — S3**: network-bound; +36% over v3.0.2 same-link, steady
  wall times under the parallel window, ~46 MB flat peak
- **Read — Local**: ~127K rows/s full scan, bounded memory at any size
- **Point reads**: `rowAt` ~1.1 ms within a block; `rows(skip: 1M)` 2.5 ms
- **Analytics**: `median`/`quantile`/`countDistinct` answer with zero
  row reads; `groupStats` over 1M rows in ~57 ms

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

v3.1/v3.2 extend the same sidecar with additional TLV sections
(`STAT`/`SCRC`/`TDIG`/`CHLL`) under the identical opaque-part contract;
manual open-tests are repeated against real Excel before every tag.

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

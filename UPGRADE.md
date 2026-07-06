# Upgrade Guide

## Upgrading from v3.2.2 to v3.3.0

No breaking API changes — every call site works unchanged, the base
`Source` contract is unchanged, and classic writer output stays
byte-identical (the new `compact()` and `syncAtGroupBoundaries()` modes
are opt-in). There is **one behavioral change** to know about, plus a lot
of new opt-in capability.

### 1. S3 multipart uploads are now synchronous by default (memory fix)

`S3MultipartSink`'s default `concurrency` changed from **4 to 1**. This
fixes a real memory leak: the parallel (async) upload path let the AWS
SDK's promise graph retain every part's body, so S3 write memory grew
with the file (~130 MB at 3M rows, unbounded beyond) — the advertised
"stream millions of rows to S3 with bounded memory" was silently
O(file size). The synchronous default keeps memory flat at ~part size
regardless of file size.

- **You almost certainly want the new default.** On bandwidth-bound links
  it is no slower than the old parallel default (both are network-bound)
  and it is steadier.
- **To restore parallel uploads** (only worth it on a high-latency link
  where you have measured a throughput win), opt back in explicitly:

  ```php
  // Per sink:
  new S3MultipartSink($s3, $bucket, $key, concurrency: 4);
  // Or via config/env for forDisk() writers:
  // XLSX_STREAM_S3_CONCURRENCY=4
  ```

  The parallel path now runs a per-part `gc_collect_cycles()` to bound its
  footprint, but it is still a higher, sawtooth memory profile than the
  synchronous default.

### 2. New opt-in capabilities (nothing to change)

All additive — adopt as you like:

- **Query engine:** `rowsWhereAll()`, `estimatedRows()`/`explain()`,
  `topRows()`, `sampleRows()`, column addressing by header name
  (`columnStats('amount')`), `Readers\Bucket::month()` for `groupStats()`.
- **Integrity:** `$reader->verify()` (block-granularity CRC check);
  `new S3MultipartSink(..., verifyParts: true)` so S3 rejects a corrupted
  part.
- **Writer ergonomics:** `queryable([1,2,3])` (index + zone maps +
  sketches in one call); `compact()` (r-less cells, ~52–62 % smaller
  compressed sheets — opt-in while it accrues field mileage);
  `syncAtGroupBoundaries(col)` (zero-scan `groupStats` on grouped exports).
- **Observability:** `$reader->onFullScan(fn ($ctx) => …)` to detect when
  a query isn't using the index.

If you implement your own `Source`, note that bounded ranged reads are a
new optional capability interface (`SupportsBoundedStream`) — you do NOT
need to implement it; sources without it simply fetch to EOF and the
reader caps the read.

## Upgrading from v3.2.0 / v3.2.1 to v3.2.2

No breaking API changes — every call site works unchanged. One
correctness change you should know about, and one opt-in.

### 1. Queries on auto-split workbooks now cover the whole table

If a file was written past Excel's 1,048,576-row sheet ceiling (the
writer auto-splits it into continuation sheets), the query surface —
`rowCount()`, `rows()`, `rowAt()`, `rowRange()`, `rowsWhere()`,
`findRow()`, `columnStats()`, `groupStats()`, `quantile()`/`median()`,
`countDistinct()`, `shards()` — now answers for the whole continuation
chain with continuous global row numbers, instead of silently covering
only the active sheet.

- If you never exceed one sheet, nothing changes (verified within ±1 %
  of v3.2.1 on every single-sheet read path).
- Intentional multi-sheet workbooks (`newSheet()` tables with different
  headers, or same-schema sheets that aren't exactly full) keep their
  per-sheet semantics.
- If you had built a workaround that loops `onSheetIndex(...)` over an
  auto-split file and sums per-sheet answers yourself, remove it —
  you would now be double-counting: one plain `columnStats()` call is
  the whole table.
- `shards()` on an auto-split file now returns shards for every chain
  sheet (previously only the first). Workers using `rowsForShard()`
  need no changes — keys stay local to each shard's sheet, and the
  documented "skip `$rn === 1`" header guidance now simply applies per
  sheet.

### 2. Opt in to the config file actually working

Published `config/xlsx-stream.php` keys were never read before v3.2.2.
They now apply at writer construction — but only when the config
carries `'version' => 2`, so existing published copies stay inert and
nothing changes until you opt in.

**The safe way to opt in** is re-publishing the file, which ships the
new defaults (they match the code: compression level 5, flush interval
10,000, part size 8 MB, concurrency 4):

```bash
php artisan vendor:publish --tag=xlsx-stream-config --force
```

Then set the env vars you actually want (`XLSX_STREAM_COMPRESSION_LEVEL`,
`XLSX_STREAM_BUFFER_FLUSH_INTERVAL`, `XLSX_STREAM_S3_PART_SIZE`,
`XLSX_STREAM_S3_CONCURRENCY`). Code-level setters always override
config values.

> **Do NOT just add `'version' => 2` to your old published file.** The
> pre-3.2.2 file shipped stale defaults that contradict the code —
> `compression_level => 1` while the writer's real default has been 5
> since v3.1. Adding the marker to that file as-is would activate
> level 1 and silently change your output size/speed. Re-publish with
> `--force` (or review every value first, then add the marker).

Keys that had no implementation behind them are gone from the file:
`logging.*` (use the `onProgress()` callback), `memory.file_buffer_size`
and `max_rows_per_sheet` (never configurable). Transient S3 retries are
configured on your own `S3Client` (`'retries' => N`) — the AWS SDK's
retry middleware owns them; the sink adds one last-resort re-dispatch.

## Upgrading from v3.0.x to v3.1

No breaking API changes — every v3.0 call site works unchanged. Two
behavioral notes:

### 1. Default compression level is now 5 (was 6)

Output **bytes** change (the ZIP stays fully valid and Excel-compatible);
file size grows ~0.2 % while writes get ~20 % faster. If anything in your
pipeline compares file hashes or you want the exact old bytes back:

```php
$writer->setCompressionLevel(6);
```

### 2. Column stats ride the sidecar without breaking old readers

Files written **without** `withColumnStats()` keep emitting the KXSI
sidecar byte-identically. Files written **with** it append a TLV
"STAT" section after the core body while the version byte stays 2 —
the v3.0.x decoder parses the core and ignores trailing bytes, so
older readers keep full `rowAt`/`rowRange`/`rowCount` random access on
stats-bearing files; only the new `columnStats()`/`rowsWhere()`/
`findRow()` surface requires v3.1 on the reading side.

## Upgrading from v1.x to v2.0

v2.0 fixes data-corruption bugs and modernizes the dependency matrix. Below is
every behavior change a v1.x user might hit, with a concrete migration step.

### Composer constraints

| Requirement | v1.x | v2.0 |
|---|---|---|
| PHP | `^8.0` | `^8.1` |
| Laravel | 9 / 10 / 11 / 12 | 10 / 11 / 12 / 13 |
| `aws/aws-sdk-php` | `^3.0` | `^3.300` |

If your project is on PHP 8.0 or Laravel 9, **stay on v1.0.x** — both are
past upstream EOL.

```bash
composer require kolay/xlsx-stream:^2.0
```

### Behavior changes

#### 1. Boolean values now produce native Excel boolean cells

```php
$writer->writeRow([true, false]);
```

| | v1.x output | v2.0 output |
|---|---|---|
| `true` | string cell `"1"` | boolean cell TRUE |
| `false` | empty cell | boolean cell FALSE |

If a downstream parser depended on the v1.x string output, cast to int or
string yourself before passing:

```php
$writer->writeRow([(int) $isActive, $verified ? 'Y' : 'N']);
```

#### 2. `\DateTimeInterface` is now supported (was a fatal error)

```php
$writer->writeRow([new DateTime('2026-01-15 10:30')]);
```

In v1.x this threw `Error: Object of class DateTime could not be converted
to string`. In v2.0 it produces a real Excel date cell formatted as
`yyyy-mm-dd hh:mm:ss` and sortable as a date.

If you were previously formatting dates yourself and want the v1.x behavior:

```php
$writer->writeRow([$dt->format('Y-m-d H:i:s')]);
```

#### 3. Long numeric strings are preserved as text

```php
$writer->writeRow(['12345678901234567890']);
```

| | v1.x output | v2.0 output |
|---|---|---|
| 16+ digit numeric string | scientific notation `1.23E+19` (precision loss) | preserved as text |
| `+90555` | `90555` (sign dropped) | preserved as text |

This is a data-correctness fix. If you actually wanted float behavior, cast
explicitly:

```php
$writer->writeRow([(float) $value]);
```

#### 4. State guards now throw controlled exceptions

| Misuse | v1.x | v2.0 |
|---|---|---|
| `writeRow()` before `startFile()` | silently broken XLSX | `XlsxStreamException` |
| `startFile()` twice | silently broken XLSX | `XlsxStreamException` |
| `finishFile()` twice | fatal `TypeError` | `XlsxStreamException` |
| `writeRow()` after `finishFile()` | silently broken XLSX | `XlsxStreamException` |

#### 5. Excel's 16,384 column limit is enforced

Headers or rows wider than 16,384 cells now throw `XlsxStreamException`
on `startFile()` / `writeRow()`. v1.x silently produced unopenable XLSX.

### Operational changes

#### `S3MultipartSink::abort()` failures are now logged

If `abortMultipartUpload` itself fails (network blip, permissions revoked
mid-flight), v2.0 emits one line via `error_log()`:

```
[kolay/xlsx-stream] Failed to abort multipart upload — orphan upload may incur
S3 charges. bucket=… key=… uploadId=… error=…
```

We strongly recommend pairing this with an S3 lifecycle rule on
`AbortIncompleteMultipartUpload` so orphan parts are bounded automatically.

### Nothing-to-do changes (safe upgrades)

- Repo hygiene additions: `.gitattributes`, `phpstan.neon`, `pint.json`,
  GitHub Actions CI — none affect runtime.
- Minor cleanup of redundant constructor initialization in
  `SinkableXlsxWriter`. Same observable behavior.

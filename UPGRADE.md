# Upgrade Guide

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

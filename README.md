# Kolay XLSX Stream

[![Latest Version on Packagist](https://img.shields.io/packagist/v/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![Tests](https://img.shields.io/github/actions/workflow/status/turgutahmet/kolay-xlsx-stream/tests.yml?branch=main&label=tests&style=flat-square)](https://github.com/turgutahmet/kolay-xlsx-stream/actions/workflows/tests.yml)
[![Total Downloads](https://img.shields.io/packagist/dt/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![License](https://img.shields.io/packagist/l/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)
[![PHP Version](https://img.shields.io/packagist/php-v/kolay/xlsx-stream.svg?style=flat-square)](https://packagist.org/packages/kolay/xlsx-stream)

High-performance XLSX streaming writer for Laravel with **zero disk I/O** and **direct S3 support**. Perfect for exporting millions of rows without memory issues.

## Why This Package?

### The Problem with Existing Solutions

Most PHP Excel libraries (PHPSpreadsheet, Spout, Laravel Excel) have critical limitations:

- **Memory Issues**: Load entire documents in RAM (unusable for large files)
- **Disk I/O**: Write temporary files then upload to S3 (2x I/O, slow)
- **No True Streaming**: Can't stream directly to S3

### Our Solution

- **Zero Disk I/O**: Direct streaming to S3 using multipart upload
- **Constant Memory**: O(1) memory usage - only 32MB buffer regardless of file size
- **Blazing Fast**: 2-3x faster than alternatives
- **Production Tested**: Successfully exported 4.65 million rows (500MB+ files)

## Performance Comparison

### Latest Benchmark — v2.2 (May 2026)

Re-measured on an Apple Silicon laptop with PHP 8.2.28 and AWS SDK
3.379 against the same `xlsx-test-package` bucket in `us-east-2`. The
workload is identical to the v1.x baseline below (8 columns,
mixed types, compression level 1).

| Rows | Local Speed | Local Time | S3 Speed | S3 Memory | S3 Time | File Size |
|------|-------------|------------|----------|-----------|---------|-----------|
| 100 | 33,784 rows/s | 0.00s | 106 rows/s | 0 MB | 0.94s | 0.01 MB |
| 500 | 153,942 rows/s | 0.00s | 481 rows/s | 0 MB | 1.04s | 0.02 MB |
| 1,000 | 152,792 rows/s | 0.01s | 875 rows/s | 0 MB | 1.14s | 0.04 MB |
| 5,000 | 160,639 rows/s | 0.03s | 3,342 rows/s | 0 MB | 1.50s | 0.2 MB |
| 10,000 | 163,765 rows/s | 0.06s | 4,726 rows/s | 0 MB | 2.12s | 0.4 MB |
| 25,000 | 173,144 rows/s | 0.14s | 11,357 rows/s | 0 MB | 2.20s | 1 MB |
| 50,000 | 175,682 rows/s | 0.28s | 20,014 rows/s | 2 MB | 2.50s | 2 MB |
| 100,000 | 175,130 rows/s | 0.57s | 22,372 rows/s | 4 MB | 4.47s | 4 MB |
| 250,000 | 170,788 rows/s | 1.46s | 62,340 rows/s | 12 MB (±6) | 4.01s | 10 MB |
| 500,000 | 171,361 rows/s | 2.92s | 83,525 rows/s | 20 MB (±18) | 5.99s | 20 MB |
| 750,000 | 168,742 rows/s | 4.44s | 93,753 rows/s | 30 MB (±26) | 8.00s | 30 MB |
| 1,000,000 | 161,250 rows/s | 6.20s | 95,070 rows/s | 40 MB (±38) | 10.52s | 40 MB |
| 1,500,000 | 168,616 rows/s | 8.90s | 103,311 rows/s | 60 MB (±58) | 14.52s | 60 MB |
| 2,000,000 | 166,369 rows/s | 12.02s | 95,629 rows/s | 79 MB (±77) | 20.91s | 79 MB |
| 3,000,000 | – | – | 105,291 rows/s | 119 MB (±117) | 28.49s | 119 MB |
| 4,000,000 | – | – | 106,815 rows/s | 158 MB (±156) | 37.45s | 158 MB |
| 4,500,000 | – | – | 106,715 rows/s | 178 MB (±178) | 42.17s | 178 MB |

#### What changed since v1.x

- **S3 throughput is up roughly 2–3×** for any workload above 50K rows
  (1M: 95K rows/s vs 43K, 4.5M: 107K rows/s vs 46K). Most of the win
  comes from updated `aws/aws-sdk-php` (3.379+) and a faster network on
  the measurement machine — the multipart-upload code path itself is
  unchanged.
- **Local throughput is ~5–10% lower** than the v1.x numbers — the cost
  of the v2.0+ per-cell type detection (boolean cells, `DateTimeInterface`
  → serial date, big-integer-string preservation). It's a deliberate
  trade-off: v1.x produced silently broken cells for those types.
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

### Performance Highlights *(v2.2, May 2026)*

- **Local File System**: ~160,000–175,000 rows/second with true O(1) memory
- **S3 Streaming**: 90,000–110,000 rows/second above 1M rows (2–3× the v1.x baseline)
- **Memory Efficiency**: Local uses <2 MB, S3 averages 40 MB per million rows
- **Multi-sheet Support**: Automatic sheet creation at Excel's 1,048,576 row limit
- **Production Ready**: Successfully tested with 4.5 million rows

### Comparison with Other Libraries

| Package | 1M Rows Time | Memory Usage | Disk Usage | S3 Support |
|---------|--------------|--------------|------------|------------|
| PHPSpreadsheet | ❌ Crashes | ~8GB | Full file | Indirect |
| Spout | ~60 sec | ~100MB+ | Full file | Indirect |
| Laravel Excel | ~90 sec | ~500MB+ | Full file | Indirect |
| **Kolay XLSX Stream (Local)** | ✅ **6.2 sec** | ✅ **0 MB** | ✅ **Zero** | N/A |
| **Kolay XLSX Stream (S3)** | ✅ **10.5 sec** | ✅ **40MB avg** | ✅ **Zero** | ✅ **Direct** |

## Requirements

- PHP 8.1+
- Laravel 10, 11, 12 or 13
- AWS SDK (only if using S3 streaming)

> Upgrading from v1.x? See [UPGRADE.md](UPGRADE.md) for the v2.0 migration guide.

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

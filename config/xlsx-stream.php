<?php

/*
|--------------------------------------------------------------------------
| kolay/xlsx-stream configuration
|--------------------------------------------------------------------------
|
| Precedence: code-level setters > this config > package defaults.
| Values here are applied when a writer is constructed inside Laravel;
| any setter you call afterwards (setCompressionLevel(),
| setBufferFlushInterval(), an explicit $partSize argument) wins.
| Outside Laravel the file is simply never read.
|
| The `version` key is a compatibility gate, not a user setting. Copies
| of this file published before v3.2.2 listed keys the package never
| read (writer/s3/memory/logging) with stale values that contradicted
| the code defaults — e.g. compression_level 1 while the writer's real
| default is 5. Honouring those old copies retroactively would have
| silently changed existing applications' output, so the package only
| applies configs that declare `'version' => 2` (this file). If you
| still have a pre-3.2.2 copy published, re-publish it:
|
|     php artisan vendor:publish --tag=xlsx-stream-config --force
|
| Keys with no runtime counterpart (memory.*, logging.*,
| writer.max_rows_per_sheet) were removed in v3.2.2 and do not come
| back — the row-per-sheet ceiling is Excel's, not configurable.
|
*/

return [
    'version' => 2,

    'writer' => [
        // Deflate level 1-9. 5 is the measured knee of the size/speed
        // curve for XLSX-shaped XML (within ~0.2% of level 6's size at
        // ~20% less wall time).
        'compression_level' => env('XLSX_STREAM_COMPRESSION_LEVEL', 5),

        // Rows buffered before each flush to the sink.
        'buffer_flush_interval' => env('XLSX_STREAM_BUFFER_FLUSH_INTERVAL', 10000),
    ],

    's3' => [
        // Multipart part size in bytes for forDisk() writers. S3's
        // minimum for non-final parts is 5 MB; the sink silently
        // raises anything smaller.
        'part_size' => env('XLSX_STREAM_S3_PART_SIZE', 8 * 1024 * 1024),

        // Max part uploads in flight at once for forDisk() writers.
        // Memory ceiling is roughly part_size x (concurrency + 1);
        // 1 = strictly sequential uploads. Writers that need per-call
        // control construct S3MultipartSink directly — its constructor
        // takes both part size and concurrency.
        'concurrency' => env('XLSX_STREAM_S3_CONCURRENCY', 4),

        // Why there are no retry keys here: transient upload errors are
        // retried by the AWS SDK's own retry middleware — configure it
        // on YOUR S3Client ('retries' => N); after the SDK gives up the
        // sink re-dispatches the part once synchronously as a last
        // resort. A package-level retry_attempts/retry_delay_ms pair
        // (as pre-3.2.2 copies listed) would be a dead copy shadowing
        // the SDK setting, so those keys are intentionally absent.
        // Likewise the old logging.* keys: the runtime counterpart is
        // the $writer->onProgress(callable) callback, not a config key.
    ],
];

<?php

return [
    /*
    |--------------------------------------------------------------------------
    | Default S3 Settings
    |--------------------------------------------------------------------------
    |
    | These settings will be used for S3 operations if not specified
    |
    */
    's3' => [
        'part_size' => env('XLSX_STREAM_S3_PART_SIZE', 32 * 1024 * 1024), // 32MB default
        'retry_attempts' => env('XLSX_STREAM_S3_RETRY_ATTEMPTS', 3),
        'retry_delay_ms' => env('XLSX_STREAM_S3_RETRY_DELAY', 100),
    ],

    /*
    |--------------------------------------------------------------------------
    | Writer Performance Settings
    |--------------------------------------------------------------------------
    |
    | These settings control the performance characteristics of the writer
    |
    */
    'writer' => [
        // Compression level (1-9). 1 = fastest, 9 = best compression
        'compression_level' => env('XLSX_STREAM_COMPRESSION_LEVEL', 1),
        
        // Number of rows to buffer before flushing to output
        'buffer_flush_interval' => env('XLSX_STREAM_BUFFER_FLUSH_INTERVAL', 10000),
        
        // Maximum rows per sheet (Excel limit is 1,048,576)
        'max_rows_per_sheet' => env('XLSX_STREAM_MAX_ROWS_PER_SHEET', 1048575),
    ],

    /*
    |--------------------------------------------------------------------------
    | Memory Settings
    |--------------------------------------------------------------------------
    |
    | Control memory usage of the package
    |
    */
    'memory' => [
        // File write buffer size
        'file_buffer_size' => env('XLSX_STREAM_FILE_BUFFER_SIZE', 1024 * 1024), // 1MB
    ],

    /*
    |--------------------------------------------------------------------------
    | Logging
    |--------------------------------------------------------------------------
    |
    | Control what gets logged
    |
    */
    'logging' => [
        'enabled' => env('XLSX_STREAM_LOGGING', true),
        'channel' => env('XLSX_STREAM_LOG_CHANNEL', null), // null = default channel
        'log_progress' => env('XLSX_STREAM_LOG_PROGRESS', false), // Log progress updates
        'progress_interval' => env('XLSX_STREAM_PROGRESS_INTERVAL', 10000), // Log every N rows
    ],
];
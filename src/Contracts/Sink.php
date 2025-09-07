<?php

namespace Kolay\XlsxStream\Contracts;

/**
 * Sink interface for streaming output abstraction
 * 
 * Allows writing to different destinations (S3, file, memory)
 * without changing the writer logic
 */
interface Sink
{
    /**
     * Write data to the sink
     * 
     * @param string $data Binary data to write
     * @return void
     * @throws \RuntimeException on write failure
     */
    public function write(string $data): void;
    
    /**
     * Close the sink successfully
     * Finalize any pending operations
     * 
     * @return void
     * @throws \RuntimeException on finalization failure
     */
    public function close(): void;
    
    /**
     * Abort and cleanup
     * Called on errors to cleanup partial data
     * 
     * @return void
     */
    public function abort(): void;
    
    /**
     * Get total bytes written
     * 
     * @return int
     */
    public function getBytesWritten(): int;
}
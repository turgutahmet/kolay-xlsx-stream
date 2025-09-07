<?php

namespace Kolay\XlsxStream\Sinks;

use Kolay\XlsxStream\Contracts\Sink;

/**
 * File Sink - Local file system implementation
 * 
 * Simple wrapper around file operations
 * Provides same interface as S3MultipartSink for compatibility
 */
class FileSink implements Sink
{
    private $handle;
    private string $path;
    private int $bytesWritten = 0;
    private bool $closed = false;
    
    /**
     * @param string $path File path to write to
     * @param string $mode File open mode (default: 'wb')
     */
    public function __construct(string $path, string $mode = 'wb')
    {
        $this->path = $path;
        
        // Ensure directory exists
        $dir = dirname($path);
        if (!is_dir($dir)) {
            if (!mkdir($dir, 0755, true) && !is_dir($dir)) {
                throw new \RuntimeException("Failed to create directory: {$dir}");
            }
        }
        
        // Open file
        $this->handle = fopen($path, $mode);
        if ($this->handle === false) {
            throw new \RuntimeException("Failed to open file: {$path}");
        }
        
        // Set write buffer for better performance
        stream_set_write_buffer($this->handle, 1 << 20); // 1MB buffer
    }
    
    /**
     * Write data to file
     */
    public function write(string $data): void
    {
        if ($this->closed) {
            throw new \RuntimeException('Cannot write to closed sink');
        }
        
        $written = fwrite($this->handle, $data);
        
        if ($written === false) {
            throw new \RuntimeException("Failed to write to file: {$this->path}");
        }
        
        $this->bytesWritten += $written;
    }
    
    /**
     * Close the file
     */
    public function close(): void
    {
        if ($this->closed) {
            return;
        }
        
        if ($this->handle) {
            fflush($this->handle);
            fclose($this->handle);
        }
        
        $this->closed = true;
    }
    
    /**
     * Abort and cleanup
     */
    public function abort(): void
    {
        if ($this->closed) {
            return;
        }
        
        // Close file handle
        if ($this->handle) {
            fclose($this->handle);
        }
        
        // Delete partial file
        if (file_exists($this->path)) {
            unlink($this->path);
        }
        
        $this->closed = true;
    }
    
    /**
     * Get total bytes written
     */
    public function getBytesWritten(): int
    {
        return $this->bytesWritten;
    }
    
    /**
     * Get file path
     */
    public function getPath(): string
    {
        return $this->path;
    }
    
    /**
     * Destructor - ensure file is closed
     */
    public function __destruct()
    {
        if (!$this->closed && $this->handle) {
            fclose($this->handle);
        }
    }
}
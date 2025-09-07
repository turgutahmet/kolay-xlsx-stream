<?php

namespace Kolay\XlsxStream\Writers;

use Kolay\XlsxStream\Contracts\Sink;
use Kolay\XlsxStream\Sinks\FileSink;

/**
 * Sinkable XLSX Writer - S3 & Local File Compatible
 * 
 * Extends BaseXlsxWriter to use Sink abstraction
 * Enables direct S3 streaming without disk I/O
 */
class SinkableXlsxWriter extends BaseXlsxWriter
{
    private Sink $sink;
    
    /**
     * Create writer with Sink
     * 
     * @param Sink $sink Output sink (S3, file, etc)
     */
    public function __construct(Sink $sink)
    {
        $this->sink = $sink;
        
        // Initialize inherited properties
        $this->centralDirectory = [];
        $this->currentOffset = 0;
        $this->sheets = [];
        $this->currentSheetIndex = 0;
        $this->columns = [];
        $this->deflateCtx = null;
        $this->crcContext = null;
        $this->sheetCrc = 0;
        $this->sheetUncompressedSize = 0;
        $this->sheetCompressedSize = 0;
        $this->sheetOffset = 0;
        $this->currentSheetRow = 0;
        $this->totalRows = 0;
        $this->bufferFlushInterval = 1000;
        $this->rowBuffer = '';
        $this->rowBufferCount = 0;
        $this->deflateLevel = 6;
    }
    
    /**
     * Create from file path (backward compatibility)
     */
    public static function createForFile(string $path): self
    {
        return new self(new FileSink($path));
    }
    
    /**
     * Override all write operations to use sink
     */
    protected function writeToDest(string $data): void
    {
        $this->sink->write($data);
        $this->currentOffset += strlen($data);
    }
    
    /**
     * Finalize the XLSX file
     */
    public function finishFile(): array
    {
        try {
            // Close current sheet if open
            if ($this->currentSheetRow > 0) {
                $this->flushRowBuffer();
                $this->finishCurrentSheet();
            }
            
            // Write workbook files
            $this->writeStaticFile('xl/_rels/workbook.xml.rels', $this->getWorkbookRelsXml());
            $this->writeStaticFile('xl/workbook.xml', $this->getWorkbookXml());
            
            // Write Content_Types.xml LAST
            $this->writeStaticFile('[Content_Types].xml', $this->getContentTypesXml());
            
            // Write central directory
            $this->writeCentralDirectory();
            
            // Close sink successfully
            $this->sink->close();
            
            return [
                'bytes' => $this->currentOffset,
                'rows' => $this->totalRows,
                'sheets' => count($this->sheets),
                'sheet_details' => $this->sheets,
                'sink_bytes' => $this->sink->getBytesWritten()
            ];
            
        } catch (\Throwable $e) {
            // Abort sink on error
            $this->sink->abort();
            throw $e;
        }
    }
}
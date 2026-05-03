<?php

namespace Kolay\XlsxStream\Writers;

use Kolay\XlsxStream\Contracts\Sink;
use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;

/**
 * Sinkable XLSX Writer - S3 & Local File Compatible
 *
 * Extends BaseXlsxWriter to use Sink abstraction.
 * Enables direct S3 streaming without disk I/O.
 */
class SinkableXlsxWriter extends BaseXlsxWriter
{
    private Sink $sink;

    public function __construct(Sink $sink)
    {
        $this->sink = $sink;
    }

    /**
     * Create from file path (backward compatibility)
     */
    public static function createForFile(string $path): self
    {
        return new self(new FileSink($path));
    }

    protected function writeToDest(string $data): void
    {
        $this->sink->write($data);
        $this->currentOffset += strlen($data);
    }

    /**
     * Finalize the XLSX file. On any failure during finalization the sink is
     * aborted so partial output (e.g. orphan S3 multipart uploads) is cleaned up.
     */
    public function finishFile(): array
    {
        if ($this->closed) {
            throw XlsxStreamException::writerAlreadyClosed();
        }
        if (!$this->started) {
            throw XlsxStreamException::headersNotSet();
        }

        try {
            if ($this->currentSheetRow > 0) {
                $this->flushRowBuffer();
                $this->finishCurrentSheet();
            }

            $this->writeStaticFile('xl/_rels/workbook.xml.rels', $this->getWorkbookRelsXml());
            $this->writeStaticFile('xl/workbook.xml', $this->getWorkbookXml());
            $this->writeStaticFile('[Content_Types].xml', $this->getContentTypesXml());

            $this->writeCentralDirectory();

            $this->sink->close();
            $this->closed = true;

            return [
                'bytes' => $this->currentOffset,
                'rows' => $this->totalRows,
                'sheets' => count($this->sheets),
                'sheet_details' => $this->sheets,
                'sink_bytes' => $this->sink->getBytesWritten(),
            ];
        } catch (\Throwable $e) {
            $this->sink->abort();
            $this->closed = true;
            throw $e;
        }
    }
}
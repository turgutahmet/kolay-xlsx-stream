<?php

namespace Kolay\XlsxStream\Writers;

use Aws\S3\S3Client;
use Kolay\XlsxStream\Contracts\Sink;
use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sinks\S3MultipartSink;

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

    /**
     * Create a writer that streams to a Laravel filesystem disk.
     *
     * Supported drivers:
     *  - 's3' (or any disk with `driver: s3`) → S3MultipartSink with credentials
     *    pulled from `config('filesystems.disks.{$disk}')`
     *  - 'local' / 'public' (any local-style driver) → FileSink at the disk's
     *    resolved path
     *
     * Examples:
     *
     *   $writer = SinkableXlsxWriter::forDisk('s3', 'exports/report.xlsx');
     *   $writer = SinkableXlsxWriter::forDisk('local', 'exports/report.xlsx');
     *
     * Pass `$putObjectParams` (e.g. `['ACL' => 'public-read']`) for S3-only
     * options. They are ignored for local disks.
     *
     * @throws XlsxStreamException when the disk is not configured or the
     *                             driver is not supported.
     */
    public static function forDisk(
        string $disk,
        string $path,
        array $putObjectParams = [],
        int $partSize = S3MultipartSink::DEFAULT_PART_SIZE
    ): self {
        if (! function_exists('config')) {
            throw new XlsxStreamException(
                'forDisk() requires Laravel — config() helper is not available. '.
                'Use createForFile() or pass a Sink to the constructor instead.'
            );
        }

        $config = config("filesystems.disks.{$disk}");
        if (! is_array($config)) {
            throw new XlsxStreamException(
                "Disk [{$disk}] is not configured in filesystems.disks."
            );
        }

        $driver = $config['driver'] ?? null;

        if ($driver === 's3') {
            $client = new S3Client([
                'version' => 'latest',
                'region' => $config['region'] ?? 'us-east-1',
                'credentials' => [
                    'key' => $config['key'] ?? '',
                    'secret' => $config['secret'] ?? '',
                ],
                'endpoint' => $config['endpoint'] ?? null,
                'use_path_style_endpoint' => $config['use_path_style_endpoint'] ?? false,
            ]);

            $sink = new S3MultipartSink(
                $client,
                $config['bucket'] ?? '',
                $path,
                $partSize,
                $putObjectParams
            );

            return new self($sink);
        }

        if (in_array($driver, ['local', null], true)) {
            $root = $config['root'] ?? storage_path('app');
            $fullPath = rtrim($root, '/').'/'.ltrim($path, '/');

            return new self(new FileSink($fullPath));
        }

        throw new XlsxStreamException(
            "Disk [{$disk}] driver [{$driver}] is not supported. ".
            'Supported drivers: s3, local.'
        );
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

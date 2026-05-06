<?php

namespace Kolay\XlsxStream\Readers;

use Aws\S3\S3Client;
use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Sources\S3RangeSource;

/**
 * Public-facing streaming reader.
 *
 * Construction is via factory methods, never `new`:
 *
 *     StreamingXlsxReader::fromFile('/path/to/big.xlsx')
 *     StreamingXlsxReader::fromS3($s3Client, 'bucket', 'key.xlsx')
 *     StreamingXlsxReader::from(new LocalFileSource(...))
 *
 * Once constructed the reader exposes:
 *
 *   - sheets()                   list of {name, sheetId, entry}
 *   - onSheet('Reports')         select sheet by name
 *   - onSheetIndex(0)            select sheet by tab order
 *   - header()                   first row of the selected sheet
 *   - rows(skip, limit)          generator over data rows
 *   - chunked(N)                 generator over batches of N rows
 *   - rowCount()                 total rows in selected sheet (full scan)
 *
 * RAM is bounded — independent of sheet size. Each call to rows() /
 * chunked() opens a fresh forward-only inflate stream so the API is
 * replayable from the caller's perspective without memoising the
 * dataset in memory.
 *
 * Files written by SinkableXlsxWriter use inline strings exclusively
 * (t="inlineStr") — no sharedStrings.xml is needed. Files written by
 * other XLSX writers (PhpSpreadsheet, openpyxl, Apache POI, …) typically
 * deduplicate strings into xl/sharedStrings.xml; the reader detects this
 * and loads the table into memory transparently. Archives whose
 * sharedStrings.xml exceeds SST_RAM_THRESHOLD compressed bytes are
 * refused with a clear error — an on-disk variant for very large tables
 * is tracked as a future addition.
 */
class StreamingXlsxReader
{
    /**
     * Compressed-size threshold at which xl/sharedStrings.xml stops
     * fitting comfortably in RAM. ~99% of real-world XLSX files have a
     * sst well below this. Files above the threshold get a clear error
     * pointing at the limitation.
     */
    public const SST_RAM_THRESHOLD = 20 * 1024 * 1024;

    private Source $source;
    private ZipDirectory $cd;

    /** @var list<array{name: string, sheetId: int, entry: string}> */
    private array $sheets;

    private string $currentEntry;
    private ?array $cachedHeader = null;
    private ?SharedStrings $sst = null;
    private bool $sstResolved = false;
    private ?RandomAccessIndex $randomAccessIndex = null;
    private bool $indexResolved = false;

    private function __construct(Source $source)
    {
        $this->source = $source;
        $this->cd = ZipDirectory::fromSource($source);
        $this->sheets = WorkbookResolver::resolve($source, $this->cd);

        if ($this->sheets === []) {
            throw XlsxReadException::corruptCentralDirectory('workbook contains no sheets');
        }

        $this->currentEntry = $this->sheets[0]['entry'];
    }

    public static function from(Source $source): self
    {
        return new self($source);
    }

    public static function fromFile(string $path): self
    {
        return new self(new LocalFileSource($path));
    }

    public static function fromS3(S3Client $s3, string $bucket, string $key): self
    {
        return new self(new S3RangeSource($s3, $bucket, $key));
    }

    /**
     * Sheets in workbook.xml order (matches the user-visible tab order).
     *
     * @return list<array{name: string, sheetId: int, entry: string}>
     */
    public function sheets(): array
    {
        return $this->sheets;
    }

    public function onSheet(string $name): self
    {
        foreach ($this->sheets as $sheet) {
            if ($sheet['name'] === $name) {
                $this->currentEntry = $sheet['entry'];
                $this->cachedHeader = null;

                return $this;
            }
        }

        throw XlsxReadException::entryNotFound("sheet named '{$name}'");
    }

    public function onSheetIndex(int $index): self
    {
        if (! isset($this->sheets[$index])) {
            throw XlsxReadException::entryNotFound("sheet at index {$index}");
        }
        $this->currentEntry = $this->sheets[$index]['entry'];
        $this->cachedHeader = null;

        return $this;
    }

    /**
     * Return the first row of the selected sheet. Cached after first call,
     * so subsequent invocations do not re-open the underlying stream.
     *
     * @return array<int, mixed>
     */
    public function header(): array
    {
        if ($this->cachedHeader !== null) {
            return $this->cachedHeader;
        }

        foreach ($this->openSheetReader()->rows() as $row) {
            return $this->cachedHeader = $row;
        }

        return $this->cachedHeader = [];
    }

    /**
     * Yield rows from the selected sheet. By default returns every row
     * including the header; pass skip=1 to drop the header, or skip=N
     * to start at row N+1.
     *
     * @return \Generator<int, array<int, mixed>>
     */
    public function rows(int $skip = 0, ?int $limit = null): \Generator
    {
        if ($skip < 0) {
            $skip = 0;
        }
        if ($limit !== null && $limit <= 0) {
            return;
        }

        $emitted = 0;
        $skipped = 0;

        foreach ($this->openSheetReader()->rows() as $row) {
            if ($skipped < $skip) {
                $skipped++;

                continue;
            }
            yield $row;
            $emitted++;
            if ($limit !== null && $emitted >= $limit) {
                return;
            }
        }
    }

    /**
     * Yield rows in fixed-size batches. Last batch may be shorter than
     * $size if the sheet length is not divisible. Designed for bulk DB
     * inserts where the caller wants amortised round-trip cost.
     *
     * @return \Generator<int, list<array<int, mixed>>>
     */
    public function chunked(int $size, int $skip = 0): \Generator
    {
        if ($size < 1) {
            throw new \InvalidArgumentException('Chunk size must be at least 1');
        }

        $batch = [];
        foreach ($this->rows($skip) as $row) {
            $batch[] = $row;
            if (count($batch) >= $size) {
                yield $batch;
                $batch = [];
            }
        }
        if ($batch !== []) {
            yield $batch;
        }
    }

    /**
     * Total row count of the selected sheet. Currently O(N) — performs a
     * full inflate scan. A future Mode A* implementation will return O(1)
     * when an xl/_kxs/index.bin sidecar is present.
     */
    /**
     * Total row count including header. O(1) when the file carries a
     * matching xl/_kxs/index.bin sidecar (born-indexed); O(N) full
     * inflate scan otherwise. Both call sites are covered by the same
     * tests so the result is identical either way.
     */
    public function rowCount(): int
    {
        $index = $this->loadRandomAccessIndex();
        if ($index !== null) {
            $total = $index->totalRows($this->currentEntry);
            if ($total !== null) {
                return $total;
            }
        }

        return iterator_count($this->openSheetReader()->rows());
    }

    /**
     * Return a single row by 1-based row number — row 1 is the header,
     * row 2 is the first data row. Returns null when $rowNumber is past
     * the end of the sheet.
     *
     * Cost:
     *   - With xl/_kxs/index.bin → O(period) — fresh inflate from the
     *     nearest sync point + scan up to $period rows. Effectively
     *     constant time bounded by the writer-chosen sync period.
     *   - Without the sidecar → O(N) — full inflate scan from the
     *     start until the target row is reached.
     *
     * @return array<int, mixed>|null
     */
    public function rowAt(int $rowNumber): ?array
    {
        if ($rowNumber < 1) {
            return null;
        }

        [$compOffset, $startingRow] = $this->seekTarget($rowNumber);

        foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startingRow) as $rn => $row) {
            if ($rn === $rowNumber) {
                return $row;
            }
            if ($rn > $rowNumber) {
                return null;
            }
        }

        return null;
    }

    /**
     * Yield rows by 1-based inclusive range [from, to]. With an index
     * sidecar the cost is O(period + (to - from)); without, the cost
     * matches a row-skip O(from) plus the size of the range. Yielded
     * keys are the 1-based row numbers.
     *
     * @return \Generator<int, array<int, mixed>>
     */
    public function rowRange(int $from, int $to): \Generator
    {
        if ($from < 1) {
            $from = 1;
        }
        if ($to < $from) {
            return;
        }

        [$compOffset, $startingRow] = $this->seekTarget($from);

        foreach ($this->openSheetReader()->rowsFromOffset($compOffset, $startingRow) as $rn => $row) {
            if ($rn < $from) {
                continue;
            }
            if ($rn > $to) {
                return;
            }
            yield $rn => $row;
        }
    }

    /**
     * Resolve a target row to a (compressed-offset, starting-row)
     * pair the StreamingSheetReader can inflate from. Returns
     * [null, 1] when no index is present or the target precedes
     * every recorded sync point.
     *
     * @return array{0: int|null, 1: int}
     */
    private function seekTarget(int $rowNumber): array
    {
        $index = $this->loadRandomAccessIndex();
        if ($index === null) {
            return [null, 1];
        }
        $sp = $index->findSyncPoint($this->currentEntry, $rowNumber);
        if ($sp === null) {
            return [null, 1];
        }

        return [$sp['comp_offset'], $sp['row']];
    }

    public function close(): void
    {
        $this->source->close();
    }

    private function openSheetReader(): StreamingSheetReader
    {
        return new StreamingSheetReader(
            $this->source,
            $this->cd,
            $this->currentEntry,
            65536,
            $this->resolveSharedStrings(),
        );
    }

    /**
     * Lazy-load and cache the random-access index sidecar if present.
     * Files written without withRandomAccessIndex() carry no sidecar;
     * resolveRandomAccessIndex() returns null in that case and rowAt /
     * rowRange / rowCount fall back to sequential scans.
     */
    private function loadRandomAccessIndex(): ?RandomAccessIndex
    {
        if ($this->indexResolved) {
            return $this->randomAccessIndex;
        }
        $this->indexResolved = true;

        if (! $this->cd->has(RandomAccessIndex::ENTRY_PATH)) {
            return $this->randomAccessIndex = null;
        }

        $payload = $this->cd->readEntry($this->source, RandomAccessIndex::ENTRY_PATH);

        return $this->randomAccessIndex = RandomAccessIndex::decode($payload);
    }

    /**
     * Lazy-load the shared-strings table when first needed. Returns null
     * for archives that don't carry one (the kolay-xlsx-stream output
     * shape — every cell uses inlineStr). Throws a clear error for
     * tables larger than SST_RAM_THRESHOLD so callers know the file is
     * outside the supported range without their RAM filling up first.
     */
    private function resolveSharedStrings(): ?SharedStrings
    {
        if ($this->sstResolved) {
            return $this->sst;
        }

        $entry = $this->cd->entry('xl/sharedStrings.xml');
        if ($entry === null) {
            $this->sstResolved = true;

            return null;
        }

        if ($entry['compressed_size'] > self::SST_RAM_THRESHOLD) {
            $sizeMb = number_format($entry['compressed_size'] / 1024 / 1024, 1);
            throw XlsxReadException::corruptCentralDirectory(
                "xl/sharedStrings.xml is {$sizeMb} MB compressed — beyond the in-memory threshold ".
                'this reader supports. An on-disk variant for very large shared-strings tables '.
                'is tracked for a future release.'
            );
        }

        $sstXml = $this->cd->readEntry($this->source, 'xl/sharedStrings.xml');
        $this->sst = SharedStringsParser::parseInMemory($sstXml);
        $this->sstResolved = true;

        return $this->sst;
    }
}

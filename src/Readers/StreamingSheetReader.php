<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * Streams a single sheet's <row>...</row> sequence.
 *
 * Pipeline:
 *
 *     ZIP entry bytes ──► (inflate or pass-through) ──► inflated buffer
 *                                                      │
 *                                                      ▼
 *                                             drain complete <row>...</row>
 *                                                      │
 *                                                      ▼
 *                                         CellTokenizer::tokenizeRow → yield
 *
 * Compression methods supported:
 *   - DEFLATE (method 8) — typical XLSX worksheet path
 *   - STORED  (method 0) — raw uncompressed XML; some editors choose
 *     this for very small worksheets so they don't pay deflate setup
 *     cost on a few hundred bytes
 *
 * RAM is bounded: buffer holds at most one in-progress row plus a
 * read-chunk worth of bytes (64 KB compressed → up to ~256 KB inflated
 * with default zlib settings). The generator yields O(1) per row;
 * callers control accumulation.
 */
class StreamingSheetReader
{
    private const METHOD_STORED = 0;
    private const METHOD_DEFLATE = 8;

    /**
     * Hard ceiling on the in-progress row XML buffer. The reader holds
     * at most one open <row>...</row> element worth of bytes; if a
     * sheet never closes a row tag, the buffer would otherwise grow
     * without bound. 16 MB is far above any plausible legitimate row
     * (Excel's per-cell text limit is ~32 KB × 16384 columns) and far
     * below typical PHP memory budgets.
     */
    private const MAX_ROW_XML_BYTES = 16 * 1024 * 1024;

    private Source $source;
    private ZipDirectory $cd;
    private string $sheetEntry;
    private int $chunkSize;
    private ?SharedStrings $sst;

    public function __construct(
        Source $source,
        ZipDirectory $cd,
        string $sheetEntry = 'xl/worksheets/sheet1.xml',
        int $chunkSize = 65536,
        ?SharedStrings $sst = null,
    ) {
        $this->source = $source;
        $this->cd = $cd;
        $this->sheetEntry = $sheetEntry;
        $this->chunkSize = $chunkSize;
        $this->sst = $sst;
    }

    /**
     * Yield rows in document order as 0-indexed dense arrays. The first
     * yielded row is the header; callers wanting just data rows should
     * skip it (a public-facing reader will offer a header() helper).
     *
     * @return \Generator<int, array<int, mixed>>
     */
    public function rows(): \Generator
    {
        foreach ($this->rowsFromOffset(null, 1) as $row) {
            yield $row;
        }
    }

    /**
     * Lower-level row generator with explicit start position. Used by
     * the public reader for random access via xl/_kxs/index.bin sync
     * points: seek to the byte offset of the nearest sync point, init
     * a fresh inflate context (the marker is byte-aligned by
     * construction so no inflatePrime is needed), and resume row
     * decoding from $startingRowNumber.
     *
     * Generator key is the 1-based row number — row numbers match the
     * file's <row r="N"/> attribute, NOT the PHP zero-indexed array
     * convention. This gives callers a direct way to filter by row
     * range without re-deriving position.
     *
     * @return \Generator<int, array<int, mixed>>
     */
    public function rowsFromOffset(?int $compOffset, int $startingRowNumber): \Generator
    {
        $entry = $this->cd->entry($this->sheetEntry);
        if ($entry === null) {
            throw XlsxReadException::entryNotFound($this->sheetEntry);
        }

        $method = $entry['method'];
        if ($method !== self::METHOD_DEFLATE && $method !== self::METHOD_STORED) {
            throw XlsxReadException::corruptCentralDirectory(
                "worksheet '{$this->sheetEntry}' uses unsupported ZIP compression method {$method} ".
                '(only STORED and DEFLATE are supported)'
            );
        }

        $dataOffset = $this->cd->dataOffset($this->source, $this->sheetEntry);
        $startOffset = $dataOffset + ($compOffset ?? 0);
        $compRemaining = $entry['compressed_size'] - ($compOffset ?? 0);

        $stream = $this->source->streamFrom($startOffset);

        $inflate = null;
        if ($method === self::METHOD_DEFLATE) {
            $inflate = inflate_init(ZLIB_ENCODING_RAW);
            if ($inflate === false) {
                if (is_resource($stream)) {
                    fclose($stream);
                }
                throw XlsxReadException::inflateFailed('inflate_init returned false');
            }
        }

        $rowNumber = $startingRowNumber;
        $buffer = '';
        // STORED has no deflate state to flush — treat the FINISH step
        // as already done so the loop exits cleanly when input runs out.
        $finishedFlush = $method === self::METHOD_STORED;
        $emptyReads = 0;

        try {
            while (true) {
                $inflated = '';

                if ($compRemaining > 0) {
                    $compressed = fread($stream, min($this->chunkSize, $compRemaining));
                    $n = is_string($compressed) ? strlen($compressed) : 0;

                    if ($n === 0) {
                        if (feof($stream)) {
                            $compRemaining = 0;
                        } elseif (++$emptyReads >= 100) {
                            throw XlsxReadException::sourceUnreadable(
                                'source stream stalled — 100 consecutive empty reads with feof=false'
                            );
                        } else {
                            usleep(10_000); // 10ms backoff before retry
                            continue;
                        }
                    } else {
                        $emptyReads = 0;
                        $compRemaining -= $n;

                        if ($method === self::METHOD_DEFLATE) {
                            $flag = $compRemaining === 0 ? ZLIB_FINISH : ZLIB_NO_FLUSH;
                            if ($flag === ZLIB_FINISH) {
                                $finishedFlush = true;
                            }
                            $inflated = inflate_add($inflate, $compressed, $flag);
                            if ($inflated === false) {
                                throw XlsxReadException::inflateFailed('mid-stream inflate_add returned false');
                            }
                        } else {
                            // STORED: raw bytes are the "inflated" output.
                            $inflated = $compressed;
                        }
                    }
                } elseif (! $finishedFlush) {
                    $inflated = inflate_add($inflate, '', ZLIB_FINISH);
                    if ($inflated === false) {
                        throw XlsxReadException::inflateFailed('final inflate_add returned false');
                    }
                    $finishedFlush = true;
                }

                if ($inflated !== '') {
                    $buffer .= $inflated;
                }

                $cursor = 0;
                while (true) {
                    $rowStart = self::findRowOpen($buffer, $cursor);
                    if ($rowStart < 0) {
                        break;
                    }
                    $rowEnd = strpos($buffer, '</row>', $rowStart);
                    if ($rowEnd === false) {
                        break;
                    }
                    $rowXml = substr($buffer, $rowStart, $rowEnd + 6 - $rowStart);
                    $cursor = $rowEnd + 6;
                    yield $rowNumber => CellTokenizer::tokenizeRow($rowXml, $this->sst);
                    $rowNumber++;
                }

                if ($cursor > 0) {
                    $buffer = substr($buffer, $cursor);
                }

                // After draining every complete row, anything left in
                // $buffer is one in-progress row's prefix. A pathological
                // sheet that opens <row> and never closes it would let
                // the buffer grow to GB scale and OOM the process. Cap
                // it loudly instead.
                if (strlen($buffer) > self::MAX_ROW_XML_BYTES) {
                    throw XlsxReadException::corruptCentralDirectory(
                        'in-progress row XML exceeds '.(self::MAX_ROW_XML_BYTES / 1024 / 1024).
                        ' MB without a closing tag — sheet is malformed or malicious'
                    );
                }

                if ($compRemaining === 0 && $finishedFlush) {
                    break;
                }
            }
        } finally {
            if (is_resource($stream)) {
                fclose($stream);
            }
        }
    }

    /**
     * Locate the next "<row " or "<row>" at or after $start. Avoids matching
     * "<row" prefixes that belong to other element names.
     */
    private static function findRowOpen(string $s, int $start): int
    {
        $pos = $start;
        $len = strlen($s);

        while ($pos < $len) {
            $found = strpos($s, '<row', $pos);
            if ($found === false) {
                return -1;
            }
            $next = $s[$found + 4] ?? '';
            if ($next === ' ' || $next === '>' || $next === "\t" || $next === "\n" || $next === '/') {
                return $found;
            }
            $pos = $found + 4;
        }

        return -1;
    }
}

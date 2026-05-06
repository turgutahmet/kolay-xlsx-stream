<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * Streams a single sheet's <row>...</row> sequence.
 *
 * Pipeline (§3.3 / §3.5 of the spec):
 *
 *     compressed bytes ──► inflate_add (chunked) ──► inflated buffer
 *                                                   │
 *                                                   ▼
 *                                          drain complete <row>...</row>
 *                                                   │
 *                                                   ▼
 *                                      CellTokenizer::tokenizeRow → yield
 *
 * RAM is bounded: buffer holds at most one in-progress row plus a
 * compressed-chunk worth of inflated output (256 KB tunable). The
 * generator yields O(1) per row; callers control accumulation.
 */
class StreamingSheetReader
{
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
        $entry = $this->cd->entry($this->sheetEntry);
        if ($entry === null) {
            throw XlsxReadException::entryNotFound($this->sheetEntry);
        }

        $dataOffset = $this->cd->dataOffset($this->source, $this->sheetEntry);
        $compRemaining = $entry['compressed_size'];

        $stream = $this->source->streamFrom($dataOffset);
        $inflate = inflate_init(ZLIB_ENCODING_RAW);
        if ($inflate === false) {
            if (is_resource($stream)) {
                fclose($stream);
            }
            throw XlsxReadException::inflateFailed('inflate_init returned false');
        }

        $buffer = '';
        $finishedFlush = false;

        try {
            while (true) {
                $inflated = '';

                if ($compRemaining > 0) {
                    $compressed = fread($stream, min($this->chunkSize, $compRemaining));
                    $n = is_string($compressed) ? strlen($compressed) : 0;

                    if ($n === 0) {
                        if (feof($stream)) {
                            $compRemaining = 0;
                        } else {
                            // Source is open but yielded no bytes; loop again.
                            continue;
                        }
                    } else {
                        $compRemaining -= $n;
                        $flag = $compRemaining === 0 ? ZLIB_FINISH : ZLIB_NO_FLUSH;
                        if ($flag === ZLIB_FINISH) {
                            $finishedFlush = true;
                        }
                        $inflated = inflate_add($inflate, $compressed, $flag);
                        if ($inflated === false) {
                            throw XlsxReadException::inflateFailed('mid-stream inflate_add returned false');
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
                    yield CellTokenizer::tokenizeRow($rowXml, $this->sst);
                }

                if ($cursor > 0) {
                    $buffer = substr($buffer, $cursor);
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
            if ($next === ' ' || $next === '>' || $next === "\t" || $next === "\n" || $next === "/") {
                return $found;
            }
            $pos = $found + 4;
        }

        return -1;
    }
}

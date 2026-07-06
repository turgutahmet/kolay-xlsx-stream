<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Contracts\SupportsBoundedStream;
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
    private ?DateDetection $dates;

    public function __construct(
        Source $source,
        ZipDirectory $cd,
        string $sheetEntry = 'xl/worksheets/sheet1.xml',
        int $chunkSize = 65536,
        ?SharedStrings $sst = null,
        ?DateDetection $dates = null,
    ) {
        $this->source = $source;
        $this->cd = $cd;
        $this->sheetEntry = $sheetEntry;
        $this->chunkSize = $chunkSize;
        $this->sst = $sst;
        $this->dates = $dates;
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
     * $fastForwardTo skips every row BEFORE that row number without
     * tokenizing it: the inflated stream is consumed in boundary-count
     * mode — the same '</row>' counting countRows() uses, at the same
     * ~35x speedup over tokenization — until the target's preceding
     * boundary passes, then row assembly resumes from the byte after
     * it. The first yield is row $fastForwardTo (or nothing when the
     * sheet ends first). Correctness rides on the counting equivalence
     * documented on countRows(): every yield of this generator consumes
     * exactly one '</row>', self-closing empty rows included, so "skip
     * K yields" and "skip K boundaries" are the same operation. Values
     * at or below $startingRowNumber are a no-op — callers can pass
     * their target row unconditionally.
     *
     * @return \Generator<int, array<int, mixed>>
     */
    public function rowsFromOffset(?int $compOffset, int $startingRowNumber, ?int $fastForwardTo = null, ?int $compLength = null): \Generator
    {
        $rowNumber = $startingRowNumber;
        $buffer = '';

        // Boundaries (= yields, see above) still to skip before the
        // tokenizer takes over. The 5-byte carry mirrors countRows():
        // the needle is 6 bytes, so a carry alone can never complete a
        // match (no double counting) while a boundary straddling two
        // chunks is fully visible in carry+chunk.
        $skipRemaining = $fastForwardTo !== null ? max(0, $fastForwardTo - $startingRowNumber) : 0;
        $carry = '';

        foreach ($this->inflatedChunks($compOffset, $compLength) as $inflated) {
            if ($skipRemaining > 0) {
                $scan = $carry.$inflated;
                $found = substr_count($scan, '</row>');
                if ($found < $skipRemaining) {
                    $skipRemaining -= $found;
                    $carry = substr($scan, -5);

                    continue;
                }

                // The last boundary to skip closes inside this chunk.
                // Locate it and hand everything after it to the row
                // assembler: the bytes between that '</row>' and the
                // next row opening are inter-row filler findRowOpen()
                // skips naturally, so no re-alignment is needed. The
                // carry bytes double as buffer prefix here — positions
                // in $scan already account for them, and a match seen
                // by substr_count can start in the carry, so the locate
                // walk must run over the same string.
                $pos = 0;
                for ($i = 0; $i < $skipRemaining; $i++) {
                    $next = strpos($scan, '</row>', $pos);
                    if ($next === false) {
                        break; // unreachable — substr_count counted >= $skipRemaining matches
                    }
                    $pos = $next + 6;
                }
                $buffer = substr($scan, $pos);
                // Total boundaries skipped across all chunks is exactly
                // ($fastForwardTo - $startingRowNumber), one per yield
                // the tokenizer never saw.
                $rowNumber = $fastForwardTo;
                $skipRemaining = 0;
            } else {
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
                yield $rowNumber => CellTokenizer::tokenizeRow($rowXml, $this->sst, $this->dates);
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
        }
    }

    /**
     * Count the rows that rows() would yield — without tokenizing a
     * single cell.
     *
     * rows() consumes the sheet as "row-open … next '</row>'" slices:
     * every yield consumes exactly one '</row>', and every '</row>' in
     * well-formed sheetData belongs to exactly one yield. Counting
     * '</row>' occurrences in the inflated byte stream therefore gives
     * the same number as iterator_count(rows()) while skipping the
     * per-cell tokenization entirely.
     *
     * Why '</row>' and not row openings: self-closing rows
     * (<row r="5"/>, legal in external files for empty rows) produce
     * no yield of their own in rows() — the slice runs from their
     * opening to the *next* row's closing tag — and they emit no
     * '</row>' either, so the counts still agree. Counting openings
     * would overcount relative to rows() for exactly that shape.
     *
     * '</row>' cannot occur inside cell text ('<' is always escaped as
     * &lt; in XML content) and no other worksheet element matches the
     * full 6-byte pattern ('</rowBreaks>' fails the trailing '>').
     *
     * Chunk boundaries: the last 5 bytes of the previous chunk are
     * prepended before counting. The needle is 6 bytes, so the carry
     * alone can never hold a complete match (no double counting), while
     * any match straddling the boundary is fully visible in carry+chunk.
     */
    public function countRows(): int
    {
        $count = 0;
        $carry = '';

        foreach ($this->inflatedChunks(null) as $chunk) {
            $buf = $carry.$chunk;
            $count += substr_count($buf, '</row>');
            $carry = substr($buf, -5);
        }

        return $count;
    }

    /**
     * Integrity check: inflate the whole sheet once, running a CRC32 over
     * the uncompressed bytes, and compare it against the sidecar's
     * per-sync-point running-CRC pins (SCRC) and the whole-sheet CRC. This
     * is the read-side counterpart of the CRC the writer pinned at each
     * ZLIB_FULL_FLUSH boundary, so a corrupt block is localized to the two
     * sync points that bracket it. O(1) memory (streaming inflate).
     *
     * A checkpoint is `[uncompOffset, expectedCrc]` — expectedCrc must equal
     * the CRC32 of the first uncompOffset uncompressed bytes. $checkpoints
     * must be ascending by offset. `inflate_ok=false` means the compressed
     * data itself was too corrupt to inflate (a hard failure, reported not
     * thrown so callers get a full report).
     *
     * @param  list<array{0: int, 1: int}>  $checkpoints
     * @return array{sheet_crc_ok: bool, corrupt_blocks: list<int>, inflate_ok: bool}
     */
    public function verifyCrc(array $checkpoints, int $sheetCrc): array
    {
        $ctx = hash_init('crc32b');
        $consumed = 0;
        $ci = 0;
        $n = count($checkpoints);
        $corrupt = [];

        try {
            foreach ($this->inflatedChunks(null) as $chunk) {
                $len = strlen($chunk);
                $pos = 0;

                // Close out every checkpoint whose boundary falls in this chunk.
                while ($ci < $n && $checkpoints[$ci][0] <= $consumed + $len) {
                    $upto = $checkpoints[$ci][0] - $consumed;
                    if ($upto > $pos) {
                        hash_update($ctx, substr($chunk, $pos, $upto - $pos));
                        $pos = $upto;
                    }
                    if (hexdec(hash_final(hash_copy($ctx))) !== $checkpoints[$ci][1]) {
                        $corrupt[] = $ci;
                    }
                    $ci++;
                }

                if ($pos < $len) {
                    hash_update($ctx, substr($chunk, $pos));
                }
                $consumed += $len;
            }
        } catch (\Throwable) {
            return ['sheet_crc_ok' => false, 'corrupt_blocks' => $corrupt, 'inflate_ok' => false];
        }

        return [
            'sheet_crc_ok' => hexdec(hash_final($ctx)) === $sheetCrc,
            'corrupt_blocks' => $corrupt,
            'inflate_ok' => true,
        ];
    }

    /**
     * Stream the sheet entry's inflated bytes as string chunks. Shared
     * engine behind rowsFromOffset() (row assembly) and countRows()
     * (boundary counting): entry validation, the ranged read loop, the
     * inflate context and the stalled-source backoff all live here so
     * both consumers stay in lockstep.
     *
     * @return \Generator<int, string>
     */
    private function inflatedChunks(?int $compOffset, ?int $compLength = null): \Generator
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
        $entryRemaining = $entry['compressed_size'] - ($compOffset ?? 0);

        // A caller-supplied $compLength stops the read at a ZLIB_FULL_FLUSH
        // sync boundary (the run's trailing block). Because a full flush
        // already emits every byte before it, the bounded bytes are fed to
        // inflate with NO_FLUSH and the loop stops WITHOUT a ZLIB_FINISH
        // step — the exact inverse of seeking INTO a sync point, which the
        // random-access index already does. $bounded === false is the
        // historical path (read to entry end, then FINISH), preserved
        // byte-for-byte when $compLength is null or reaches the entry end.
        $compRemaining = $compLength !== null ? min($compLength, $entryRemaining) : $entryRemaining;
        $bounded = $compRemaining < $entryRemaining;

        // A source that can serve a bounded range (S3 range GET) fetches
        // only the run's bytes off the wire; others fall back to a plain
        // stream-to-EOF and the loop caps the read via $compRemaining. The
        // NO_FLUSH / skip-FINISH handling below is identical either way.
        $stream = ($bounded && $this->source instanceof SupportsBoundedStream)
            ? $this->source->streamFromRange($startOffset, $compRemaining)
            : $this->source->streamFrom($startOffset);

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
                            // Bounded reads end at a full-flush boundary, so
                            // the last chunk is still NO_FLUSH — never FINISH
                            // (the flush already emitted its output; a FINISH
                            // on a stream truncated before its BFINAL block
                            // would error).
                            $flag = ($compRemaining === 0 && ! $bounded) ? ZLIB_FINISH : ZLIB_NO_FLUSH;
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
                } elseif (! $finishedFlush && ! $bounded) {
                    $inflated = inflate_add($inflate, '', ZLIB_FINISH);
                    if ($inflated === false) {
                        throw XlsxReadException::inflateFailed('final inflate_add returned false');
                    }
                    $finishedFlush = true;
                }

                if ($inflated !== '') {
                    yield $inflated;
                }

                // Bounded runs stop as soon as their bytes are consumed
                // (no FINISH to wait for); unbounded runs wait for FINISH.
                if ($compRemaining === 0 && ($finishedFlush || $bounded)) {
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

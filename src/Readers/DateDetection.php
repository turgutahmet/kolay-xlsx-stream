<?php

namespace Kolay\XlsxStream\Readers;

/**
 * @internal
 *
 * Plumbing bundle for the reader's opt-in autoDetectDates(): everything
 * CellTokenizer needs to turn a date-styled numeric cell into a
 * DateTimeImmutable, carried as ONE optional parameter so the tokenizer
 * signature (and its default fast path) stays flat.
 *
 * Built fresh by StreamingXlsxReader::openSheetReader() each time an
 * iteration starts, so the skip set reflects the casts registered at
 * that moment.
 */
final class DateDetection
{
    /**
     * @param  list<bool>  $isDateByStyle  cellXfs index → renders as a date
     *                                     (StylesParser::dateStyleBitmap)
     * @param  \Closure  $convert  raw <v> string → DateTimeImmutable (or the
     *                             raw value back when it isn't a convertible
     *                             serial — auto-detection must never destroy
     *                             data it merely inferred about)
     * @param  array<int, true>  $skipColumns  0-based columns with an explicit
     *                                         castColumn() registered — the
     *                                         explicit cast wins, so detection
     *                                         leaves the raw serial for it
     */
    public function __construct(
        public readonly array $isDateByStyle,
        public readonly \Closure $convert,
        public readonly array $skipColumns = [],
    ) {
    }
}

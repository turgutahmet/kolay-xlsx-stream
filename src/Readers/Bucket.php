<?php

namespace Kolay\XlsxStream\Readers;

/**
 * Ready-made bucket callables for StreamingXlsxReader::groupStats() over
 * a DATE column. groupStats groups by the raw stored value, which for a
 * date cell is the Excel serial (days since 1899-12-30). These helpers
 * turn a serial into a calendar-period key so "GROUP BY month" is one
 * call instead of hand-rolling the serial → date arithmetic.
 *
 * The conversion mirrors the reader's serial convention exactly —
 * unix = (serial − 25569) × 86400, where 25569 is the serial of the Unix
 * epoch (1970-01-01) — and is evaluated in UTC so it never drifts with
 * the host timezone. Every bucket here is MONOTONE (a later serial maps
 * to an equal-or-greater key), which is what lets groupStats fold a
 * group-pure block straight from the sidecar without reading its rows.
 *
 * Example:
 *   $reader->groupStats('order_date', 'total', Bucket::month());
 *   // => one row per YYYYMM with sum/count/min/max of total
 */
final class Bucket
{
    /** Serial of the Unix epoch (1970-01-01) in the 1899-12-30 system. */
    private const UNIX_EPOCH_SERIAL = 25569;

    /** Group by calendar year: serial => YYYY (e.g. 2026). */
    public static function year(): \Closure
    {
        return static fn (float $serial): int => (int) gmdate('Y', self::toUnix($serial));
    }

    /** Group by calendar month: serial => YYYYMM (e.g. 202603). */
    public static function month(): \Closure
    {
        return static fn (float $serial): int => (int) gmdate('Ym', self::toUnix($serial));
    }

    /** Group by calendar day: serial => YYYYMMDD (e.g. 20260315). */
    public static function day(): \Closure
    {
        return static fn (float $serial): int => (int) gmdate('Ymd', self::toUnix($serial));
    }

    private static function toUnix(float $serial): int
    {
        return (int) (($serial - self::UNIX_EPOCH_SERIAL) * 86400);
    }
}

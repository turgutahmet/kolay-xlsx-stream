<?php

namespace Kolay\XlsxStream\Contracts;

/**
 * Optional capability: a Source that can describe its I/O cost so the
 * reader can plan gap-bridging (G5). When a pruned scan leaves two
 * surviving block runs separated by a short gap of non-matching blocks,
 * it is often cheaper on a high-latency source to keep ONE ranged read
 * alive straight through the gap than to pay a second round-trip. The
 * reader bridges a gap when transferring its compressed bytes costs less
 * than one round-trip: `gap_bytes / bandwidth_bps < rtt_us / 1e6`.
 *
 * Sources that do NOT implement this are treated as zero-latency (local
 * disk): no gap is ever bridged and run merging stays byte-for-byte the
 * contiguous-only behaviour. The numbers are deployment facts the Source
 * knows (region round-trip, measured throughput) — never baked constants.
 */
interface ProvidesCostHints
{
    /**
     * @return array{rtt_us: int, bandwidth_bps: int} round-trip latency in
     *                                                 microseconds and
     *                                                 throughput in bytes/sec
     */
    public function costHints(): array;
}

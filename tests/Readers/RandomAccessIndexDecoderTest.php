<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\RandomAccessIndex as WriterIndex;

/**
 * Unit tests for the reader's binary index decoder.
 *
 * Round-trips bytes produced by the writer-side encoder so any drift
 * between encoder and decoder fails immediately. Also exercises the
 * defensive failure paths — magic mismatch, version mismatch, body
 * truncation, CRC corruption — so a tampered or partial sidecar is
 * surfaced loudly instead of silently degrading random-access answers.
 */
class RandomAccessIndexDecoderTest extends TestCase
{
    public function test_decode_preserves_sync_period_and_per_sheet_data(): void
    {
        $points = [
            ['row' => 102, 'comp_offset' => 4_500, 'uncomp_offset' => 12_000],
            ['row' => 1_002, 'comp_offset' => 45_000, 'uncomp_offset' => 120_000],
        ];

        $payload = WriterIndex::encode(100, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 1_001,
            'sheet_crc32' => 0xDEADBEEF,
            'sync_points' => $points,
        ]]);

        $idx = ReaderIndex::decode($payload);

        $this->assertSame(100, $idx->syncPeriod());
        $this->assertSame(1_001, $idx->totalRows('xl/worksheets/sheet1.xml'));
        $this->assertNull($idx->totalRows('xl/worksheets/sheet99.xml'));
        $this->assertSame($points, $idx->syncPoints('xl/worksheets/sheet1.xml'));
    }

    public function test_find_sync_point_returns_largest_below_or_equal_target(): void
    {
        $payload = WriterIndex::encode(100, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 1_002,
            'sheet_crc32' => 0,
            'sync_points' => [
                ['row' => 102, 'comp_offset' => 1_000, 'uncomp_offset' => 10_000],
                ['row' => 502, 'comp_offset' => 5_000, 'uncomp_offset' => 50_000],
                ['row' => 902, 'comp_offset' => 9_000, 'uncomp_offset' => 90_000],
            ],
        ]]);

        $idx = ReaderIndex::decode($payload);
        $entry = 'xl/worksheets/sheet1.xml';

        // Below first sync point → null (caller streams from beginning).
        $this->assertNull($idx->findSyncPoint($entry, 50));
        $this->assertNull($idx->findSyncPoint($entry, 101));

        // Exact sync-point match returns that point.
        $this->assertSame(102, $idx->findSyncPoint($entry, 102)['row']);
        $this->assertSame(502, $idx->findSyncPoint($entry, 502)['row']);

        // Between sync points → previous one.
        $this->assertSame(102, $idx->findSyncPoint($entry, 250)['row']);
        $this->assertSame(502, $idx->findSyncPoint($entry, 800)['row']);

        // Beyond last sync point → still last sync point (caller scans
        // forward through the tail).
        $this->assertSame(902, $idx->findSyncPoint($entry, 999)['row']);
        $this->assertSame(902, $idx->findSyncPoint($entry, 100_000)['row']);
    }

    public function test_multi_sheet_decoded_independently(): void
    {
        $payload = WriterIndex::encode(50, [
            [
                'entry' => 'xl/worksheets/sheet1.xml',
                'total_rows' => 100,
                'sheet_crc32' => 0,
                'sync_points' => [['row' => 52, 'comp_offset' => 5, 'uncomp_offset' => 50]],
            ],
            [
                'entry' => 'xl/worksheets/sheet2.xml',
                'total_rows' => 200,
                'sheet_crc32' => 0,
                'sync_points' => [
                    ['row' => 52, 'comp_offset' => 6, 'uncomp_offset' => 60],
                    ['row' => 102, 'comp_offset' => 12, 'uncomp_offset' => 120],
                ],
            ],
        ]);

        $idx = ReaderIndex::decode($payload);

        $this->assertSame(100, $idx->totalRows('xl/worksheets/sheet1.xml'));
        $this->assertSame(200, $idx->totalRows('xl/worksheets/sheet2.xml'));
        $this->assertCount(1, $idx->syncPoints('xl/worksheets/sheet1.xml'));
        $this->assertCount(2, $idx->syncPoints('xl/worksheets/sheet2.xml'));
    }

    public function test_zero_sync_points_for_short_sheets(): void
    {
        $payload = WriterIndex::encode(10_000, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 50, // shorter than sync period — no points emitted
            'sheet_crc32' => 0,
            'sync_points' => [],
        ]]);

        $idx = ReaderIndex::decode($payload);

        $this->assertSame(50, $idx->totalRows('xl/worksheets/sheet1.xml'));
        $this->assertSame([], $idx->syncPoints('xl/worksheets/sheet1.xml'));
        $this->assertNull($idx->findSyncPoint('xl/worksheets/sheet1.xml', 25));
    }

    public function test_rejects_bad_magic(): void
    {
        $payload = "ABCD".str_repeat("\0", 12);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/magic.*KXSI/i');
        ReaderIndex::decode($payload);
    }

    public function test_rejects_unsupported_version(): void
    {
        $body = '';
        $header = WriterIndex::MAGIC.pack('CCv', 99, 0, 0).pack('V', 0).pack('V', crc32($body));

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/version 99/i');
        ReaderIndex::decode($header);
    }

    public function test_rejects_short_payload(): void
    {
        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/too short/i');
        ReaderIndex::decode("KXSI");
    }

    public function test_rejects_corrupted_body_via_crc_mismatch(): void
    {
        $payload = WriterIndex::encode(100, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 10,
            'sheet_crc32' => 0,
            'sync_points' => [],
        ]]);

        // Flip a single byte inside the body — CRC32 instantly invalidates.
        $tampered = substr($payload, 0, 16).chr(ord($payload[16]) ^ 0xFF).substr($payload, 17);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/CRC32 mismatch/i');
        ReaderIndex::decode($tampered);
    }

    public function test_rejects_zero_sync_period(): void
    {
        // syncPeriod is the cadence at which the writer planted sync
        // points. A zero value is structurally meaningless — refuse the
        // sidecar before any seek is attempted.
        $payload = WriterIndex::encode(0, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 10,
            'sheet_crc32' => 0,
            'sync_points' => [],
        ]]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/syncPeriod must be greater than zero/');
        ReaderIndex::decode($payload);
    }

    public function test_rejects_non_monotonic_sync_point_rows(): void
    {
        // Sync points must be strictly increasing on row — findSyncPoint
        // does a forward walk and assumes monotonicity. A descending or
        // duplicate row would break that contract silently.
        $payload = WriterIndex::encode(100, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 1000,
            'sheet_crc32' => 0,
            'sync_points' => [
                ['row' => 100, 'comp_offset' => 1000, 'uncomp_offset' => 5000],
                ['row' => 50, 'comp_offset' => 2000, 'uncomp_offset' => 10000],
            ],
        ]]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/not strictly increasing on row/');
        ReaderIndex::decode($payload);
    }

    public function test_rejects_non_monotonic_comp_offset(): void
    {
        $payload = WriterIndex::encode(100, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 1000,
            'sheet_crc32' => 0,
            'sync_points' => [
                ['row' => 100, 'comp_offset' => 5000, 'uncomp_offset' => 5000],
                ['row' => 200, 'comp_offset' => 4000, 'uncomp_offset' => 10000],
            ],
        ]]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/not strictly increasing on comp_offset/');
        ReaderIndex::decode($payload);
    }

    public function test_rejects_sync_row_far_beyond_total_rows(): void
    {
        // A sync point past totalRows + 1 (the legitimate EOF marker)
        // is structurally impossible — the corresponding rows do not
        // exist in the sheet.
        $payload = WriterIndex::encode(100, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 100,
            'sheet_crc32' => 0,
            'sync_points' => [
                ['row' => 5000, 'comp_offset' => 1000, 'uncomp_offset' => 5000],
            ],
        ]]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/exceeds totalRows/');
        ReaderIndex::decode($payload);
    }

    public function test_rejects_duplicate_sheet_path(): void
    {
        $payload = WriterIndex::encode(100, [
            [
                'entry' => 'xl/worksheets/sheet1.xml',
                'total_rows' => 100,
                'sheet_crc32' => 0,
                'sync_points' => [],
            ],
            [
                'entry' => 'xl/worksheets/sheet1.xml',
                'total_rows' => 200,
                'sheet_crc32' => 0,
                'sync_points' => [],
            ],
        ]);

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/lists sheet .+ more than once/');
        ReaderIndex::decode($payload);
    }

    public function test_rejects_zero_length_sheet_path(): void
    {
        // Hand-craft a body with pathLen=0 (the writer never produces
        // this; any sidecar in the wild that does is corrupt).
        $body = pack('v', 0);                 // pathLen
        $body .= pack('V', 10);               // total_rows
        $body .= pack('V', 0);                // sheet_crc32
        $body .= pack('V', 0);                // sync_count
        $header = WriterIndex::MAGIC.pack('CCv', WriterIndex::VERSION, 0, 1).pack('V', 100).pack('V', crc32($body));

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/sheet path length out of range/');
        ReaderIndex::decode($header.$body);
    }

    public function test_rejects_truncated_sheet_section(): void
    {
        $payload = WriterIndex::encode(100, [[
            'entry' => 'xl/worksheets/sheet1.xml',
            'total_rows' => 10,
            'sheet_crc32' => 0,
            'sync_points' => [
                ['row' => 1, 'comp_offset' => 0, 'uncomp_offset' => 0],
            ],
        ]]);

        // Lop off the last 4 bytes of the body so the sync point is incomplete.
        // CRC32 is recomputed against the truncated body to ensure we test
        // the structural-truncation guard, not the CRC guard.
        $truncatedBody = substr(substr($payload, 16), 0, -4);
        $newHeader = substr($payload, 0, 12).pack('V', crc32($truncatedBody));
        $tampered = $newHeader.$truncatedBody;

        $this->expectException(XlsxReadException::class);
        $this->expectExceptionMessageMatches('/truncated/i');
        ReaderIndex::decode($tampered);
    }
}

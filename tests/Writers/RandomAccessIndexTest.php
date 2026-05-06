<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\RandomAccessIndex;

/**
 * Unit tests for the binary encoder. Each test decodes the produced bytes
 * directly via unpack() rather than going through a parser — both sides of
 * the format contract are pinned this way, so a future reader change that
 * reads field offsets differently can't drift away from the writer.
 */
class RandomAccessIndexTest extends TestCase
{
    public function test_header_layout_is_stable(): void
    {
        $payload = RandomAccessIndex::encode(10000, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 1001, []),
        ]);

        $magic = substr($payload, 0, 4);
        $version = ord($payload[4]);
        $flags = ord($payload[5]);
        $sheetCount = unpack('v', substr($payload, 6, 2))[1];
        $syncPeriod = unpack('V', substr($payload, 8, 4))[1];
        $payloadCrc = unpack('V', substr($payload, 12, 4))[1];

        $this->assertSame('KXSI', $magic);
        $this->assertSame(1, $version);
        $this->assertSame(0, $flags);
        $this->assertSame(1, $sheetCount);
        $this->assertSame(10000, $syncPeriod);

        $body = substr($payload, 16);
        $this->assertSame(crc32($body), $payloadCrc);
    }

    public function test_single_sheet_with_zero_sync_points(): void
    {
        $payload = RandomAccessIndex::encode(10000, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 50, []),
        ]);

        $body = substr($payload, 16);
        $cursor = 0;

        $pathLen = unpack('v', substr($body, $cursor, 2))[1];
        $cursor += 2;
        $path = substr($body, $cursor, $pathLen);
        $cursor += $pathLen;
        $totalRows = unpack('V', substr($body, $cursor, 4))[1];
        $cursor += 4;
        $syncCount = unpack('V', substr($body, $cursor, 4))[1];
        $cursor += 4;

        $this->assertSame('xl/worksheets/sheet1.xml', $path);
        $this->assertSame(50, $totalRows);
        $this->assertSame(0, $syncCount);
        $this->assertSame(strlen($body), $cursor);
    }

    public function test_sync_points_round_trip_uint64_offsets(): void
    {
        // Use offsets large enough to exercise the uint64 encoding (above 2^32).
        $points = [
            ['row' => 102, 'comp_offset' => 4_500, 'uncomp_offset' => 12_000],
            ['row' => 1_000_002, 'comp_offset' => 5_000_000_000, 'uncomp_offset' => 32_000_000_000],
        ];

        $payload = RandomAccessIndex::encode(100, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 1_000_001, $points),
        ]);

        // Walk to the sync points block.
        $body = substr($payload, 16);
        $pathLen = unpack('v', substr($body, 0, 2))[1];
        $offset = 2 + $pathLen + 4 + 4; // path + total_rows + sync_count

        $decoded = [];
        for ($i = 0; $i < count($points); $i++) {
            $decoded[] = [
                'row' => unpack('P', substr($body, $offset, 8))[1],
                'comp_offset' => unpack('P', substr($body, $offset + 8, 8))[1],
                'uncomp_offset' => unpack('P', substr($body, $offset + 16, 8))[1],
            ];
            $offset += 24;
        }

        $this->assertSame($points, $decoded);
    }

    public function test_multi_sheet_sections_appended_in_order(): void
    {
        $payload = RandomAccessIndex::encode(1000, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 100, []),
            $this->sheetSection('xl/worksheets/sheet2.xml', 200, [
                ['row' => 1001, 'comp_offset' => 99, 'uncomp_offset' => 999],
            ]),
            $this->sheetSection('xl/worksheets/sheet3.xml', 300, []),
        ]);

        $sheetCount = unpack('v', substr($payload, 6, 2))[1];
        $this->assertSame(3, $sheetCount);

        $body = substr($payload, 16);
        $cursor = 0;
        $names = [];
        for ($i = 0; $i < 3; $i++) {
            $pathLen = unpack('v', substr($body, $cursor, 2))[1];
            $cursor += 2;
            $names[] = substr($body, $cursor, $pathLen);
            $cursor += $pathLen + 4; // skip total_rows
            $syncCount = unpack('V', substr($body, $cursor, 4))[1];
            $cursor += 4 + 24 * $syncCount;
        }

        $this->assertSame([
            'xl/worksheets/sheet1.xml',
            'xl/worksheets/sheet2.xml',
            'xl/worksheets/sheet3.xml',
        ], $names);
    }

    public function test_crc32_validates_against_body_only(): void
    {
        $points = [
            ['row' => 5, 'comp_offset' => 1, 'uncomp_offset' => 2],
        ];

        $payload = RandomAccessIndex::encode(2, [
            $this->sheetSection('xl/worksheets/sheet1.xml', 4, $points),
        ]);

        $headerCrc = unpack('V', substr($payload, 12, 4))[1];
        $body = substr($payload, 16);

        $this->assertSame(crc32($body), $headerCrc);

        // Corrupting the body invalidates the CRC.
        $tampered = substr($payload, 0, 16).str_repeat('X', strlen($body));
        $tamperedBody = substr($tampered, 16);
        $this->assertNotSame(crc32($tamperedBody), $headerCrc);
    }

    /**
     * @param  list<array{row: int, comp_offset: int, uncomp_offset: int}>  $points
     */
    private function sheetSection(string $entry, int $totalRows, array $points): array
    {
        return [
            'entry' => $entry,
            'total_rows' => $totalRows,
            'sync_points' => $points,
        ];
    }
}

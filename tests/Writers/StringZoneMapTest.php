<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * STRZ (G3) — per-block lexicographic string zone maps, codec round-trip.
 *
 * Increment A: the writer accumulates per-block string [min, max] (unsigned
 * byte / strcmp), caps + deferred-truncates at finishFile, encodes STRZ;
 * the reader decodes it back. The load-bearing gates are (1) SOUNDNESS —
 * every real value sits within its block's truncated [min, max]; and (2)
 * the deferred separator restores pruning on a common-prefix corpus where
 * a fixed 16-byte truncation would collapse every block.
 */
class StringZoneMapTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-strz-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    private function decodeSidecar(): ReaderIndex
    {
        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        return ReaderIndex::decode($cd->readEntry($source, ReaderIndex::ENTRY_PATH));
    }

    private const ENTRY = 'xl/worksheets/sheet1.xml';

    public function test_string_zone_maps_round_trip_and_are_sound(): void
    {
        // 2000 sorted invoice numbers, common prefix "INV-2024-DEPT-042-" (18B).
        $keys = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withStringStats([2])->setBufferFlushInterval(100)->withRandomAccessIndex(every: 100);
        $writer->startFile(['id', 'invoice']);
        for ($i = 1; $i <= 2000; $i++) {
            $k = sprintf('INV-2024-DEPT-042-%08d', $i);
            $keys[$i] = $k;
            $writer->writeRow([$i, $k]);
        }
        $writer->finishFile();

        $idx = $this->decodeSidecar();
        $strz = $idx->columnStringStats(self::ENTRY, 2);
        $this->assertNotNull($strz, 'STRZ section missing');
        $this->assertTrue($strz['sorted_asc'], 'sequential invoice numbers are ascending');

        $blocks = $strz['blocks'];
        // block_count == sync_count + 1 invariant (mirrors STAT).
        $syncCount = count($idx->syncPoints(self::ENTRY));
        $this->assertCount($syncCount + 1, $blocks);

        // SOUNDNESS: every key sits within SOME block's truncated [min,max].
        // We reconstruct the row->block mapping from block row counts.
        $rowsSeen = 0;
        $blockOf = [];
        // block 0 also folds the header (row 1); its count includes it.
        foreach ($blocks as $bi => $b) {
            // each block's count = numeric-string 'other'? no — invoice is a
            // string value, so count>0. Header "invoice" folds into block 0.
            for ($n = 0; $n < $b['count']; $n++) {
                $blockOf[$rowsSeen++] = $bi;
            }
        }
        // Verify soundness directly: for every key, at least one block whose
        // [min,max] contains it must exist (the one it was written into).
        foreach ([1, 2, 137, 999, 1000, 1500, 2000] as $i) {
            $k = $keys[$i];
            $inside = false;
            foreach ($blocks as $b) {
                if ($b['count'] === 0) {
                    continue;
                }
                if (strcmp($b['min'], $k) <= 0 && strcmp($b['max'], $k) >= 0) {
                    $inside = true;
                    break;
                }
            }
            $this->assertTrue($inside, "key {$k} outside every block's [min,max] — STRZ unsound");
        }
    }

    public function test_deferred_separator_restores_pruning_on_common_prefix(): void
    {
        // Common prefix "INVOICE-2024-Q3-" is exactly 16 bytes, so a fixed
        // 16-byte truncation would collapse every block to an identical
        // range (zero pruning). The deferred per-column length must reach
        // past it so a point query prunes to ~one block.
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withStringStats([1])->setBufferFlushInterval(100)->withRandomAccessIndex(every: 100);
        $writer->startFile(['invoice']);
        for ($i = 1; $i <= 2000; $i++) {
            $writer->writeRow([sprintf('INVOICE-2024-Q3-%08d', $i)]);
        }
        $writer->finishFile();

        $idx = $this->decodeSidecar();
        $blocks = $idx->columnStringStats(self::ENTRY, 1)['blocks'];

        // Point-query a value that exists in the middle; count surviving blocks.
        $x = sprintf('INVOICE-2024-Q3-%08d', 1500);
        $survivors = 0;
        foreach ($blocks as $b) {
            if ($b['count'] > 0 && strcmp($b['max'], $x) >= 0 && strcmp($b['min'], $x) <= 0) {
                $survivors++;
            }
        }

        // Deferred truncation reached past the 16B prefix: the sorted key
        // lands in a single block, not all ~21.
        $this->assertLessThanOrEqual(2, $survivors, 'deferred separator failed — pruning collapsed');

        // And prove the min/max actually carry bytes past the 16B prefix
        // (otherwise the survivor count above would be luck).
        $mid = $blocks[intdiv(count($blocks), 2)];
        $this->assertGreaterThan(16, strlen($mid['min']), 'truncation did not reach past the common prefix');
    }

    public function test_no_string_stats_emits_no_strz_and_stays_byte_compatible(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([1])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'name']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow([$i, 'name-'.$i]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $raw = $cd->readEntry($source, ReaderIndex::ENTRY_PATH);

        $this->assertStringNotContainsString('STRZ', $raw, 'STRZ emitted without withStringStats');
        $idx = ReaderIndex::decode($raw);
        $this->assertNull($idx->columnStringStats(self::ENTRY, 1));
        $this->assertSame([], $idx->stringStatsColumns(self::ENTRY));
    }
}

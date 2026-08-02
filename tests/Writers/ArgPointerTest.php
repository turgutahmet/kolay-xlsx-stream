<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * ARGP (F4) — per-block argmin/argmax row numbers, writer + codec
 * round-trip (Increment A). Gates: blocks align 1:1 with STAT; each
 * block's recorded rows really hold that block's min/max (verified via
 * rowAt); first-occurrence semantics on ties; count==0 blocks store {0,0};
 * withArgPointers folds columns into STAT; additive (no ARGP without opt-in).
 */
class ArgPointerTest extends TestCase
{
    private string $testFile;
    private const ENTRY = 'xl/worksheets/sheet1.xml';

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-argp-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    private function decode(): ReaderIndex
    {
        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        return ReaderIndex::decode($cd->readEntry($source, ReaderIndex::ENTRY_PATH));
    }

    public function test_arg_rows_are_1to1_with_stat_and_hold_the_block_extremes(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withArgPointers([2])->withRandomAccessIndex(every: 50)->setBufferFlushInterval(50);
        $writer->startFile(['id', 'amount']);
        // A non-monotone amount so min/max sit at varied rows within blocks.
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([$i, (($i * 37) % 100) + 0.5]);
        }
        $writer->finishFile();

        $index = $this->decode();
        $stat = $index->columnStats(self::ENTRY, 2);
        $argp = $index->argPointers(self::ENTRY, 2);

        $this->assertNotNull($argp);
        $this->assertCount(count($stat['blocks']), $argp, 'ARGP must align 1:1 with STAT blocks');

        // Cross-check every block: the recorded rows must carry the block's
        // stored min/max value (read back through the full reader).
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        foreach ($stat['blocks'] as $i => $block) {
            if ($block['count'] === 0) {
                $this->assertSame(['minRow' => 0, 'maxRow' => 0], $argp[$i]);

                continue;
            }
            $minRowCell = $reader->rowAt($argp[$i]['minRow'])[1];
            $maxRowCell = $reader->rowAt($argp[$i]['maxRow'])[1];
            $this->assertEqualsWithDelta($block['min'], (float) $minRowCell, 1e-9, "block {$i} minRow value");
            $this->assertEqualsWithDelta($block['max'], (float) $maxRowCell, 1e-9, "block {$i} maxRow value");
        }
        $reader->close();
    }

    public function test_global_extreme_row_is_recoverable(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withArgPointers([2])->withRandomAccessIndex(every: 50)->setBufferFlushInterval(50);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 300; $i++) {
            // Unique global max at row 137 (sheet row 138), min at row 200 (sheet 201).
            $v = 10.0 + ($i % 20);
            if ($i === 137) {
                $v = 9999.0;
            }
            if ($i === 200) {
                $v = -5.0;
            }
            $writer->writeRow([$i, $v]);
        }
        $writer->finishFile();

        $index = $this->decode();
        $stat = $index->columnStats(self::ENTRY, 2);
        $argp = $index->argPointers(self::ENTRY, 2);

        // Reproduce a reader-side global argmax: block with the greatest max.
        $bestBlock = null;
        $bestMax = null;
        foreach ($stat['blocks'] as $i => $b) {
            if ($b['count'] > 0 && ($bestMax === null || $b['max'] > $bestMax)) {
                $bestMax = $b['max'];
                $bestBlock = $i;
            }
        }
        $this->assertSame(138, $argp[$bestBlock]['maxRow'], 'global max at sheet row 138');

        // …and global argmin.
        $worstBlock = null;
        $worstMin = null;
        foreach ($stat['blocks'] as $i => $b) {
            if ($b['count'] > 0 && ($worstMin === null || $b['min'] < $worstMin)) {
                $worstMin = $b['min'];
                $worstBlock = $i;
            }
        }
        $this->assertSame(201, $argp[$worstBlock]['minRow'], 'global min at sheet row 201');
    }

    public function test_ties_keep_the_first_occurrence(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withArgPointers([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        // All rows equal → min == max; first data row (sheet row 2) wins both.
        for ($i = 1; $i <= 20; $i++) {
            $writer->writeRow([$i, 7.0]);
        }
        $writer->finishFile();

        $argp = $this->decode()->argPointers(self::ENTRY, 2);
        $this->assertSame(2, $argp[0]['minRow']);
        $this->assertSame(2, $argp[0]['maxRow']);
    }

    public function test_with_arg_pointers_folds_column_into_stat(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        // Only argptr opt-in — STAT must still cover the column.
        $writer->withArgPointers([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 50; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $index = $this->decode();
        $this->assertNotNull($index->columnStats(self::ENTRY, 2), 'argptr column must also carry STAT');
        $this->assertNotNull($index->argPointers(self::ENTRY, 2));
    }

    public function test_no_arg_pointers_emits_no_argp_and_is_additive(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withColumnStats([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 200; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);
        $raw = $cd->readEntry($source, ReaderIndex::ENTRY_PATH);
        $this->assertStringNotContainsString('ARGP', $raw);
        $this->assertNull(ReaderIndex::decode($raw)->argPointers(self::ENTRY, 2));
    }
}

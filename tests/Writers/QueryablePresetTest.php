<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * E3 — queryable(): one call that turns on the whole query stack
 * (random-access index + column zone maps + sketches) for a set of
 * columns, so the common "make this export queryable" setup is a
 * one-liner instead of three chained calls.
 */
class QueryablePresetTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-queryable-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_queryable_enables_index_stats_and_sketches(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->queryable([1, 2]);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 1000; $i++) {
            $writer->writeRow([$i, $i * 2.5]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);

        // Zone maps (columnStats) + point pruning (rowsWhere) + sketches
        // (quantile) must all be available — the three subsystems queryable()
        // switched on.
        $this->assertNotNull($reader->columnStats(2));
        $this->assertSame(1000, $reader->columnStats(1)['count']);
        $this->assertCount(1, iterator_to_array($reader->rowsWhere(1, '=', 500)));
        $this->assertNotNull($reader->quantile(2, 0.5)); // sketch present
        $reader->close();
    }

    public function test_queryable_matches_the_three_separate_calls_byte_for_byte(): void
    {
        $preset = $this->testFile;
        $manual = $this->testFile.'.manual.xlsx';

        $write = function (string $path, bool $usePreset): void {
            $w = new SinkableXlsxWriter(new FileSink($path));
            if ($usePreset) {
                $w->queryable([1, 2], every: 250);
            } else {
                $w->withRandomAccessIndex(every: 250)->withColumnStats([1, 2])->withColumnSketches([1, 2]);
            }
            $w->setBufferFlushInterval(250);
            $w->startFile(['id', 'amount']);
            for ($i = 1; $i <= 1000; $i++) {
                $w->writeRow([$i, $i * 2.5]);
            }
            $w->finishFile();
        };

        try {
            $write($preset, true);
            $write($manual, false);

            // Same configuration => byte-identical file.
            $this->assertSame(
                hash_file('sha256', $manual),
                hash_file('sha256', $preset),
                'queryable() preset must equal the three explicit calls'
            );
        } finally {
            @unlink($manual);
        }
    }

    public function test_queryable_without_sketches(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->queryable([1], withSketches: false);
        $writer->startFile(['id']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow([$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNotNull($reader->columnStats(1));   // stats on
        $this->assertNull($reader->quantile(1, 0.5));     // sketches off
        $reader->close();
    }

    public function test_queryable_throws_after_start(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['id']);

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxStreamException::class);
        $writer->queryable([1]);
    }
}

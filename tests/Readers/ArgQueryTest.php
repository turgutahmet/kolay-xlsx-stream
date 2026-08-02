<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * argMin()/argMax() through the public reader facade — the extreme row is
 * named straight from the ARGP sidecar (no scan) and reads back to the
 * extreme value. Gates: global extreme row + value; the returned row is a
 * rowAt coordinate; ties resolve to the first occurrence; column addressable
 * by name; null contracts (untracked column, text-only column, no index).
 */
class ArgQueryTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-argq-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_argmin_argmax_name_the_global_extreme_rows(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withArgPointers([2])->withRandomAccessIndex(every: 50)->setBufferFlushInterval(50);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 400; $i++) {
            // Non-monotone with a unique global max (row 300 → sheet 301)
            // and a unique global min (row 90 → sheet 91).
            $v = 100.0 + (($i * 31) % 50);
            if ($i === 300) {
                $v = 9999.0;
            }
            if ($i === 90) {
                $v = -12.0;
            }
            $writer->writeRow([$i, $v]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $lo = $reader->argMin('amount');
        $hi = $reader->argMax('amount');

        $this->assertNotNull($lo);
        $this->assertNotNull($hi);

        // Values equal the STAT extremes and the rows read back to them.
        $stats = $reader->columnStats('amount');
        $this->assertEqualsWithDelta($stats['min'], $lo['value'], 1e-9);
        $this->assertEqualsWithDelta($stats['max'], $hi['value'], 1e-9);

        $this->assertSame(91, $lo['row']);
        $this->assertSame(301, $hi['row']);
        $this->assertEqualsWithDelta(-12.0, (float) $reader->rowAt($lo['row'])[1], 1e-9);
        $this->assertEqualsWithDelta(9999.0, (float) $reader->rowAt($hi['row'])[1], 1e-9);
        $reader->close();
    }

    public function test_ties_return_the_first_occurrence(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withArgPointers([2])->withRandomAccessIndex(every: 25)->setBufferFlushInterval(25);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 200; $i++) {
            // A plateau at the max shared by rows 10 and 150; the earliest
            // (sheet row 11) must win. A plateau at the min at rows 40 and 170.
            $v = 5.0;
            if ($i === 10 || $i === 150) {
                $v = 42.0;
            }
            if ($i === 40 || $i === 170) {
                $v = -3.0;
            }
            $writer->writeRow([$i, $v]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSame(11, $reader->argMax('amount')['row'], 'first row achieving the max wins');
        $this->assertSame(41, $reader->argMin('amount')['row'], 'first row achieving the min wins');
        $reader->close();
    }

    public function test_null_contracts(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withArgPointers([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount', 'label']);
        for ($i = 1; $i <= 150; $i++) {
            $writer->writeRow([$i, (float) $i, 'text'.$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->argMin(1), 'untracked column has no ARGP');
        $this->assertNull($reader->argMax(3), 'untracked column has no ARGP');
        $this->assertNotNull($reader->argMin('amount'), 'tracked column answers');
        $reader->close();
    }

    public function test_text_only_column_has_no_numeric_extreme(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        // Track a text column: STAT/ARGP exist but every block is non-numeric.
        $writer->withArgPointers([2])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'label']);
        for ($i = 1; $i <= 120; $i++) {
            $writer->writeRow([$i, 'sku-'.$i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->argMin(2), 'no numeric value → no argmin');
        $this->assertNull($reader->argMax(2), 'no numeric value → no argmax');
        $reader->close();
    }

    public function test_null_without_index(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->argMin('amount'), 'no sidecar → no ARGP');
        $reader->close();
    }
}

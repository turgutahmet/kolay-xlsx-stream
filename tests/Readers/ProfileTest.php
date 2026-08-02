<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * profile() — the one-call data-profiling report, assembled entirely from
 * the sidecar sections. It is a faithful PACKAGE: every field must equal the
 * dedicated method it stands in for (columnStats, quantile, histogram,
 * countDistinct, topValues, countEmpty, correlation), the tracked columns
 * are exactly the union of the section column sets, and a file with no
 * sidecar profiles nothing.
 */
class ProfileTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-profile-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    private function writeRich(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 200)
            ->withColumnStats([2, 3])
            ->withColumnSketches([2, 3])
            ->withTopValues([4])
            ->withCorrelations([2, 3])
            ->setBufferFlushInterval(200);
        $writer->startFile(['id', 'amount', 'score', 'status']);
        $statuses = ['paid', 'pending', 'refunded'];
        for ($i = 1; $i <= 2000; $i++) {
            $amount = $i % 9 === 0 ? '' : ($i % 500) + 0.5;   // some blanks
            $score = ($i * 7919) % 1000;
            $writer->writeRow([$i, $amount, $score, $statuses[$i % 3]]);
        }
        $writer->finishFile();
    }

    public function test_profile_packages_every_field_faithfully(): void
    {
        $this->writeRich();
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $profile = $reader->profile();

        // Tracked columns are the union of the section sets: 2, 3 (STAT +
        // sketches + corr) and 4 (TOPK). id (1) carries nothing → absent.
        $this->assertSame([2, 3, 4], array_keys($profile['columns']));

        // Numeric column 2: every field equals its dedicated method.
        $c2 = $profile['columns'][2];
        $stats = $reader->columnStats(2);
        $this->assertSame('amount', $c2['name']);
        $this->assertSame($stats['count'], $c2['numeric_count']);
        $this->assertSame($stats['min'], $c2['min']);
        $this->assertSame($stats['max'], $c2['max']);
        $this->assertSame($stats['avg'], $c2['avg']);
        $this->assertSame($reader->countEmpty(2), $c2['empty_count']);
        // Percentiles carry the estimate AND its rank certificate.
        $this->assertSame($reader->quantile(2, 0.5), $c2['percentiles']['p50']['value']);
        $this->assertSame($reader->quantile(2, 0.95), $c2['percentiles']['p95']['value']);
        $exp = $reader->explainQuantile(2, 0.5);
        $this->assertSame($exp['rank_lo'], $c2['percentiles']['p50']['rank_lo'], 'certificate must match explainQuantile');
        $this->assertSame($exp['rank_hi'], $c2['percentiles']['p50']['rank_hi']);
        $this->assertEquals($reader->histogram(2), $c2['histogram']);
        $this->assertSame($reader->countDistinct(2), $c2['distinct']);
        $this->assertNull($c2['top_values'], 'column 2 is not TOPK-tracked');

        // Text column 4: categorical fields present, numeric fields null.
        $c4 = $profile['columns'][4];
        $this->assertSame('status', $c4['name']);
        $this->assertNull($c4['numeric_count']);
        $this->assertNull($c4['min']);
        $this->assertNull($c4['percentiles']['p50']['value']);
        $this->assertNull($c4['histogram']);
        $this->assertEquals($reader->topValues(4), $c4['top_values']);

        // Correlations mirror correlation().
        $this->assertEqualsWithDelta($reader->correlation(2, 3), $profile['correlations']['2,3'], 1e-12);

        // Data-row count is reported.
        $this->assertSame(2000, $profile['data_rows']);
        $reader->close();
    }

    public function test_profile_can_target_specific_columns_and_skip_histogram(): void
    {
        $this->writeRich();
        $reader = StreamingXlsxReader::fromFile($this->testFile);

        $profile = $reader->profile(['amount'], histogram: false);
        $this->assertSame([2], array_keys($profile['columns']));
        $this->assertNull($profile['columns'][2]['histogram'], 'histogram opted out');
        // Percentiles still there (cheap, same digest).
        $this->assertNotNull($profile['columns'][2]['percentiles']['p50']['value']);
        $reader->close();
    }

    public function test_percentile_keys_are_lossless(): void
    {
        $this->writeRich();
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        // p99 and p99.9 must NOT collide (the old floor(q*100) key did).
        $profile = $reader->profile(['amount'], histogram: false, percentiles: [0.5, 0.99, 0.999]);
        $keys = array_keys($profile['columns'][2]['percentiles']);
        $this->assertSame(['p50', 'p99', 'p99.9'], $keys);
        $reader->close();
    }

    public function test_negative_or_zero_column_is_rejected(): void
    {
        $this->writeRich();
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->expectException(\InvalidArgumentException::class);
        try {
            $reader->profile([0]);
        } finally {
            $reader->close();
        }
    }

    public function test_no_sidecar_reports_null_data_rows(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        // No index → do NOT scan the whole file just to count rows.
        $this->assertNull($reader->profile()['data_rows']);
        $reader->close();
    }

    public function test_profile_of_a_file_without_a_sidecar_is_empty(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'amount']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i, (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $profile = $reader->profile();
        $this->assertSame([], $profile['columns']);
        $this->assertSame([], $profile['correlations']);
        $reader->close();
    }
}

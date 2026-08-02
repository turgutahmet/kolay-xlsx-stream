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
 * CORR (6E) — pairwise Pearson correlation, writer + codec + reader round
 * trip. Gates: correlation() equals a textbook oracle over the written data;
 * only both-numeric rows feed a pair (blanks/text skipped); the pair is
 * order-independent; untracked pairs / same column / no-index return null;
 * and a file without withCorrelations carries no CORR section (additive).
 */
class CorrelationTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-corr-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    private function oracle(array $x, array $y): float
    {
        $n = count($x);
        $sx = array_sum($x);
        $sy = array_sum($y);
        $sxy = 0.0;
        $sx2 = 0.0;
        $sy2 = 0.0;
        for ($i = 0; $i < $n; $i++) {
            $sxy += $x[$i] * $y[$i];
            $sx2 += $x[$i] ** 2;
            $sy2 += $y[$i] ** 2;
        }

        return ($n * $sxy - $sx * $sy) / sqrt(($n * $sx2 - $sx ** 2) * ($n * $sy2 - $sy ** 2));
    }

    public function test_correlation_matches_the_oracle_over_three_pairs(): void
    {
        mt_srand(61);
        $price = [];
        $qty = [];
        $noise = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withCorrelations([2, 3, 4])->setBufferFlushInterval(200);
        $writer->startFile(['id', 'price', 'qty', 'noise']);
        for ($i = 1; $i <= 5000; $i++) {
            $p = mt_rand(100, 10000) / 100;
            $q = $p * 0.4 + mt_rand(-300, 300) / 100; // correlated with price
            $r = mt_rand(0, 100000) / 100;            // independent
            $price[] = $p;
            $qty[] = $q;
            $noise[] = $r;
            $writer->writeRow([$i, $p, $q, $r]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertEqualsWithDelta($this->oracle($price, $qty), $reader->correlation('price', 'qty'), 1e-9);
        $this->assertEqualsWithDelta($this->oracle($price, $noise), $reader->correlation(2, 4), 1e-9);
        // Order-independent.
        $this->assertSame($reader->correlation(2, 3), $reader->correlation(3, 2));
        $reader->close();
    }

    public function test_only_rows_where_both_are_numeric_contribute(): void
    {
        // Interleave blanks and text on each column; the pair's population is
        // exactly the rows where BOTH are numeric.
        $x = [];
        $y = [];
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withCorrelations([2, 3])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'a', 'b']);
        for ($i = 1; $i <= 1200; $i++) {
            $a = $i % 7 === 0 ? '' : (float) $i;          // blank a
            $b = $i % 5 === 0 ? 'n/a' : (float) ($i * 2); // text b
            $writer->writeRow([$i, $a, $b]);
            if (is_numeric($a) && is_numeric($b)) {
                $x[] = (float) $a;
                $y[] = (float) $b;
            }
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        // a and b are both linear in i, so over the shared rows r == 1.
        $this->assertEqualsWithDelta($this->oracle($x, $y), $reader->correlation(2, 3), 1e-9);
        $this->assertEqualsWithDelta(1.0, $reader->correlation(2, 3), 1e-9);
        $reader->close();
    }

    public function test_null_contracts(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withCorrelations([2, 3])->setBufferFlushInterval(100);
        $writer->startFile(['id', 'a', 'b', 'c']);
        for ($i = 1; $i <= 300; $i++) {
            $writer->writeRow([$i, (float) $i, (float) ($i % 10), (float) $i]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNotNull($reader->correlation(2, 3), 'tracked pair answers');
        $this->assertNull($reader->correlation(2, 4), 'untracked pair (col 4) → null');
        $this->assertNull($reader->correlation(2, 2), 'same column → null');
        $reader->close();
    }

    public function test_no_correlations_emits_no_corr_and_is_additive(): void
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
        $this->assertStringNotContainsString('CORR', $raw);

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->correlation(1, 2), 'no CORR section → null');
        $reader->close();
    }

    public function test_null_without_index(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'a', 'b']);
        for ($i = 1; $i <= 100; $i++) {
            $writer->writeRow([$i, (float) $i, (float) ($i * 2)]);
        }
        $writer->finishFile();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertNull($reader->correlation(2, 3), 'no sidecar → null');
        $reader->close();
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\ProvidesCostHints;
use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * G5 — gap-bridging: when a pruned scan leaves two surviving block runs
 * separated by a short gap, a high-latency Source (via ProvidesCostHints)
 * bridges the gap into ONE ranged read when the gap's compressed bytes
 * transfer faster than a fresh round-trip. Zero-latency Sources (local,
 * no hints) never bridge — run merging stays byte-for-byte the
 * contiguous-only behaviour. Bridging is a fetch-count optimization only;
 * results must always equal the full-scan oracle.
 */
class GapBridgeTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-bridge-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * col2 = 500 only for the first 100 and last 100 ids, else 1 — so a
     * `col2 = 500` query survives the first and last blocks with a big
     * pruned gap between them (two runs).
     */
    private function writeFixture(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 100);
        $writer->withColumnStats([1, 2]);
        $writer->setBufferFlushInterval(100);
        $writer->startFile(['id', 'flag']);
        for ($i = 1; $i <= 2000; $i++) {
            $flag = ($i <= 100 || $i > 1900) ? 500 : 1;
            $writer->writeRow([$i, $flag]);
        }
        $writer->finishFile();
    }

    /** Spy wrapping a source, counting streamFrom calls. Optionally advertises cost hints. */
    private function makeSource(bool $withHints): object
    {
        $inner = new LocalFileSource($this->testFile);

        if ($withHints) {
            return new class ($inner) implements ProvidesCostHints, Source {
                public int $streamCalls = 0;

                public function __construct(private LocalFileSource $inner) {}

                public function costHints(): array
                {
                    // Absurdly generous: a huge byte budget fits any gap,
                    // so every gap should bridge.
                    return ['rtt_us' => 1_000_000, 'bandwidth_bps' => 10_000_000_000];
                }

                public function size(): int
                {
                    return $this->inner->size();
                }

                public function range(int $offset, int $length): string
                {
                    return $this->inner->range($offset, $length);
                }

                public function streamFrom(int $offset, ?int $length = null)
                {
                    $this->streamCalls++;

                    return $this->inner->streamFrom($offset, $length);
                }

                public function close(): void
                {
                    $this->inner->close();
                }
            };
        }

        return new class ($inner) implements Source {
            public int $streamCalls = 0;

            public function __construct(private LocalFileSource $inner) {}

            public function size(): int
            {
                return $this->inner->size();
            }

            public function range(int $offset, int $length): string
            {
                return $this->inner->range($offset, $length);
            }

            public function streamFrom(int $offset, ?int $length = null)
            {
                $this->streamCalls++;

                return $this->inner->streamFrom($offset, $length);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };
    }

    private function expectedRows(): array
    {
        // ids 1..100 and 1901..2000 carry flag 500.
        return array_merge(range(1, 100), range(1901, 2000));
    }

    public function test_no_hints_source_does_not_bridge(): void
    {
        $this->writeFixture();
        $spy = $this->makeSource(withHints: false);
        $reader = StreamingXlsxReader::from($spy);

        $spy->streamCalls = 0;
        $hits = iterator_to_array($reader->rowsWhere(2, '=', 500));

        $this->assertSame($this->expectedRows(), array_map(fn ($r) => (int) $r[0], array_values($hits)));
        // First-block run and last-block run are separate fetches.
        $this->assertGreaterThanOrEqual(2, $spy->streamCalls, 'zero-latency source must not bridge the gap');
        $reader->close();
    }

    public function test_high_latency_source_bridges_gap_into_one_fetch(): void
    {
        $this->writeFixture();
        $spy = $this->makeSource(withHints: true);
        $reader = StreamingXlsxReader::from($spy);

        $spy->streamCalls = 0;
        $hits = iterator_to_array($reader->rowsWhere(2, '=', 500));

        // Same rows — bridging is a fetch optimization, never a result change.
        $this->assertSame($this->expectedRows(), array_map(fn ($r) => (int) $r[0], array_values($hits)));
        // The two runs collapse into a single ranged read.
        $this->assertSame(1, $spy->streamCalls, 'high-latency source must bridge the gap into one fetch');
        $reader->close();
    }
}

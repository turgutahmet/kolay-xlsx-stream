<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Contracts\SupportsBoundedStream;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * The I/O contract of profile() and the sidecar query family — pinned with a
 * stream-counting source so a regression here is caught, not shipped. The
 * numeric queries touch no stream at all; profile() reads NO data rows and
 * names columns with a single BOUNDED header read (streamFromRange, not an
 * open-ended streamFrom to EOF — the leak an earlier profile() had).
 */
class ProfileIoTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-profile-io-'.uniqid('', true).'.xlsx';

        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->withRandomAccessIndex(every: 200)
            ->withColumnStats([2, 3])
            ->withColumnSketches([2, 3])
            ->withCorrelations([2, 3])
            ->setBufferFlushInterval(200);
        $writer->startFile(['id', 'amount', 'score']);
        for ($i = 1; $i <= 5000; $i++) {
            $writer->writeRow([$i, ($i % 500) + 0.5, ($i * 7919) % 1000]);
        }
        $writer->finishFile();
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    private function countingSource(): Source
    {
        $inner = new LocalFileSource($this->testFile);

        return new class ($inner) implements Source, SupportsBoundedStream {
            /** @var list<int> offsets of open-ended streamFrom() calls */
            public array $unbounded = [];

            /** @var list<int> lengths of bounded streamFromRange() calls */
            public array $boundedLengths = [];

            public function __construct(private Source $inner)
            {
            }

            public function size(): int
            {
                return $this->inner->size();
            }

            public function range(int $offset, int $length): string
            {
                return $this->inner->range($offset, $length);
            }

            public function streamFrom(int $offset)
            {
                $this->unbounded[] = $offset;

                return $this->inner->streamFrom($offset);
            }

            public function streamFromRange(int $offset, int $length)
            {
                $this->boundedLengths[] = $length;
                if ($this->inner instanceof SupportsBoundedStream) {
                    return $this->inner->streamFromRange($offset, $length);
                }

                return $this->inner->streamFrom($offset);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };
    }

    public function test_numeric_queries_touch_no_stream(): void
    {
        $spy = $this->countingSource();
        $reader = StreamingXlsxReader::from($spy);
        // Warm the index (opened via range(), not a stream), then clear.
        $reader->columnStats(2);
        $spy->unbounded = [];
        $spy->boundedLengths = [];

        $reader->columnStats(2);
        $reader->quantile(2, 0.95);
        $reader->countDistinct(2);
        $reader->correlation(2, 3);
        $reader->exactQuantile(2, 0.5, maxScanBlocks: 0); // budget 0 → never scans

        $this->assertSame([], $spy->unbounded, 'sidecar queries must open no stream');
        $this->assertSame([], $spy->boundedLengths, 'and no bounded read either');
        $reader->close();
    }

    public function test_name_addressed_query_bounds_its_header_read(): void
    {
        // Resolving a column NAME reads the header — it must be the same
        // bounded read, not an open-ended stream to EOF (the DX headline
        // $reader->quantile('amount', 0.5) must not pull a multi-GB sheet).
        $spy = $this->countingSource();
        $reader = StreamingXlsxReader::from($spy);
        $reader->columnStats(2); // warm index
        $spy->unbounded = [];
        $spy->boundedLengths = [];

        $reader->quantile('amount', 0.5);

        $this->assertSame([], $spy->unbounded, 'name resolution must not open an unbounded stream');
        $this->assertNotEmpty($spy->boundedLengths, 'the header is read once, bounded');
        $this->assertLessThan($spy->size() / 2, max($spy->boundedLengths));

        // The header now shared one cache: a following profile() re-reads
        // nothing (name resolution and header() no longer double-read).
        $spy->boundedLengths = [];
        $reader->profile();
        $this->assertSame([], $spy->boundedLengths, 'header cache is shared across name lookup and profile()');
        $reader->close();
    }

    public function test_corrupt_bounded_read_falls_back_not_empties_the_header(): void
    {
        // A sidecar whose block table is wrong makes the BOUNDED header read
        // yield nothing. header() must fall toward the full read and return
        // the true header — never trust the sidecar to shape the answer, and
        // never cache the empty (SPEC §7: a bad sidecar slows, not changes).
        $inner = new LocalFileSource($this->testFile);
        $sabotage = new class ($inner) implements Source, SupportsBoundedStream {
            public function __construct(private Source $inner)
            {
            }

            public function size(): int
            {
                return $this->inner->size();
            }

            public function range(int $offset, int $length): string
            {
                return $this->inner->range($offset, $length);
            }

            public function streamFrom(int $offset)
            {
                return $this->inner->streamFrom($offset);
            }

            public function streamFromRange(int $offset, int $length)
            {
                // Simulate a block table that points at nothing: the bounded
                // read gets an empty stream.
                return fopen('php://memory', 'r');
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };

        $reader = StreamingXlsxReader::from($sabotage);
        $this->assertSame(['id', 'amount', 'score'], $reader->header(), 'fall back to the full read');
        // Cached, but the fallback value, not the empty bounded result.
        $this->assertSame(['id', 'amount', 'score'], $reader->header());
        $reader->close();
    }

    public function test_profile_reads_no_data_rows_and_bounds_the_header_read(): void
    {
        $spy = $this->countingSource();
        $reader = StreamingXlsxReader::from($spy);
        $reader->columnStats(2); // warm index
        $spy->unbounded = [];
        $spy->boundedLengths = [];

        $profile = $reader->profile();
        $this->assertNotEmpty($profile['columns']);

        // The ONLY read is the header, and it is BOUNDED — never an
        // open-ended streamFrom to EOF (the pre-fix leak).
        $this->assertSame([], $spy->unbounded, 'profile() must not open an unbounded (to-EOF) stream');
        $this->assertNotEmpty($spy->boundedLengths, 'the header is read once, bounded');
        // Bounded to a fraction of the sheet — the first block, not the file.
        $this->assertLessThan($spy->size() / 2, max($spy->boundedLengths), 'header read must not span the sheet');

        // A second profile() with the header now cached reads nothing at all.
        $spy->unbounded = [];
        $spy->boundedLengths = [];
        $reader->profile();
        $this->assertSame([], $spy->unbounded);
        $this->assertSame([], $spy->boundedLengths, 'cached header → profile() is pure sidecar');
        $reader->close();
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Foundation tests for the reader's ZIP container parser.
 *
 * Round-trips real XLSX files produced by SinkableXlsxWriter and asserts
 * that ZipDirectory recovers every entry the writer emitted.
 */
class ZipDirectoryTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-zipdir-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    public function test_parses_central_directory_of_minimal_xlsx(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a', 'b']);
        $writer->writeRow(['x', 'y']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        // Every XLSX must include these standard parts.
        $this->assertTrue($cd->has('[Content_Types].xml'));
        $this->assertTrue($cd->has('xl/workbook.xml'));
        $this->assertTrue($cd->has('xl/worksheets/sheet1.xml'));
        $this->assertTrue($cd->has('_rels/.rels'));
        $this->assertTrue($cd->has('xl/_rels/workbook.xml.rels'));
    }

    public function test_entry_metadata_is_consistent(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['col1', 'col2', 'col3']);
        for ($i = 0; $i < 100; $i++) {
            $writer->writeRow(["row{$i}-a", "row{$i}-b", "row{$i}-c"]);
        }
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $sheet = $cd->entry('xl/worksheets/sheet1.xml');
        $this->assertNotNull($sheet);
        $this->assertGreaterThan(0, $sheet['compressed_size']);
        $this->assertGreaterThan($sheet['compressed_size'], $sheet['uncompressed_size']);
        $this->assertGreaterThan(0, $sheet['offset']);
        $this->assertSame(8, $sheet['method']); // DEFLATE
    }

    public function test_data_offset_points_to_inflatable_payload(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['greeting']);
        $writer->writeRow(['hello world']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $sheet = $cd->entry('xl/worksheets/sheet1.xml');
        $offset = $cd->dataOffset($source, 'xl/worksheets/sheet1.xml');
        $compressed = $source->range($offset, $sheet['compressed_size']);

        $inflated = inflate_init(ZLIB_ENCODING_RAW);
        $xml = inflate_add($inflated, $compressed, ZLIB_FINISH);

        $this->assertNotFalse($xml);
        $this->assertStringContainsString('<row r="1"', $xml);
        $this->assertStringContainsString('hello world', $xml);
    }

    public function test_entry_lookup_returns_null_for_missing_name(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['x']);
        $writer->writeRow(['y']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $this->assertNull($cd->entry('xl/no-such-thing.xml'));
        $this->assertFalse($cd->has('xl/no-such-thing.xml'));
    }

    public function test_data_offset_throws_for_unknown_entry(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['x']);
        $writer->writeRow(['y']);
        $writer->finishFile();

        $source = new LocalFileSource($this->testFile);
        $cd = ZipDirectory::fromSource($source);

        $this->expectException(XlsxReadException::class);
        $cd->dataOffset($source, 'xl/no-such-thing.xml');
    }

    public function test_rejects_non_zip_input(): void
    {
        file_put_contents($this->testFile, 'definitely not a zip archive');

        $source = new LocalFileSource($this->testFile);

        $this->expectException(XlsxReadException::class);
        ZipDirectory::fromSource($source);
    }

    public function test_local_file_source_reports_size_and_serves_ranges(): void
    {
        file_put_contents($this->testFile, str_repeat('abcdefghij', 1000)); // 10000 bytes

        $source = new LocalFileSource($this->testFile);
        $this->assertSame(10000, $source->size());

        $head = $source->range(0, 10);
        $this->assertSame('abcdefghij', $head);

        $mid = $source->range(50, 5);
        $this->assertSame('abcde', $mid);

        $tail = $source->range(9990, 10);
        $this->assertSame('abcdefghij', $tail);
    }

    public function test_local_file_source_stream_from_offset(): void
    {
        file_put_contents($this->testFile, str_repeat('0123456789', 100)); // 1000 bytes

        $source = new LocalFileSource($this->testFile);
        $stream = $source->streamFrom(990);

        $tail = '';
        while (! feof($stream)) {
            $chunk = fread($stream, 64);
            if ($chunk === false) {
                break;
            }
            $tail .= $chunk;
        }
        fclose($stream);

        $this->assertSame('0123456789', $tail);
    }

    public function test_local_file_source_tail_returns_suffix_and_size(): void
    {
        file_put_contents($this->testFile, str_repeat('abcdefghij', 1000)); // 10000 bytes

        $source = new LocalFileSource($this->testFile);

        ['data' => $data, 'size' => $size] = $source->tail(10);
        $this->assertSame('abcdefghij', $data);
        $this->assertSame(10000, $size);

        // Suffix larger than the file returns the whole content.
        ['data' => $all, 'size' => $size2] = $source->tail(1_000_000);
        $this->assertSame(10000, strlen($all));
        $this->assertSame(10000, $size2);
    }

    public function test_read_entry_coalesces_lfh_and_body_into_one_range_fetch(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a', 'b']);
        $writer->writeRow(['x', 'y']);
        $writer->finishFile();

        $inner = new LocalFileSource($this->testFile);
        $spy = new RangeCountingSource($inner);
        $cd = ZipDirectory::fromSource($spy);

        $spy->resetRangeCalls();

        // First read: LFH + body must arrive in a single range() call —
        // the old flow paid a 30-byte LFH fetch plus a body fetch.
        $xml = $cd->readEntry($spy, 'xl/workbook.xml');
        $this->assertStringContainsString('<sheets>', $xml);
        $this->assertSame(1, $spy->rangeCallCount(), 'coalesced read must be one range fetch');
        $this->assertSame(0, $spy->lfhRangeCallCount(), 'no separate 30-byte LFH fetch expected');

        // The coalesced read populates the dataOffset cache, so a
        // later dataOffset() is free and a re-read costs one body fetch.
        $spy->resetRangeCalls();
        $cd->dataOffset($spy, 'xl/workbook.xml');
        $this->assertSame(0, $spy->rangeCallCount(), 'dataOffset must come from the coalesced read cache');

        $cd->readEntry($spy, 'xl/workbook.xml');
        $this->assertSame(1, $spy->rangeCallCount(), 'cached-offset re-read is a single body fetch');
    }

    public function test_read_entry_above_coalesce_cap_falls_back_to_two_step(): void
    {
        // Incompressible payload > 1 MB: compressed_size exceeds the
        // coalesce cap, so the coalesced prefetch must NOT trigger —
        // big bodies keep the lazy LFH + body flow.
        $big = random_bytes(1_200_000);

        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->testFile, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true);
        $zip->addFromString('big.bin', $big);
        $zip->close();

        $inner = new LocalFileSource($this->testFile);
        $spy = new RangeCountingSource($inner);
        $cd = ZipDirectory::fromSource($spy);

        $this->assertGreaterThan(1024 * 1024, $cd->entry('big.bin')['compressed_size']);

        $spy->resetRangeCalls();
        $payload = $cd->readEntry($spy, 'big.bin');

        $this->assertSame($big, $payload);
        $this->assertSame(2, $spy->rangeCallCount(), 'oversize entry must use LFH fetch + body fetch');
        $this->assertSame(1, $spy->lfhRangeCallCount());
    }

    public function test_reader_construction_costs_three_range_fetches(): void
    {
        // Static RTT budget for opening a workbook on a range source
        // without suffix support: 1 tail read (EOCD + CD) + 1 coalesced
        // workbook.xml + 1 coalesced workbook.xml.rels. Guards against
        // regressions re-introducing the duplicate workbook.xml fetch
        // or the separate LFH fetches. (Suffix-capable sources also
        // fold the size() lookup into the tail read.)
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['a']);
        $writer->writeRow(['x']);
        $writer->finishFile();

        $spy = new RangeCountingSource(new LocalFileSource($this->testFile));
        \Kolay\XlsxStream\Readers\StreamingXlsxReader::from($spy);

        $this->assertSame(3, $spy->rangeCallCount());
    }

    public function test_s3_range_source_returns_empty_string_for_non_positive_length(): void
    {
        // length <= 0 would render as `bytes=X-(X-1)` — an unsatisfiable
        // range S3 answers with 416. The guard must return '' without
        // ever touching the network (the fake client would fail loudly).
        $s3 = new \Aws\S3\S3Client([
            'region' => 'eu-west-1',
            'version' => 'latest',
            'credentials' => ['key' => 'test', 'secret' => 'test'],
        ]);
        $source = new \Kolay\XlsxStream\Sources\S3RangeSource($s3, 'bucket', 'key');

        $this->assertSame('', $source->range(100, 0));
        $this->assertSame('', $source->range(100, -5));
    }

    public function test_data_offset_is_cached_per_entry(): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['x']);
        $writer->writeRow(['y']);
        $writer->finishFile();

        $inner = new LocalFileSource($this->testFile);
        $spy = new RangeCountingSource($inner);
        $cd = ZipDirectory::fromSource($spy);

        $spy->resetRangeCalls();

        $first = $cd->dataOffset($spy, 'xl/worksheets/sheet1.xml');
        $this->assertSame(1, $spy->lfhRangeCallCount(), 'first call must hit Source::range for the LFH');

        $second = $cd->dataOffset($spy, 'xl/worksheets/sheet1.xml');
        $this->assertSame($first, $second);
        $this->assertSame(
            1,
            $spy->lfhRangeCallCount(),
            'cached offset must be returned without a second LFH range fetch'
        );
    }
}

/**
 * Counts range() calls — every call, plus the 30-byte fetches
 * specifically (the Local File Header path used by
 * ZipDirectory::dataOffset). Deliberately does NOT implement
 * SupportsSuffixRange so the two-step size()+range tail flow stays
 * exercised.
 */
class RangeCountingSource implements \Kolay\XlsxStream\Contracts\Source
{
    private int $lfhCalls = 0;

    private int $rangeCalls = 0;

    public function __construct(private LocalFileSource $inner) {}

    public function size(): int
    {
        return $this->inner->size();
    }

    public function range(int $offset, int $length): string
    {
        $this->rangeCalls++;
        if ($length === 30) {
            $this->lfhCalls++;
        }

        return $this->inner->range($offset, $length);
    }

    public function streamFrom(int $offset)
    {
        return $this->inner->streamFrom($offset);
    }

    public function close(): void
    {
        $this->inner->close();
    }

    public function resetRangeCalls(): void
    {
        $this->lfhCalls = 0;
        $this->rangeCalls = 0;
    }

    public function lfhRangeCallCount(): int
    {
        return $this->lfhCalls;
    }

    public function rangeCallCount(): int
    {
        return $this->rangeCalls;
    }
}

<?php

namespace Kolay\XlsxStream\Tests\SpecVectors;

use Kolay\XlsxStream\Readers\RandomAccessIndex as ReaderIndex;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * KXSI conformance suite (SPEC.md §8). Reads the COMMITTED fixture
 * files — it regenerates nothing — so these assertions pin the on-disk
 * format against accidental drift: any writer or decoder change that
 * alters the sidecar bytes or their meaning fails here and forces a
 * deliberate spec review (regenerate via tests/SpecVectors/generate.php
 * and commit the diff alongside the SPEC.md change).
 *
 * Other implementations (a TS reader is planned) consume the same
 * .xlsx / .expected.json / .sidecar.hex triples as their conformance
 * input.
 */
class SpecVectorsTest extends TestCase
{
    /** @return array<string, array{string}> */
    public static function vectors(): array
    {
        return [
            'plain indexed' => ['vector-01-plain-indexed'],
            'indexed + stats' => ['vector-02-stats'],
            'multi-sheet + stats' => ['vector-03-multisheet'],
            'sorted + unsorted stats' => ['vector-04-sorted'],
            'sketches (TDIG + CHLL)' => ['vector-05-sketches'],
        ];
    }

    #[\PHPUnit\Framework\Attributes\DataProvider('vectors')]
    public function test_sidecar_bytes_match_committed_hexdump(string $name): void
    {
        $payload = $this->sidecarPayload($name);
        $expectedHex = preg_replace('/\s+/', '', (string) file_get_contents($this->path($name, 'sidecar.hex')));

        $this->assertSame(
            $expectedHex,
            bin2hex($payload),
            "sidecar bytes for {$name} drifted from the committed hexdump — ".
            'a format change must be deliberate: update SPEC.md and regenerate the vectors'
        );
    }

    #[\PHPUnit\Framework\Attributes\DataProvider('vectors')]
    public function test_decoded_sidecar_matches_expected_json(string $name): void
    {
        $golden = json_decode((string) file_get_contents($this->path($name, 'expected.json')), true);
        $index = ReaderIndex::decode($this->sidecarPayload($name));

        $this->assertSame($golden['sync_period'], $index->syncPeriod());

        foreach ($golden['sheets'] as $sheet) {
            $entry = $sheet['entry'];

            $this->assertSame($sheet['total_rows'], $index->totalRows($entry), $entry);
            $this->assertSame($sheet['sheet_crc32'], $index->sheetCrc32($entry), $entry);
            $this->assertSame($sheet['sync_points'], $index->syncPoints($entry), $entry);
            $this->assertSame($sheet['sync_point_crcs'], $index->syncPointCrcs($entry), $entry);

            $actualStats = [];
            foreach ($index->statsColumns($entry) as $col) {
                $actualStats[(string) $col] = $index->columnStats($entry, $col);
            }
            // assertEquals: JSON stores integral float64 stats (e.g. a
            // sum of 4545.0) as JSON numbers that decode to int — the
            // values are identical, only the PHP scalar type differs.
            $this->assertEquals($sheet['column_stats'], $actualStats, $entry);

            // Sketch goldens pin the ESTIMATORS against the committed
            // bytes: quantile interpolation over TDIG centroids and the
            // CHLL count must reproduce the recorded values exactly
            // (both are deterministic arithmetic over the fixed bytes).
            // Vectors committed before TDIG/CHLL existed carry no key.
            $actualSketches = [];
            foreach ($index->digestColumns($entry) as $col) {
                $digest = $index->columnDigest($entry, $col);
                $quantiles = [];
                foreach (['0', '0.25', '0.5', '0.75', '1'] as $q) {
                    $quantiles[$q] = $digest->quantile((float) $q);
                }
                $actualSketches[(string) $col] = [
                    'quantiles' => $quantiles,
                    'numeric_count' => $digest->count(),
                    'distinct' => $index->columnHll($entry, $col)?->count(),
                ];
            }
            $this->assertEquals($sheet['column_sketches'] ?? [], $actualSketches, $entry);
        }
    }

    /**
     * The sketch vector must also answer through the public reader
     * facade — sidecar-only quantile/median/countDistinct on a real
     * committed workbook, including the null contracts (text column
     * has no numeric digest values; untracked column has no sketch).
     */
    public function test_sketch_vector_answers_through_reader_facade(): void
    {
        $golden = json_decode((string) file_get_contents($this->path('vector-05-sketches', 'expected.json')), true);
        $expected = $golden['sheets'][0]['column_sketches'];

        $reader = StreamingXlsxReader::fromFile($this->path('vector-05-sketches', 'xlsx'));

        $this->assertEquals($expected['2']['quantiles']['0.5'], $reader->median(2));
        $this->assertEquals($expected['2']['quantiles']['0.25'], $reader->quantile(2, 0.25));
        $this->assertEquals($expected['2']['quantiles']['1'], $reader->quantile(2, 1.0));
        $this->assertSame($expected['2']['distinct'], $reader->countDistinct(2));
        $this->assertSame($expected['3']['distinct'], $reader->countDistinct(3));

        $this->assertNull($reader->median(3), 'text column carries no numeric digest values');
        $this->assertNull($reader->quantile(1, 0.5), 'untracked column carries no sketch');
        $this->assertNull($reader->countDistinct(1), 'untracked column carries no sketch');
        $reader->close();
    }

    #[\PHPUnit\Framework\Attributes\DataProvider('vectors')]
    public function test_scrc_values_verify_against_inflated_sheet_prefixes(string $name): void
    {
        // Semantics, recomputed from scratch: SCRC value k MUST equal
        // crc32 of the sheet's uncompressed bytes [0, uncomp_offset_k),
        // and the core sheet_crc32 MUST equal crc32 of the whole entry
        // (which in turn matches the live ZIP central directory).
        $source = new LocalFileSource($this->path($name, 'xlsx'));
        $cd = ZipDirectory::fromSource($source);
        $index = ReaderIndex::decode($cd->readEntry($source, ReaderIndex::ENTRY_PATH));

        $golden = json_decode((string) file_get_contents($this->path($name, 'expected.json')), true);

        foreach ($golden['sheets'] as $sheet) {
            $entry = $sheet['entry'];
            $sheetXml = $cd->readEntry($source, $entry);

            $points = $index->syncPoints($entry);
            $crcs = $index->syncPointCrcs($entry);
            $this->assertSameSize($points, $crcs, $entry);

            foreach ($points as $k => $sp) {
                $this->assertSame(
                    crc32(substr($sheetXml, 0, $sp['uncomp_offset'])),
                    $crcs[$k],
                    "{$name} {$entry} SCRC[{$k}]"
                );
            }

            $this->assertSame(crc32($sheetXml), $index->sheetCrc32($entry), $entry);
            $this->assertSame($cd->entry($entry)['crc32'], $index->sheetCrc32($entry), $entry);
        }
    }

    #[\PHPUnit\Framework\Attributes\DataProvider('vectors')]
    public function test_fixture_opens_through_streaming_reader(string $name): void
    {
        // The vectors are real workbooks, not synthetic blobs — the full
        // reader must serve indexed row counts and random access on them.
        $golden = json_decode((string) file_get_contents($this->path($name, 'expected.json')), true);
        $reader = StreamingXlsxReader::fromFile($this->path($name, 'xlsx'));

        foreach ($golden['sheets'] as $i => $sheet) {
            $reader->onSheetIndex($i);
            $this->assertSame($sheet['total_rows'], $reader->rowCount(), $sheet['entry']);
            $this->assertNotNull($reader->rowAt($sheet['total_rows']), $sheet['entry']);
            $this->assertNull($reader->rowAt($sheet['total_rows'] + 1), $sheet['entry']);
        }
    }

    private function path(string $name, string $ext): string
    {
        return __DIR__."/{$name}.{$ext}";
    }

    private function sidecarPayload(string $name): string
    {
        $source = new LocalFileSource($this->path($name, 'xlsx'));
        $cd = ZipDirectory::fromSource($source);

        return $cd->readEntry($source, ReaderIndex::ENTRY_PATH);
    }
}

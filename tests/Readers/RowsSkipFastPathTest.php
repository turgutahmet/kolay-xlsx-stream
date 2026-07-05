<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;
use ZipArchive;

/**
 * Pins the public rows(skip, limit) contract and verifies the v3.2 skip
 * fast path against it.
 *
 * The contract these tests freeze is exactly what v3.1 shipped:
 *
 *   - rows() yields SEQUENTIALLY RE-KEYED rows — generator keys are
 *     0-based and count EMITTED rows, regardless of skip. They are NOT
 *     sheet row numbers (rowRange/rowsWhere carry those instead).
 *   - skip counts yielded rows from the top of the sheet INCLUDING the
 *     header: skip=0 yields the header first, skip=1 drops exactly it.
 *   - rows(skip) is always equal to "rows() with the first skip yields
 *     dropped" — the fast path may change how rows are skipped (index
 *     seek + boundary scan) but never what is yielded.
 *
 * The oracle throughout is a full rows() materialisation compared with
 * array_slice — any divergence between the skip machinery and the plain
 * sequential path fails loudly.
 */
class RowsSkipFastPathTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-skip-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * 500 data rows. With $indexed, sync every 100 rows -> ~5 blocks so
     * skips land before, on, and after sync points.
     */
    private function writeFixture(bool $indexed): void
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        if ($indexed) {
            $writer->withRandomAccessIndex(every: 100);
            $writer->setBufferFlushInterval(100);
        }
        $writer->startFile(['id', 'name']);
        for ($i = 1; $i <= 500; $i++) {
            $writer->writeRow([$i, 'name-'.$i]);
        }
        $writer->finishFile();
    }

    /**
     * @return list<array<int, mixed>>
     */
    private function assertSkipContractHolds(StreamingXlsxReader $reader): array
    {
        $full = iterator_to_array($reader->rows());

        // Keys are 0-based sequential; the header is the first yield.
        $this->assertSame(range(0, count($full) - 1), array_keys($full));
        $this->assertSame(['id', 'name'], $full[0]);

        // Skips spanning "before / on / just past a sync point", deep
        // skips, the exact end, and past-the-end.
        foreach ([1, 3, 99, 100, 101, 250, 449, 500, 501, 999] as $skip) {
            $got = iterator_to_array($reader->rows($skip));
            $this->assertSame(
                array_slice($full, $skip),
                $got,
                "rows(skip={$skip}) diverged from the sequential contract"
            );
            // Re-keying: emitted rows always count from 0.
            $this->assertSame(
                $got === [] ? [] : range(0, count($got) - 1),
                array_keys($got),
                "rows(skip={$skip}) keys are not sequentially re-keyed"
            );
        }

        // limit interplay on the skipping path.
        $this->assertSame(array_slice($full, 250, 10), iterator_to_array($reader->rows(250, 10)));
        $this->assertSame(array_slice($full, 499, 5), iterator_to_array($reader->rows(499, 5)));
        $this->assertSame([], iterator_to_array($reader->rows(250, 0)));

        // Negative skip clamps to 0.
        $this->assertSame($full, iterator_to_array($reader->rows(-3)));

        return $full;
    }

    public function test_rows_skip_contract_with_index(): void
    {
        $this->writeFixture(indexed: true);
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSkipContractHolds($reader);
        $reader->close();
    }

    public function test_rows_skip_contract_without_index(): void
    {
        $this->writeFixture(indexed: false);
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertSkipContractHolds($reader);
        $reader->close();
    }

    public function test_rows_skip_applies_casts_and_feeds_chunked(): void
    {
        $this->writeFixture(indexed: true);
        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $reader->castColumn(0, 'int');

        // Casts must apply on the skipping path exactly as on the
        // sequential path (they run at yield time, after any skip).
        $rows = iterator_to_array($reader->rows(250, 2));
        $this->assertSame(250, $rows[0][0]);
        $this->assertSame(251, $rows[1][0]);

        // chunked() delegates to rows($skip) — same rows, batched.
        $batches = iterator_to_array($reader->chunked(100, 401));
        $this->assertCount(1, $batches);
        $this->assertCount(100, $batches[0]);
        $this->assertSame(401, $batches[0][0][0]);
        $this->assertSame(500, $batches[0][99][0]);
        $reader->close();
    }

    /**
     * The indexed fast path must SEEK, not scan: a deep skip has to open
     * the sheet stream at a sync-point offset past the sheet start.
     * Identical results alone can't prove that (a regression back to a
     * full scan would still pass the oracle), so spy on streamFrom().
     */
    public function test_rows_skip_seeks_via_index(): void
    {
        $this->writeFixture(indexed: true);

        $spy = new class(new LocalFileSource($this->testFile)) implements Source
        {
            /** @var list<int> */
            public array $streamOffsets = [];

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
                $this->streamOffsets[] = $offset;

                return $this->inner->streamFrom($offset, $length);
            }

            public function close(): void
            {
                $this->inner->close();
            }
        };

        $reader = StreamingXlsxReader::from($spy);

        $reader->rows()->current(); // sequential path: opens at the sheet start
        $sheetStart = min($spy->streamOffsets);
        $spy->streamOffsets = [];

        // Skip deep into the last index block; the stream must open past
        // the sheet start (at a sync-point offset).
        iterator_to_array($reader->rows(450, 1));
        $this->assertNotEmpty($spy->streamOffsets);
        $this->assertGreaterThan($sheetStart, min($spy->streamOffsets));
        $reader->close();
    }

    /**
     * The no-index boundary-scan skip counts '</row>' closings instead
     * of tokenizing. Its equivalence argument (see countRows) leans on
     * self-closing rows yielding nothing — so exercise exactly that
     * external-writer shape, plus sparse cells and inter-row whitespace,
     * for EVERY possible skip against the tokenized full read.
     */
    public function test_rows_skip_matches_on_external_sheet_with_self_closing_rows(): void
    {
        // <row r="3"/> is a legal empty row: rows() folds it into the
        // next row's slice, so it must stay invisible to skip counting.
        $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            "<sheetData>\n".
            '<row r="1"><c r="A1" t="inlineStr"><is><t>h1</t></is></c></row>'."\n".
            '<row r="2"><c r="A2" t="n"><v>1</v></c><c r="B2" t="inlineStr"><is><t>a</t></is></c></row>'."\n".
            '<row r="3"/>'."\n".
            '<row r="4"><c r="B4" t="n"><v>2</v></c></row>'."\n".
            '<row r="5"><c r="A5" t="n"><v>3</v></c></row>'."\n".
            '<row r="6"><c r="A6" t="inlineStr"><is><t>tail</t></is></c></row>'."\n".
            '</sheetData></worksheet>';

        $zip = new ZipArchive();
        $this->assertTrue($zip->open($this->testFile, ZipArchive::CREATE | ZipArchive::OVERWRITE));
        $zip->addFromString('[Content_Types].xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'.
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'.
            '<Default Extension="xml" ContentType="application/xml"/>'.
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'.
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.
            '</Types>');
        $zip->addFromString('_rels/.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'.
            '</Relationships>');
        $zip->addFromString('xl/workbook.xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.
            '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
        $zip->addFromString('xl/_rels/workbook.xml.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'.
            '</Relationships>');
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheetXml);
        $zip->close();

        $reader = StreamingXlsxReader::fromFile($this->testFile);
        $full = iterator_to_array($reader->rows());
        $this->assertCount(5, $full); // 6 <row> opens, but r=3 self-closes into r=4's slice

        for ($skip = 0; $skip <= count($full) + 1; $skip++) {
            $this->assertSame(
                array_slice($full, $skip),
                iterator_to_array($reader->rows($skip)),
                "external-shape rows(skip={$skip}) diverged"
            );
        }
        $reader->close();
    }
}

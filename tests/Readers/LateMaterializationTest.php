<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StreamingXlsxReader;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * Late materialization (G4) through the public reader. Its promise is a
 * pure query-path optimization: rowsWhere returns EXACTLY the same rows
 * whether the scan probes-then-materializes or tokenizes eagerly, and the
 * auto planner engages it only when the sidecar predicts a selective
 * predicate (so the ≈100%-selectivity case, where probing is overhead,
 * stays on the eager path). These gates pin byte-identity across the
 * on/off × classic/compact matrix and the planner's decision boundary.
 */
class LateMaterializationTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir().'/kxs-latemat-'.uniqid('', true).'.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->testFile);
        parent::tearDown();
    }

    /**
     * Wide fixture: an id, a uniform numeric 'amount' (stat+sketch tracked
     * so the planner has an estimate), a text 'status' with a handful of
     * distinct values (STRZ tracked), and padding columns so the deferred
     * body parse is worth measuring.
     */
    private function writeWide(int $rows = 4000, bool $compact = false): void
    {
        $header = ['id', 'amount', 'status'];
        for ($c = 3; $c < 14; $c++) {
            $header[] = 'c'.$c;
        }

        $w = new SinkableXlsxWriter(new FileSink($this->testFile));
        $w->withRandomAccessIndex(every: 400);
        $w->withColumnStats([2]);
        $w->withColumnSketches([2]);
        $w->withStringStats([3]);
        if ($compact) {
            $w->compact();
        }
        $w->setBufferFlushInterval(400);
        $w->startFile($header);
        $statuses = ['alpha', 'bravo', 'charlie', 'delta'];
        for ($i = 1; $i <= $rows; $i++) {
            $row = [$i, ($i * 7919) % 1000, $statuses[$i % 4]];
            for ($c = 3; $c < 14; $c++) {
                $row[] = 'cell-'.$i.'-'.$c;
            }
            $w->writeRow($row);
        }
        $w->finishFile();
    }

    /** Full result map (row number => full row) for a numeric predicate. */
    private function collect(StreamingXlsxReader $r, int|string $col, string $op, $val, ?bool $mode): array
    {
        $r->useLateMaterialization($mode);

        return iterator_to_array($r->rowsWhere($col, $op, $val));
    }

    public function test_numeric_results_are_byte_identical_on_vs_off(): void
    {
        $this->writeWide();

        foreach ([['>=', 990], ['>=', 500], ['>=', 100], ['<', 50], ['between', 200]] as $case) {
            [$op, $val] = $case;
            $v2 = $op === 'between' ? 300 : null;

            $r = StreamingXlsxReader::fromFile($this->testFile);
            $r->useLateMaterialization(false);
            $eager = iterator_to_array($r->rowsWhere('amount', $op, $val, $v2));
            $r->useLateMaterialization(true);
            $late = iterator_to_array($r->rowsWhere('amount', $op, $val, $v2));
            $r->close();

            $this->assertSame(array_keys($eager), array_keys($late), "row keys for {$op} {$val}");
            $this->assertEquals($eager, $late, "full rows for {$op} {$val}");
            $this->assertNotEmpty($eager, "predicate {$op} {$val} should match some rows");
        }
    }

    public function test_classic_and_compact_agree_across_the_late_mat_matrix(): void
    {
        // 2×2: {classic, compact} × {late on, off}. Compact rows are r-less
        // (positional) with <c/> gaps — the shape tokenizeColumn's maxIdx+1
        // branch must resolve. All four cells must equal the classic-eager
        // oracle for the same predicate.
        $this->writeWide(compact: false);
        $r = StreamingXlsxReader::fromFile($this->testFile);
        $oracle = $this->collect($r, 'amount', '>=', 700, false);
        $r->close();
        $this->assertNotEmpty($oracle);

        $r = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertEquals($oracle, $this->collect($r, 'amount', '>=', 700, true), 'classic + late');
        $r->close();

        $this->writeWide(compact: true);
        $r = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertEquals($oracle, $this->collect($r, 'amount', '>=', 700, false), 'compact + eager');
        $this->assertEquals($oracle, $this->collect($r, 'amount', '>=', 700, true), 'compact + late');
        $r->close();
    }

    public function test_auto_planner_turns_off_near_full_selectivity(): void
    {
        $this->writeWide();
        $r = StreamingXlsxReader::fromFile($this->testFile);

        // Selective (~1%): planner engages late materialization.
        $this->assertTrue(
            $r->explain([['amount', '>=', 990]])['lateMaterialization'],
            'a selective predicate should late-materialize'
        );

        // ≈100% selectivity: every scanned row matches, so probing first is
        // pure overhead — the named regression the gate must prevent.
        $this->assertFalse(
            $r->explain([['amount', '>=', 1]])['lateMaterialization'],
            'a predicate matching ~everything must stay eager'
        );
        $r->close();
    }

    public function test_auto_planner_stays_off_without_a_selectivity_estimate(): void
    {
        // 'amount' has STAT (zone maps) but NO sketch → no selectivity
        // estimate → auto must not engage (no regression on the eager path).
        $header = ['id', 'amount'];
        for ($c = 2; $c < 14; $c++) {
            $header[] = 'c'.$c;
        }
        $w = new SinkableXlsxWriter(new FileSink($this->testFile));
        $w->withRandomAccessIndex(every: 200)->withColumnStats([2])->setBufferFlushInterval(200);
        $w->startFile($header);
        for ($i = 1; $i <= 1000; $i++) {
            $row = [$i, $i % 100];
            for ($c = 2; $c < 14; $c++) {
                $row[] = 'x'.$i.'-'.$c;
            }
            $w->writeRow($row);
        }
        $w->finishFile();

        $r = StreamingXlsxReader::fromFile($this->testFile);
        $this->assertFalse(
            $r->explain([['amount', '>=', 99]])['lateMaterialization'],
            'no sketch estimate → auto late-mat stays off'
        );
        // Forcing it on must still be byte-identical.
        $eager = $this->collect($r, 'amount', '>=', 99, false);
        $late = $this->collect($r, 'amount', '>=', 99, true);
        $this->assertEquals($eager, $late);
        $r->close();
    }

    public function test_string_predicate_late_mat_is_identical(): void
    {
        // STRZ column: a string rowsWhere over a candidate block set must be
        // byte-identical under late materialization (forced on), and the
        // auto planner engages when the prune leaves few surviving blocks.
        $this->writeWide();
        $r = StreamingXlsxReader::fromFile($this->testFile);

        $eager = $this->collect($r, 'status', '=', 'charlie', false);
        $late = $this->collect($r, 'status', '=', 'charlie', true);
        $this->assertNotEmpty($eager);
        $this->assertSame(array_keys($eager), array_keys($late));
        $this->assertEquals($eager, $late, 'string predicate results must match under late-mat');

        // Auto mode must be byte-identical too, whichever way the STRZ
        // prune-ratio gate decides — a selective '=' (few blocks survive)
        // and an unselective '>=' that every value clears.
        $this->assertEquals($eager, $this->collect($r, 'status', '=', 'charlie', null), 'auto string result');
        $this->assertEquals(
            $this->collect($r, 'status', '>=', 'a', false),
            $this->collect($r, 'status', '>=', 'a', null),
            'auto string result on an unselective predicate'
        );
        $r->close();
    }

    public function test_absent_predicate_column_is_handled(): void
    {
        // A predicate column that some rows omit: the probe returns '' (a
        // non-match), exactly as the eager path reads '' — no row wrongly
        // kept or dropped.
        $header = ['id', 'amount'];
        for ($c = 2; $c < 12; $c++) {
            $header[] = 'c'.$c;
        }
        $w = new SinkableXlsxWriter(new FileSink($this->testFile));
        $w->withRandomAccessIndex(every: 200)->withColumnStats([2])->withColumnSketches([2])->setBufferFlushInterval(200);
        $w->startFile($header);
        for ($i = 1; $i <= 1000; $i++) {
            // Every 5th row leaves amount blank (empty string → non-numeric).
            $row = [$i, $i % 5 === 0 ? '' : ($i % 200)];
            for ($c = 2; $c < 12; $c++) {
                $row[] = 'y'.$i.'-'.$c;
            }
            $w->writeRow($row);
        }
        $w->finishFile();

        $r = StreamingXlsxReader::fromFile($this->testFile);
        $eager = $this->collect($r, 'amount', '>=', 150, false);
        $late = $this->collect($r, 'amount', '>=', 150, true);
        $r->close();
        $this->assertNotEmpty($eager);
        $this->assertEquals($eager, $late);
    }
}

<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\CellTokenizer;
use Kolay\XlsxStream\Readers\InMemorySharedStrings;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * tokenizeColumn() — the late-materialization single-column extractor.
 * Its whole contract is: for any row XML and any target index, return the
 * SAME value tokenizeRow() places at that index (which the predicate then
 * sees), while parsing at most the target cell's body. The oracle here is
 * therefore tokenizeRow itself: for a battery of cell shapes and every
 * index (including gaps and out-of-range), the two must agree cell-for-cell.
 */
class TokenizeColumnTest extends TestCase
{
    /**
     * Cross-check tokenizeColumn against tokenizeRow at every index from 0
     * up to two past the last populated column (so absent / beyond-range
     * targets are covered), for one row XML blob.
     */
    private function assertMatchesOracle(string $rowXml, ?InMemorySharedStrings $sst = null, int $probeUpTo = 12): void
    {
        $full = CellTokenizer::tokenizeRow($rowXml, $sst);
        for ($idx = 0; $idx <= $probeUpTo; $idx++) {
            // tokenizeRow yields a dense array up to its max column; a
            // target past that is absent. The predicate path reads
            // $row[$idx] ?? null, but '' and null are equivalent to both
            // cellMatches and cellMatchesString (neither ever matches), so
            // the extractor is free to normalise absent to ''.
            $expected = $full[$idx] ?? '';
            $actual = CellTokenizer::tokenizeColumn($rowXml, $idx, $sst);
            $this->assertSame($expected, $actual, "index {$idx} of: {$rowXml}");
        }
    }

    public function test_writer_shaped_dense_row_every_shape(): void
    {
        // numeric, numeric-with-style, bool true/false, inlineStr,
        // formula-cached, error literal — the shapes the tokenizer knows.
        $xml = '<row r="2">'
            .'<c r="A2" t="n"><v>123.45</v></c>'
            .'<c r="B2" t="n" s="3"><v>45292</v></c>'
            .'<c r="C2" t="b"><v>1</v></c>'
            .'<c r="D2" t="b"><v>0</v></c>'
            .'<c r="E2" t="inlineStr"><is><t>hello</t></is></c>'
            .'<c r="F2" t="str"><v>cached</v></c>'
            .'<c r="G2" t="e"><v>#N/A</v></c>'
            .'</row>';
        $this->assertMatchesOracle($xml);
    }

    public function test_inline_string_rich_and_preserve_and_entities(): void
    {
        $xml = '<row r="5">'
            .'<c r="A5" t="inlineStr"><is><r><t>foo</t></r><r><t>bar</t></r></is></c>'
            .'<c r="B5" t="inlineStr"><is><t xml:space="preserve">  ws  </t></is></c>'
            .'<c r="C5" t="inlineStr"><is><t>a &amp; b &lt; c</t></is></c>'
            .'<c r="D5" t="n"><v>7</v></c>'
            .'</row>';
        $this->assertMatchesOracle($xml);
    }

    public function test_sparse_row_with_gaps(): void
    {
        // Cells at A, C, F — gaps at B, D, E must read as '' from both.
        $xml = '<row r="9">'
            .'<c r="A9" t="n"><v>10</v></c>'
            .'<c r="C9" t="inlineStr"><is><t>mid</t></is></c>'
            .'<c r="F9" t="n"><v>99</v></c>'
            .'</row>';
        $this->assertMatchesOracle($xml);
    }

    public function test_self_closing_and_empty_cells(): void
    {
        $xml = '<row r="3">'
            .'<c r="A3"/>'
            .'<c r="B3" t="n"><v>5</v></c>'
            .'<c r="C3"></c>'
            .'<c r="D3" t="inlineStr"><is><t></t></is></c>'
            .'</row>';
        $this->assertMatchesOracle($xml);
    }

    public function test_out_of_order_cells(): void
    {
        // An external writer may emit columns out of document order; the
        // extractor must still return the right value for a given index
        // (it may not early-stop on a forward jump). tokenizeRow rebuilds
        // densely, so it is the oracle.
        $xml = '<row r="4">'
            .'<c r="C4" t="n"><v>30</v></c>'
            .'<c r="A4" t="n"><v>10</v></c>'
            .'<c r="B4" t="inlineStr"><is><t>mid</t></is></c>'
            .'</row>';
        $this->assertMatchesOracle($xml);
    }

    public function test_cells_without_ref_use_positional_index(): void
    {
        // No r="" → tokenizeRow assigns the next position (maxIdx+1);
        // tokenizeColumn must mirror that exactly.
        $xml = '<row r="6">'
            .'<c t="n"><v>1</v></c>'
            .'<c t="n"><v>2</v></c>'
            .'<c t="inlineStr"><is><t>three</t></is></c>'
            .'</row>';
        $this->assertMatchesOracle($xml);
    }

    public function test_shared_string_resolution(): void
    {
        $sst = new InMemorySharedStrings(['alpha', 'beta', 'gamma']);
        $xml = '<row r="7">'
            .'<c r="A7" t="n"><v>0</v></c>'
            .'<c r="B7" t="s"><v>2</v></c>'
            .'<c r="C7" t="s"><v>0</v></c>'
            .'</row>';
        $this->assertMatchesOracle($xml, $sst);
    }

    public function test_empty_row(): void
    {
        $this->assertSame('', CellTokenizer::tokenizeColumn('<row r="8"></row>', 0));
        $this->assertSame('', CellTokenizer::tokenizeColumn('<row r="8"/>', 3));
    }
}

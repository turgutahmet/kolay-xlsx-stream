<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\CellTokenizer;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * Unit tests for the cell-level tokenizer. Operates on raw <row>...</row>
 * blobs without any I/O so it isolates parser correctness from streaming
 * concerns.
 */
class CellTokenizerTest extends TestCase
{
    public function test_inline_string_cell(): void
    {
        $row = '<row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row>';
        $this->assertSame(['hello'], CellTokenizer::tokenizeRow($row));
    }

    public function test_numeric_cell(): void
    {
        $row = '<row r="1"><c r="A1" t="n"><v>42.5</v></c></row>';
        $this->assertSame(['42.5'], CellTokenizer::tokenizeRow($row));
    }

    public function test_numeric_without_explicit_type(): void
    {
        $row = '<row r="1"><c r="A1"><v>123</v></c></row>';
        $this->assertSame(['123'], CellTokenizer::tokenizeRow($row));
    }

    public function test_boolean_cell_true(): void
    {
        $row = '<row r="1"><c r="A1" t="b"><v>1</v></c></row>';
        $this->assertSame([true], CellTokenizer::tokenizeRow($row));
    }

    public function test_boolean_cell_false(): void
    {
        $row = '<row r="1"><c r="A1" t="b"><v>0</v></c></row>';
        $this->assertSame([false], CellTokenizer::tokenizeRow($row));
    }

    public function test_self_closing_empty_cell(): void
    {
        $row = '<row r="1"><c r="A1"/><c r="B1" t="inlineStr"><is><t>x</t></is></c></row>';
        $this->assertSame(['', 'x'], CellTokenizer::tokenizeRow($row));
    }

    public function test_xml_entities_are_decoded(): void
    {
        $row = '<row r="1"><c r="A1" t="inlineStr"><is><t>foo &amp; &lt;bar&gt; &quot;qux&quot;</t></is></c></row>';
        $this->assertSame(['foo & <bar> "qux"'], CellTokenizer::tokenizeRow($row));
    }

    public function test_xml_space_preserve_keeps_whitespace(): void
    {
        $row = '<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">  spacey  </t></is></c></row>';
        $this->assertSame(['  spacey  '], CellTokenizer::tokenizeRow($row));
    }

    public function test_rich_text_runs_are_concatenated(): void
    {
        $row = '<row r="1"><c r="A1" t="inlineStr"><is><r><t>foo</t></r><r><t>bar</t></r></is></c></row>';
        $this->assertSame(['foobar'], CellTokenizer::tokenizeRow($row));
    }

    public function test_sparse_row_uses_r_attribute_authority(): void
    {
        // Cells B1 and D1 only — A1 and C1 missing entirely.
        $row = '<row r="1">'
            .'<c r="B1" t="inlineStr"><is><t>second</t></is></c>'
            .'<c r="D1" t="inlineStr"><is><t>fourth</t></is></c>'
            .'</row>';

        $this->assertSame(['', 'second', '', 'fourth'], CellTokenizer::tokenizeRow($row));
    }

    public function test_columns_resolve_letters_to_zero_indexed_position(): void
    {
        $this->assertSame(0, CellTokenizer::columnLettersToIndex('A'));
        $this->assertSame(25, CellTokenizer::columnLettersToIndex('Z'));
        $this->assertSame(26, CellTokenizer::columnLettersToIndex('AA'));
        $this->assertSame(27, CellTokenizer::columnLettersToIndex('AB'));
        $this->assertSame(701, CellTokenizer::columnLettersToIndex('ZZ'));
        $this->assertSame(702, CellTokenizer::columnLettersToIndex('AAA'));
    }

    public function test_unicode_passes_through_verbatim(): void
    {
        $row = '<row r="1"><c r="A1" t="inlineStr"><is><t>İstanbul 🌊 中文</t></is></c></row>';
        $this->assertSame(['İstanbul 🌊 中文'], CellTokenizer::tokenizeRow($row));
    }

    public function test_style_attribute_is_ignored(): void
    {
        // The "s" attribute identifies a style index — purely visual, no
        // value impact. Tokenizer should not let it interfere with parsing.
        $row = '<row r="1"><c r="A1" s="3" t="n"><v>45292</v></c></row>';
        $this->assertSame(['45292'], CellTokenizer::tokenizeRow($row));
    }

    public function test_empty_inline_string(): void
    {
        $row = '<row r="1"><c r="A1" t="inlineStr"><is><t></t></is></c></row>';
        $this->assertSame([''], CellTokenizer::tokenizeRow($row));
    }

    public function test_self_closing_t_inside_inline_string(): void
    {
        $row = '<row r="1"><c r="A1" t="inlineStr"><is><t/></is></c></row>';
        $this->assertSame([''], CellTokenizer::tokenizeRow($row));
    }

    public function test_multiple_cells_mixed_types(): void
    {
        $row = '<row r="1">'
            .'<c r="A1" t="n"><v>1</v></c>'
            .'<c r="B1" t="inlineStr"><is><t>two</t></is></c>'
            .'<c r="C1" t="b"><v>1</v></c>'
            .'<c r="D1"/>'
            .'<c r="E1" t="n"><v>5.5</v></c>'
            .'</row>';

        $this->assertSame(['1', 'two', true, '', '5.5'], CellTokenizer::tokenizeRow($row));
    }

    public function test_empty_row_returns_empty_array(): void
    {
        $this->assertSame([], CellTokenizer::tokenizeRow('<row r="1"/>'));
        $this->assertSame([], CellTokenizer::tokenizeRow('<row r="1"></row>'));
    }
}

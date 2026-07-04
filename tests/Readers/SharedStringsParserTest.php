<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\SharedStringsParser;
use Kolay\XlsxStream\Tests\TestCase;

class SharedStringsParserTest extends TestCase
{
    public function test_simple_string_entries(): void
    {
        $xml = $this->wrap(
            '<si><t>Hello</t></si>'.
            '<si><t>World</t></si>'
        );

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame(2, $sst->count());
        $this->assertSame('Hello', $sst->get(0));
        $this->assertSame('World', $sst->get(1));
    }

    public function test_xml_entities_decoded(): void
    {
        $xml = $this->wrap('<si><t>foo &amp; bar &lt;x&gt;</t></si>');

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame('foo & bar <x>', $sst->get(0));
    }

    public function test_xml_space_preserve_keeps_whitespace(): void
    {
        $xml = $this->wrap('<si><t xml:space="preserve">  spacey  </t></si>');

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame('  spacey  ', $sst->get(0));
    }

    public function test_rich_text_runs_concatenate(): void
    {
        $xml = $this->wrap(
            '<si>'.
                '<r><rPr><b/></rPr><t>Bold</t></r>'.
                '<r><t> </t></r>'.
                '<r><t>normal</t></r>'.
            '</si>'
        );

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame('Bold normal', $sst->get(0));
    }

    public function test_self_closing_si_yields_empty_string(): void
    {
        $xml = $this->wrap('<si/><si><t>after</t></si>');

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame(2, $sst->count());
        $this->assertSame('', $sst->get(0));
        $this->assertSame('after', $sst->get(1));
    }

    public function test_self_closing_t_inside_si_yields_empty(): void
    {
        $xml = $this->wrap('<si><t/></si>');

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame(1, $sst->count());
        $this->assertSame('', $sst->get(0));
    }

    public function test_unicode_passes_through(): void
    {
        $xml = $this->wrap('<si><t>İstanbul 🌊 中文</t></si>');

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame('İstanbul 🌊 中文', $sst->get(0));
    }

    public function test_empty_table(): void
    {
        $xml = '<?xml version="1.0"?><sst xmlns="..." count="0" uniqueCount="0"></sst>';

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame(0, $sst->count());
    }

    public function test_index_order_matches_document_order(): void
    {
        $xml = $this->wrap(
            '<si><t>zero</t></si>'.
            '<si><t>one</t></si>'.
            '<si><t>two</t></si>'.
            '<si><t>three</t></si>'
        );

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame('zero', $sst->get(0));
        $this->assertSame('one', $sst->get(1));
        $this->assertSame('two', $sst->get(2));
        $this->assertSame('three', $sst->get(3));
    }

    public function test_does_not_match_si_prefixes_in_other_tags(): void
    {
        $xml = $this->wrap(
            '<sidetable>noise</sidetable>'.
            '<si><t>real</t></si>'
        );

        $sst = SharedStringsParser::parseInMemory($xml);

        $this->assertSame(1, $sst->count());
        $this->assertSame('real', $sst->get(0));
    }

    public function test_out_of_range_index_throws(): void
    {
        $sst = SharedStringsParser::parseInMemory($this->wrap('<si><t>only</t></si>'));

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxReadException::class);
        $this->expectExceptionMessageMatches('/shared-string index 1 out of range/');

        $sst->get(1);
    }

    public function test_negative_index_throws(): void
    {
        $sst = SharedStringsParser::parseInMemory($this->wrap('<si><t>only</t></si>'));

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxReadException::class);
        $this->expectExceptionMessageMatches('/shared-string index -1 out of range/');

        $sst->get(-1);
    }

    public function test_streaming_chunks_split_at_every_byte_boundary(): void
    {
        // The whole point of the incremental parser: an <si> entry (or
        // the '<si' needle itself, or an entity, or a quoted attribute)
        // straddling a chunk edge must parse identically to the
        // one-shot path. Exhaustively split a document that exercises
        // every recognised shape at EVERY byte position.
        $xml = $this->wrap(
            '<si><t>plain</t></si>'.
            '<si/>'.
            '<si><t xml:space="preserve">  ws  </t></si>'.
            '<si><r><rPr><b/></rPr><t>Bold</t></r><r><t> tail</t></r></si>'.
            '<sidetable>noise</sidetable>'.
            '<si><t>a &amp; b</t></si>'.
            '<si><t>İstanbul 🌊</t></si>'
        );
        $expected = ['plain', '', '  ws  ', 'Bold tail', 'a & b', 'İstanbul 🌊'];

        $len = strlen($xml);
        for ($split = 1; $split < $len; $split++) {
            $parser = new SharedStringsParser();
            $parser->push(substr($xml, 0, $split));
            $parser->push(substr($xml, $split));
            $sst = $parser->finish();

            $this->assertSame(count($expected), $sst->count(), "split at byte {$split}");
            foreach ($expected as $i => $value) {
                $this->assertSame($value, $sst->get($i), "entry {$i}, split at byte {$split}");
            }
        }
    }

    public function test_streaming_single_byte_chunks(): void
    {
        // Degenerate chunking — every push carries one byte, so every
        // carry path (partial '<si', unfinished opening tag, unfinished
        // body) runs on every entry.
        $xml = $this->wrap('<si><t>one</t></si><si/><si><t>two &lt;3</t></si>');

        $parser = new SharedStringsParser();
        foreach (str_split($xml) as $byte) {
            $parser->push($byte);
        }
        $sst = $parser->finish();

        $this->assertSame(3, $sst->count());
        $this->assertSame('one', $sst->get(0));
        $this->assertSame('', $sst->get(1));
        $this->assertSame('two <3', $sst->get(2));
    }

    public function test_truncated_trailing_entry_is_dropped(): void
    {
        // EOF in the middle of an entry — the incomplete tail parses to
        // nothing, matching the pre-streaming parser's break-on-missing
        // '</si>' behaviour.
        $parser = new SharedStringsParser();
        $parser->push('<sst><si><t>whole</t></si><si><t>cut off he');
        $sst = $parser->finish();

        $this->assertSame(1, $sst->count());
        $this->assertSame('whole', $sst->get(0));
    }

    public function test_unclosed_si_entry_exceeding_carry_cap_throws(): void
    {
        // An <si> that never closes would otherwise accumulate the whole
        // remaining document in the carry buffer — the 16 MB cap keeps
        // the streaming parse's memory bound honest on malicious input.
        $parser = new SharedStringsParser();
        $parser->push('<sst><si><t>');

        $this->expectException(\Kolay\XlsxStream\Exceptions\XlsxReadException::class);
        $this->expectExceptionMessageMatches('/<si> entry exceeds 16 MB/');

        // 17 MB of body that never reaches '</si>'.
        $filler = str_repeat('x', 1024 * 1024);
        for ($i = 0; $i < 17; $i++) {
            $parser->push($filler);
        }
    }

    private function wrap(string $body): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            $body.
            '</sst>';
    }
}

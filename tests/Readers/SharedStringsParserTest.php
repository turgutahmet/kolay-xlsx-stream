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

    private function wrap(string $body): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            $body.
            '</sst>';
    }
}

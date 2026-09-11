<?php

namespace Kolay\XlsxStream\Tests\Templates;

use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\SharedStringTable;

/**
 * The shared-string table: OOXML's canonical text path (Excel writes it),
 * and the one place template mode grows memory.
 *
 * Two hard rules. Indices are POSITIONAL, and the template's header cells
 * already point at them, so seeded entries keep their index and are copied
 * verbatim — a rich-text <si> is never reinterpreted. And the dictionary is
 * capped: past maxUnique, intern() declines and the caller falls back to an
 * inline string, so a column of a million distinct values cannot turn the
 * writer's bounded memory into an unbounded dictionary.
 *
 * intern() takes text the writer has ALREADY escaped, so the package keeps a
 * single escaper (BaseXlsxWriter::fastXmlEscape) rather than a second copy
 * that could drift.
 */
class SharedStringTableTest extends TestCase
{
    private function seedTable(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="3">'
            .'<si><t>Employee</t></si>'
            .'<si><t>Leave days</t></si>'
            .'<si><r><rPr><b/></rPr><t>Rich</t></r><r><t> header</t></r></si>'
            .'</sst>';
    }

    public function test_seeded_indices_are_preserved_and_new_text_appends(): void
    {
        $sst = SharedStringTable::fromXml($this->seedTable());

        $this->assertSame(3, $sst->uniqueCount(), 'three <si> in the seed');
        // Header cells already reference 0..2 — ours must start at 3.
        $this->assertSame(3, $sst->intern('Ahmet'));
        $this->assertSame(4, $sst->intern('Mehmet'));
        $this->assertSame(3, $sst->intern('Ahmet'), 'repeat resolves to the same index');
    }

    public function test_no_seed_starts_at_zero_and_dedupes(): void
    {
        $sst = SharedStringTable::fromXml(null);

        $this->assertSame(0, $sst->intern('a'));
        $this->assertSame(1, $sst->intern('b'));
        $this->assertSame(0, $sst->intern('a'));
        $this->assertSame(2, $sst->uniqueCount());
        $this->assertSame(3, $sst->referenceCount(), 'every hit is a reference');
    }

    public function test_untouched_seed_is_returned_byte_identical(): void
    {
        $xml = $this->seedTable();
        $this->assertSame($xml, SharedStringTable::fromXml($xml)->toXml());
    }

    public function test_appending_updates_both_counts_and_keeps_the_seed_verbatim(): void
    {
        $sst = SharedStringTable::fromXml($this->seedTable());
        $sst->intern('Ahmet');
        $sst->intern('Ahmet');
        $out = $sst->toXml();

        // seed count 4 + 2 new references; uniqueCount 3 + 1.
        $this->assertStringContainsString('count="6"', $out);
        $this->assertStringContainsString('uniqueCount="4"', $out);
        // The rich-text entry survives untouched — never reinterpreted.
        $this->assertStringContainsString('<si><r><rPr><b/></rPr><t>Rich</t></r><r><t> header</t></r></si>', $out);
        // Ours is appended last.
        $this->assertStringContainsString('<si><t>Ahmet</t></si></sst>', $out);
    }

    public function test_generates_a_valid_table_when_there_is_no_seed(): void
    {
        $sst = SharedStringTable::fromXml(null);
        $sst->intern('solo');
        $out = $sst->toXml();

        $this->assertStringStartsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst ', $out);
        $this->assertStringContainsString('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', $out);
        $this->assertStringContainsString('count="1"', $out);
        $this->assertStringContainsString('uniqueCount="1"', $out);
        $this->assertStringEndsWith('<si><t>solo</t></si></sst>', $out);
    }

    public function test_whitespace_edges_get_xml_space_preserve(): void
    {
        $sst = SharedStringTable::fromXml(null);
        $sst->intern('  padded  ');
        $sst->intern('tight');
        $out = $sst->toXml();

        $this->assertStringContainsString('<si><t xml:space="preserve">  padded  </t></si>', $out);
        $this->assertStringContainsString('<si><t>tight</t></si>', $out);
    }

    public function test_saturation_declines_new_text_but_keeps_serving_known_text(): void
    {
        $sst = SharedStringTable::fromXml(null, maxUnique: 2);

        $this->assertSame(0, $sst->intern('a'));
        $this->assertSame(1, $sst->intern('b'));
        $this->assertFalse($sst->isSaturated());

        // The dictionary is full: a NEW string is declined so the caller can
        // fall back to an inline string, keeping memory bounded.
        $this->assertNull($sst->intern('c'));
        $this->assertTrue($sst->isSaturated());
        // Known strings still resolve — no cliff for the common values.
        $this->assertSame(0, $sst->intern('a'));
        $this->assertSame(2, $sst->uniqueCount(), 'declined text never enters the table');
    }

    public function test_seed_without_count_attributes_is_still_counted(): void
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<si><t>a</t></si><si><t>b</t></si></sst>';
        $sst = SharedStringTable::fromXml($xml);

        // Indices come from the actual <si> elements, not a trusted attribute.
        $this->assertSame(2, $sst->uniqueCount());
        $this->assertSame(2, $sst->intern('c'));
        $this->assertStringContainsString('uniqueCount="3"', $sst->toXml());
    }
}

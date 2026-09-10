<?php

namespace Kolay\XlsxStream\Tests\Templates;

use Kolay\XlsxStream\Styles\StyleRegistry;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * Seeding the style registry from a template's styles.xml.
 *
 * The template's stylesheet is the AUTHORITY: it carries borders, cell
 * styles, dxfs and table styles this package does not model, so the seeded
 * registry never regenerates it. It appends — new fonts, fills and xfs go
 * at the END of the existing tables and the count attributes are bumped —
 * exactly the "open one seam, copy the rest" rule the sheet cut follows.
 *
 * The dominant case registers nothing at all, and then styles.xml must come
 * out byte-for-byte as it went in.
 */
class TemplateStyleSeedTest extends TestCase
{
    /** A realistic foreign stylesheet: blocks we model and blocks we do not. */
    private function templateStyles(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            .'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            .'<numFmts count="2">'
            .'<numFmt numFmtId="165" formatCode="#,##0.00\\ &quot;TL&quot;"/>'
            .'<numFmt numFmtId="170" formatCode="dd/mm/yyyy"/>'
            .'</numFmts>'
            .'<fonts count="2" x14ac:knownFonts="1">'
            .'<font><sz val="11"/><name val="Calibri"/></font>'
            .'<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
            .'</fonts>'
            .'<fills count="3">'
            .'<fill><patternFill patternType="none"/></fill>'
            .'<fill><patternFill patternType="gray125"/></fill>'
            .'<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>'
            .'</fills>'
            .'<borders count="2">'
            .'<border><left/><right/><top/><bottom/><diagonal/></border>'
            .'<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>'
            .'</borders>'
            .'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            .'<cellXfs count="4">'
            .'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            .'<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1"/>'
            .'<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>'
            .'<xf numFmtId="170" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>'
            .'</cellXfs>'
            .'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            .'<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'
            .'<tableStyles count="0" defaultTableStyle="TableStyleMedium2"/>'
            .'<colors><indexedColors><rgbColor rgb="FF000000"/></indexedColors></colors>'
            .'</styleSheet>';
    }

    public function test_no_registration_returns_the_template_stylesheet_byte_identical(): void
    {
        $xml = $this->templateStyles();
        $registry = StyleRegistry::fromStylesXml($xml);

        // The overwhelmingly common case: the caller styles nothing, so the
        // template's stylesheet must survive untouched.
        $this->assertSame($xml, $registry->toXml());
    }

    public function test_new_style_ids_continue_after_the_template_tables(): void
    {
        $registry = StyleRegistry::fromStylesXml($this->templateStyles());

        // The template already has 4 cellXfs (ids 0..3); ours starts at 4.
        $id = $registry->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
        $this->assertSame(4, $id);
        $this->assertSame(5, $registry->registerRowStyle(['bold' => true]));
    }

    public function test_registration_appends_and_bumps_counts(): void
    {
        $registry = StyleRegistry::fromStylesXml($this->templateStyles());
        $registry->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
        $out = $registry->toXml();

        // Counts follow the appends: fonts 2→3, fills 3→4, cellXfs 4→5.
        $this->assertStringContainsString('<fonts count="3"', $out);
        $this->assertStringContainsString('<fills count="4"', $out);
        $this->assertStringContainsString('<cellXfs count="5"', $out);

        // The template's own entries are still first and unchanged.
        $this->assertStringContainsString('<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>', $out);
        $this->assertStringContainsString('<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>', $out);

        // Ours is appended at the end of cellXfs, referencing the appended
        // font (2) and fill (3).
        $this->assertMatchesRegularExpression(
            '/<xf numFmtId="0" fontId="2" fillId="3"[^>]*\/><\/cellXfs>/',
            $out
        );
        $this->assertStringContainsString('FF9C0006', $out, 'appended font colour');
        $this->assertStringContainsString('FFFFC7CE', $out, 'appended fill colour');
    }

    public function test_blocks_the_package_does_not_model_survive(): void
    {
        $registry = StyleRegistry::fromStylesXml($this->templateStyles());
        $registry->registerRowStyle(['bold' => true]);
        $out = $registry->toXml();

        foreach ([
            '<borders count="2">',
            '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>',
            '<cellStyleXfs count="1">',
            '<cellStyles count="1">',
            '<dxfs count="1">',
            '<tableStyles count="0" defaultTableStyle="TableStyleMedium2"/>',
            '<colors><indexedColors>',
        ] as $block) {
            $this->assertStringContainsString($block, $out, "seed block lost: {$block}");
        }
    }

    public function test_numfmt_ids_continue_after_the_templates_highest(): void
    {
        $registry = StyleRegistry::fromStylesXml($this->templateStyles());

        // Template's highest custom id is 170 — ours must not collide with it
        // (and must not restart at the classic 164).
        $id = $registry->registerColumnFormat('0.000');
        $out = $registry->toXml();

        $this->assertSame(4, $id, 'style id continues the cellXfs table');
        $this->assertStringContainsString('<numFmt numFmtId="171" formatCode="0.000"/>', $out);
        $this->assertStringContainsString('<numFmts count="3">', $out);
        $this->assertStringNotContainsString('numFmtId="164"', $out);
    }

    public function test_seeding_a_stylesheet_without_numfmts_inserts_the_block(): void
    {
        $xml = str_replace(
            '<numFmts count="2"><numFmt numFmtId="165" formatCode="#,##0.00\\ &quot;TL&quot;"/><numFmt numFmtId="170" formatCode="dd/mm/yyyy"/></numFmts>',
            '',
            $this->templateStyles()
        );
        $registry = StyleRegistry::fromStylesXml($xml);
        $this->assertSame($xml, $registry->toXml(), 'untouched when nothing is registered');

        $registry->registerColumnFormat('0.000');
        $out = $registry->toXml();
        // numFmts must be created, and it must come first inside styleSheet.
        $this->assertMatchesRegularExpression('/<styleSheet[^>]*><numFmts count="1">/', $out);
        // A stylesheet with no custom formats starts our ids at the 164 floor.
        $this->assertStringContainsString('<numFmt numFmtId="164" formatCode="0.000"/>', $out);
    }

    public function test_classic_registry_is_untouched_by_the_seeding_feature(): void
    {
        // Byte-oracle for the non-template path: an unseeded registry must
        // still produce exactly the stylesheet it always has.
        $classic = new StyleRegistry();
        $xml = $classic->toXml();

        $this->assertStringStartsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet', $xml);
        $this->assertStringContainsString('<cellXfs count="2">', $xml);
        $this->assertStringContainsString('<borders count="1">', $xml);
        $this->assertStringContainsString('numFmtId="164"', $xml);
    }

    public function test_seed_tables_are_sized_from_elements_not_the_count_attribute(): void
    {
        // `count` is optional in the schema, and a producer may leave it
        // stale. Trusting it would put our first appended xf at index 0,
        // colliding with the template's own xf 0 — silently wrong styling
        // rather than an error. Sizes come from the elements themselves.
        $xml = str_replace(
            ['<fonts count="2" x14ac:knownFonts="1">', '<fills count="3">', '<cellXfs count="4">'],
            ['<fonts x14ac:knownFonts="1">', '<fills count="99">', '<cellXfs>'],
            $this->templateStyles()
        );

        $registry = StyleRegistry::fromStylesXml($xml);
        $this->assertSame($xml, $registry->toXml(), 'still byte-identical when nothing is registered');

        // 4 real <xf> in cellXfs → ours is 4, NOT 0.
        $id = $registry->registerRowStyle(['fill' => '#FFC7CE', 'color' => '#9C0006']);
        $this->assertSame(4, $id);

        $out = $registry->toXml();
        // Missing counts are created; the stale 99 is corrected to the truth.
        $this->assertStringContainsString('<cellXfs count="5">', $out);
        $this->assertStringContainsString('count="4"', $out, 'fills 3 real + 1 appended');
        $this->assertMatchesRegularExpression('/<fonts[^>]*count="3"/', $out);
        // The appended xf references the appended font/fill by true position.
        $this->assertMatchesRegularExpression('/<xf numFmtId="0" fontId="2" fillId="3"[^>]*\/><\/cellXfs>/', $out);
    }

    public function test_child_counting_is_scoped_to_its_own_block(): void
    {
        // <fill> also lives inside <dxf>, and <xf> inside <cellStyleXfs>;
        // counting document-wide would inflate both offsets.
        $registry = StyleRegistry::fromStylesXml($this->templateStyles());

        // cellStyleXfs holds one <xf> and dxfs one <fill> — neither may count.
        $this->assertSame(4, $registry->registerRowStyle(['bold' => true]), 'cellStyleXfs xf not counted');
        $registry2 = StyleRegistry::fromStylesXml($this->templateStyles());
        $registry2->registerRowStyle(['fill' => '#123456']);
        $this->assertStringContainsString('fillId="3"', $registry2->toXml(), 'dxf fill not counted');
    }
}

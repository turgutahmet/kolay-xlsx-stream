<?php

namespace Kolay\XlsxStream\Tests\Readers;

use Kolay\XlsxStream\Readers\StylesParser;
use Kolay\XlsxStream\Styles\StyleRegistry;
use Kolay\XlsxStream\Tests\TestCase;
use PHPUnit\Framework\Attributes\DataProvider;

class StylesParserTest extends TestCase
{
    /**
     * The format-code date scanner: d/m/y/h/s tokens count only OUTSIDE
     * quoted sections, bracket sections, escapes and _/* layout chars.
     */
    #[DataProvider('formatCodeProvider')]
    public function test_format_code_date_detection(string $code, bool $isDate): void
    {
        $this->assertSame(
            $isDate,
            StylesParser::formatCodeIsDate($code),
            "format code '{$code}' should".($isDate ? '' : ' NOT').' be a date'
        );
    }

    /**
     * @return array<string, array{string, bool}>
     */
    public static function formatCodeProvider(): array
    {
        return [
            // — dates —
            'iso date' => ['yyyy-mm-dd', true],
            'uppercase tokens' => ['YYYY-MM-DD', true],
            'builtin 14 shape' => ['m/d/yyyy', true],
            'month name' => ['d-mmm-yy', true],
            'time' => ['mm:ss', true],
            'datetime with quoted T' => ['yyyy-mm-dd"T"hh:mm:ss', true],
            'am/pm time' => ['h:mm AM/PM', true],
            'locale prefix then date' => ['[$-409]m/d/yyyy', true],
            'color prefix then date' => ['[Red]yyyy-mm-dd', true],
            // — not dates —
            'decimal' => ['#,##0.00', false],
            'quoted m literal' => ['0.00" m"', false],
            'color only' => ['[Red]0.00', false],
            'general' => ['General', false],
            'text placeholder' => ['@', false],
            'scientific E is not a date token' => ['0.00E+00', false],
            'quoted words with slash' => ['"yes/no"', false],
            'percent' => ['0.00%', false],
            'escaped m' => ['0.00\\m', false],
            'underscore layout m' => ['0.00_m', false],
            'asterisk fill m' => ['0.00*m', false],
            'currency with escaped space' => ['#,##0.00\\ ₺', false],
            'unterminated quote is conservative' => ['0.00"m', false],
            'unterminated bracket is conservative' => ['[Redm', false],
            'empty' => ['', false],
        ];
    }

    public function test_bitmap_indexes_cellxfs_and_ignores_cellstylexfs(): void
    {
        // PhpSpreadsheet-shaped stylesheet: a <cellStyleXfs> whose xf
        // carries a DATE numFmtId (it must NOT leak into the bitmap —
        // cells never address that table) followed by the real cellXfs.
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.
            '<numFmts count="2">'.
            '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>'.
            '<numFmt numFmtId="165" formatCode="#,##0.00"/>'.
            '</numFmts>'.
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'.
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'.
            '<borders count="1"><border/></borders>'.
            '<cellStyleXfs count="1"><xf numFmtId="14" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'.
            '<cellXfs count="6">'.
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'.
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'.
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'.
            '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'.
            '<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'.
            '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'.
            '</cellXfs>'.
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'.
            '</styleSheet>';

        $this->assertSame([
            false, // 0: General
            true,  // 1: custom yyyy-mm-dd
            false, // 2: custom #,##0.00
            true,  // 3: builtin 14 (m/d/yyyy)
            true,  // 4: builtin 22 (m/d/yyyy h:mm)
            false, // 5: builtin 2 (0.00)
        ], StylesParser::dateStyleBitmap($xml));
    }

    public function test_builtin_id_families(): void
    {
        // 14-22 dates/times, 45-47 elapsed; neighbours on either side
        // (13 fraction, 23 unassigned, 44 accounting, 48 scientific,
        // 49 text) are not dates.
        $xfs = '';
        $ids = [13, 14, 22, 23, 44, 45, 47, 48, 49];
        foreach ($ids as $id) {
            $xfs .= '<xf numFmtId="'.$id.'"/>';
        }
        $xml = '<styleSheet><cellXfs count="'.count($ids).'">'.$xfs.'</cellXfs></styleSheet>';

        $this->assertSame(
            [false, true, true, false, false, true, true, false, false],
            StylesParser::dateStyleBitmap($xml)
        );
    }

    public function test_xf_without_numfmtid_defaults_to_general(): void
    {
        $xml = '<styleSheet><cellXfs count="2">'.
            '<xf fontId="0" fillId="0"/>'.
            '<xf numFmtId="14"/>'.
            '</cellXfs></styleSheet>';

        $this->assertSame([false, true], StylesParser::dateStyleBitmap($xml));
    }

    public function test_missing_cellxfs_yields_empty_bitmap(): void
    {
        $this->assertSame([], StylesParser::dateStyleBitmap('<styleSheet><fonts count="0"/></styleSheet>'));
    }

    public function test_entity_encoded_quotes_in_format_code_are_honoured(): void
    {
        // formatCode='0.00" m"' arrives XML-encoded as 0.00&quot; m&quot;
        // — decoding must happen before the token scan or the quoted m
        // would incorrectly flag the style as a date.
        $xml = '<styleSheet>'.
            '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00&quot; m&quot;"/></numFmts>'.
            '<cellXfs count="1"><xf numFmtId="164"/></cellXfs>'.
            '</styleSheet>';

        $this->assertSame([false], StylesParser::dateStyleBitmap($xml));
    }

    public function test_parses_this_packages_writer_stylesheet(): void
    {
        // Cross-check against the real thing: StyleRegistry is exactly
        // what SinkableXlsxWriter emits. Index 1 is the historical
        // datetime style (numFmtId 164 = 'yyyy-mm-dd hh:mm:ss'), and
        // registered date/decimal presets land where the registry says.
        $registry = new StyleRegistry();
        $dateId = $registry->registerColumnFormat('date');       // yyyy-mm-dd
        $decimalId = $registry->registerColumnFormat('decimal'); // #,##0.00
        $builtinId = $registry->registerBuiltinNumFmt(StyleRegistry::BUILTIN_NUMFMT_DATE);

        $bitmap = StylesParser::dateStyleBitmap($registry->toXml());

        $this->assertFalse($bitmap[0], 'default style must not be a date');
        $this->assertTrue($bitmap[1], 'legacy datetime style (numFmtId 164) must be a date');
        $this->assertTrue($bitmap[$dateId]);
        $this->assertFalse($bitmap[$decimalId]);
        $this->assertTrue($bitmap[$builtinId]);
    }
}

<?php

namespace Kolay\XlsxStream\Readers;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxReadException;

/**
 * @internal
 *
 * Resolves the workbook's sheet table into ZIP entry paths.
 *
 * XLSX stores the sheet listing in two files that must be cross-referenced:
 *
 *   xl/workbook.xml
 *     <sheets>
 *       <sheet name="Sheet1" sheetId="1" r:id="rId3"/>
 *       <sheet name="Reports" sheetId="2" r:id="rId4"/>
 *     </sheets>
 *
 *   xl/_rels/workbook.xml.rels
 *     <Relationship Id="rId3" Target="worksheets/sheet1.xml" Type="..."/>
 *     <Relationship Id="rId4" Target="worksheets/sheet2.xml" Type="..."/>
 *
 * The r:id link gives the entry path; "Target" is relative to "xl/" (the
 * directory containing workbook.xml) but tools sometimes write absolute
 * paths starting with "/" — both forms are handled.
 *
 * Returns sheets in workbook.xml order, which is the user-visible tab
 * order. Sheet IDs are not always sequential and not always equal to the
 * tab order, so callers MUST address sheets by name or by ordinal index
 * here, not by sheetId.
 */
class WorkbookResolver
{
    private const NS_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';

    /**
     * @return list<array{name: string, sheetId: int, entry: string}>
     */
    public static function resolve(Source $source, ZipDirectory $cd): array
    {
        if (! $cd->has('xl/workbook.xml')) {
            throw XlsxReadException::corruptCentralDirectory('xl/workbook.xml is missing');
        }
        if (! $cd->has('xl/_rels/workbook.xml.rels')) {
            throw XlsxReadException::corruptCentralDirectory('xl/_rels/workbook.xml.rels is missing');
        }

        $workbookXml = $cd->readEntry($source, 'xl/workbook.xml');
        $relsXml = $cd->readEntry($source, 'xl/_rels/workbook.xml.rels');

        $sheets = self::parseSheets($workbookXml);
        $rels = self::parseRels($relsXml);

        $resolved = [];
        foreach ($sheets as $sheet) {
            $rId = $sheet['rId'];
            if (! isset($rels[$rId])) {
                throw XlsxReadException::corruptCentralDirectory(
                    "sheet '{$sheet['name']}' references unknown relationship id '{$rId}'"
                );
            }
            $entry = self::normaliseTarget($rels[$rId]);
            if (! $cd->has($entry)) {
                throw XlsxReadException::corruptCentralDirectory(
                    "sheet '{$sheet['name']}' resolves to '{$entry}' which is not in the archive"
                );
            }
            $resolved[] = [
                'name' => $sheet['name'],
                'sheetId' => $sheet['sheetId'],
                'entry' => $entry,
            ];
        }

        return $resolved;
    }

    /**
     * @return list<array{name: string, sheetId: int, rId: string}>
     */
    private static function parseSheets(string $xml): array
    {
        $doc = self::loadXml($xml, 'xl/workbook.xml');

        $sheets = [];
        foreach ($doc->getElementsByTagName('sheet') as $node) {
            if (! $node instanceof \DOMElement) {
                continue;
            }
            $name = $node->getAttribute('name');
            $sheetId = (int) $node->getAttribute('sheetId');
            $rId = $node->getAttributeNS(
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'id'
            );
            // Some writers use the unprefixed "id" attribute; fall back.
            if ($rId === '') {
                $rId = $node->getAttribute('r:id');
            }
            if ($name === '' || $rId === '') {
                continue;
            }
            $sheets[] = [
                'name' => $name,
                'sheetId' => $sheetId,
                'rId' => $rId,
            ];
        }

        return $sheets;
    }

    /**
     * @return array<string, string> rId => Target
     */
    private static function parseRels(string $xml): array
    {
        $doc = self::loadXml($xml, 'xl/_rels/workbook.xml.rels');

        $rels = [];
        foreach ($doc->getElementsByTagNameNS(self::NS_RELATIONSHIPS, 'Relationship') as $node) {
            if (! $node instanceof \DOMElement) {
                continue;
            }
            $id = $node->getAttribute('Id');
            $target = $node->getAttribute('Target');
            if ($id === '' || $target === '') {
                continue;
            }
            $rels[$id] = $target;
        }

        // Fallback for documents that don't declare the relationships namespace
        // (rare but seen in third-party XLSX writers).
        if ($rels === []) {
            foreach ($doc->getElementsByTagName('Relationship') as $node) {
                if (! $node instanceof \DOMElement) {
                    continue;
                }
                $id = $node->getAttribute('Id');
                $target = $node->getAttribute('Target');
                if ($id !== '' && $target !== '') {
                    $rels[$id] = $target;
                }
            }
        }

        return $rels;
    }

    /**
     * Resolve a Relationship Target into a full ZIP entry path.
     *
     * Relative targets are anchored to "xl/" (the directory containing
     * workbook.xml). Absolute targets (starting with "/") are taken as-is
     * with the leading slash stripped — this matches the OPC convention.
     */
    private static function normaliseTarget(string $target): string
    {
        if ($target === '') {
            return '';
        }
        if ($target[0] === '/') {
            return ltrim($target, '/');
        }

        return 'xl/'.ltrim($target, './');
    }

    private static function loadXml(string $xml, string $name): \DOMDocument
    {
        $doc = new \DOMDocument();
        $previous = libxml_use_internal_errors(true);
        $loaded = $doc->loadXML($xml, LIBXML_NONET);
        $errors = libxml_get_errors();
        libxml_clear_errors();
        libxml_use_internal_errors($previous);

        if (! $loaded) {
            $first = $errors[0]->message ?? 'unknown XML parse error';
            throw XlsxReadException::corruptCentralDirectory("could not parse {$name}: ".trim($first));
        }

        return $doc;
    }
}

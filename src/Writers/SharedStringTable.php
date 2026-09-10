<?php

namespace Kolay\XlsxStream\Writers;

/**
 * The workbook's shared-string table — OOXML's canonical text path, and the
 * one structure in template mode that grows with the data.
 *
 * Cells written as `t="s"` carry an INDEX into this table instead of their
 * own text. That is what Excel itself emits, it shrinks a file whose columns
 * repeat (statuses, regions, names), and every reader resolves it to a plain
 * string — unlike an inline string, which PhpSpreadsheet hands back as a
 * RichText object. The classic writer keeps `inlineStr` as its default (an
 * O(1)-memory choice); this table is the opt-in.
 *
 * **Seeded indices are immutable.** A template's header cells already point
 * at its `<si>` entries by position, so the seed's entries keep their indices
 * and are copied VERBATIM — a rich-text `<si>` with `<r>`/`<rPr>` runs is
 * never reinterpreted, the same "do not read what you only need to copy"
 * rule the sheet seam and the style table follow. New text is appended after
 * them and is not de-duplicated against the seed (that would mean parsing it);
 * in practice a template's table holds only headers, so the overlap is nil.
 *
 * **The dictionary is capped.** Past `maxUnique`, `intern()` returns null and
 * the caller writes an inline string instead. Bounded memory is the package's
 * promise; a column of a million distinct values must not quietly turn it
 * into a million-entry dictionary.
 *
 * `intern()` takes text the writer has ALREADY escaped, so escaping lives in
 * exactly one place (BaseXlsxWriter::fastXmlEscape) instead of a second copy
 * here that could drift out of step.
 */
final class SharedStringTable
{
    /**
     * Default dictionary ceiling. Roughly 500k short strings — tens of MB,
     * far above any real report's distinct-text count and far below the point
     * where a runaway column would exhaust a queue worker.
     */
    public const DEFAULT_MAX_UNIQUE = 500_000;

    private const NAMESPACE_URI = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    /**
     * Escaped text => shared-string index, in insertion order.
     *
     * This single map is also the render order: PHP preserves insertion
     * order, so toXml() walks the keys instead of keeping a parallel list.
     * Measured at 200k unique strings, the parallel list cost 1.40x (not the
     * 2x one might assume — the key and the value share one refcounted
     * zend_string, so the overhead is the second array's buckets, ~20 bytes
     * an entry). Keys that look numeric come back losslessly through a
     * (string) cast, verified for "123", "0123", "1.5", " 7", "+8", "007".
     *
     * @var array<string, int>
     */
    private array $index = [];

    private bool $saturated = false;

    private function __construct(
        private ?string $seedXml,
        private int $seedUnique,
        private int $seedReferences,
        private int $maxUnique,
        private int $references = 0,
    ) {
    }

    /**
     * Build a table, optionally seeded from a template's sharedStrings.xml.
     * Pass null when the template has none (or for the classic writer).
     */
    public static function fromXml(?string $xml, int $maxUnique = self::DEFAULT_MAX_UNIQUE): self
    {
        if ($xml === null || trim($xml) === '') {
            return new self(null, 0, 0, max(0, $maxUnique));
        }

        return new self(
            $xml,
            self::countEntries(self::extractBody($xml)),
            self::attributeValue($xml, 'count'),
            max(0, $maxUnique),
        );
    }

    /**
     * Resolve escaped text to its shared-string index, adding it if new.
     * Returns null when the dictionary is full and the text is not already
     * in it — the caller then writes the cell as an inline string.
     */
    public function intern(string $escapedText): ?int
    {
        if (isset($this->index[$escapedText])) {
            $this->references++;

            return $this->index[$escapedText];
        }

        if ($this->uniqueCount() >= $this->maxUnique) {
            $this->saturated = true;

            return null;
        }

        $position = $this->seedUnique + \count($this->index);
        $this->index[$escapedText] = $position;
        $this->references++;

        return $position;
    }

    /** True once the ceiling has turned at least one string away. */
    public function isSaturated(): bool
    {
        return $this->saturated;
    }

    /** Distinct entries the table will write, seed included. */
    public function uniqueCount(): int
    {
        return $this->seedUnique + \count($this->index);
    }

    /** Total `t="s"` references, seed's own header cells included. */
    public function referenceCount(): int
    {
        return $this->seedReferences + $this->references;
    }

    /**
     * True when the table is an untouched template seed, so the template's
     * own sharedStrings.xml can be moved into the output rather than
     * re-rendered from its bytes.
     */
    public function isSeedUnmodified(): bool
    {
        return $this->seedXml !== null && $this->index === [];
    }

    public function isEmpty(): bool
    {
        return $this->uniqueCount() === 0;
    }

    /**
     * The table as sharedStrings.xml. An untouched seed comes back
     * byte-for-byte; otherwise the seed's own bytes are kept and ours are
     * spliced in before `</sst>` with both counts refreshed.
     */
    public function toXml(): string
    {
        if ($this->index === [] && $this->seedXml !== null) {
            return $this->seedXml;
        }

        $fragments = '';
        foreach ($this->index as $text => $_) {
            $fragments .= '<si>'.self::renderText((string) $text).'</si>';
        }

        if ($this->seedXml === null) {
            return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                .'<sst xmlns="'.self::NAMESPACE_URI.'"'
                .' count="'.$this->referenceCount().'"'
                .' uniqueCount="'.$this->uniqueCount().'">'
                .$fragments
                .'</sst>';
        }

        $xml = self::withCount($this->seedXml, 'count', $this->referenceCount());
        $xml = self::withCount($xml, 'uniqueCount', $this->uniqueCount());

        $close = strrpos($xml, '</sst');
        if ($close === false) {
            return $xml;
        }

        return substr($xml, 0, $close).$fragments.substr($xml, $close);
    }

    /**
     * `<t>` with the writer's own whitespace rule: only a leading or
     * trailing space/tab needs xml:space, and paying for it unconditionally
     * would bloat every entry.
     */
    private static function renderText(string $escaped): string
    {
        $last = $escaped === '' ? '' : $escaped[\strlen($escaped) - 1];
        $first = $escaped === '' ? '' : $escaped[0];

        if ($first === ' ' || $first === "\t" || $last === ' ' || $last === "\t") {
            return '<t xml:space="preserve">'.$escaped.'</t>';
        }

        return '<t>'.$escaped.'</t>';
    }

    /** Everything between the `<sst …>` open tag and `</sst>`, verbatim. */
    private static function extractBody(string $xml): string
    {
        if (! preg_match('/<(?:[A-Za-z_][\w.\-]*:)?sst\b[^>]*?(\/?)>/', $xml, $m, PREG_OFFSET_CAPTURE)) {
            return '';
        }
        if ($m[1][0] === '/') {
            return ''; // <sst/> — an empty table
        }

        $start = $m[0][1] + \strlen($m[0][0]);
        $end = strrpos($xml, '</');

        return $end === false || $end < $start ? '' : substr($xml, $start, $end - $start);
    }

    /**
     * Count `<si>` elements. Positional indices are the contract, so they are
     * derived from the elements themselves rather than a count attribute a
     * producer may have left stale.
     */
    private static function countEntries(string $body): int
    {
        return preg_match_all('/<(?:[A-Za-z_][\w.\-]*:)?si(?=[\s\/>])/', $body);
    }

    private static function attributeValue(string $xml, string $name): int
    {
        return preg_match('/<(?:[A-Za-z_][\w.\-]*:)?sst\b[^>]*?\b'.$name.'="(\d+)"/', $xml, $m)
            ? (int) $m[1]
            : 0;
    }

    /** Set (or add) a count attribute on the `<sst>` open tag. */
    private static function withCount(string $xml, string $name, int $value): string
    {
        $updated = preg_replace_callback(
            '/(<(?:[A-Za-z_][\w.\-]*:)?sst\b[^>]*?)\b'.$name.'="\d+"/',
            static fn (array $m): string => $m[1].$name.'="'.$value.'"',
            $xml,
            1,
            $hits
        );

        if ($hits > 0 && $updated !== null) {
            return $updated;
        }

        $added = preg_replace_callback(
            '/<(?:[A-Za-z_][\w.\-]*:)?sst\b[^>]*?(?=\/?>)/',
            static fn (array $m): string => $m[0].' '.$name.'="'.$value.'"',
            $xml,
            1
        );

        return $added ?? $xml;
    }
}

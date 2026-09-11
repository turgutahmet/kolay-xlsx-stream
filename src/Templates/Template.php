<?php

namespace Kolay\XlsxStream\Templates;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Readers\WorkbookResolver;
use Kolay\XlsxStream\Readers\ZipDirectory;
use Kolay\XlsxStream\Sources\LocalFileSource;
use Kolay\XlsxStream\Sources\StringSource;

/**
 * A workbook used as a layout: headers, merges, the style table, column
 * widths, freeze panes and rich text come from it; the rows come from the
 * streaming writer.
 *
 * The template is opened through the ordinary reader stack — `ZipDirectory`
 * for the central directory, `WorkbookResolver` for name → entry — so no new
 * parsing surface is introduced for the archive itself. It is **parsed once
 * and reused**: the same instance can back eight sheets in one job, or be
 * cached for the life of a process.
 *
 * The package deliberately does not understand the template. It reads the
 * sheet list, cuts one sheet at `<sheetData>` (see TemplateSheet), and treats
 * every other entry as opaque bytes to copy.
 */
class Template
{
    /** @var list<array{name: string, sheetId: int, entry: string}> */
    private array $sheets;

    /** @var array<string, string> zip entry => raw sheet XML, read once */
    private array $sheetXml = [];

    /** @var array<string, TemplateSheet> "name|dataStartRow" => cut sheet */
    private array $sheetCache = [];

    private function __construct(
        private Source $source,
        private ZipDirectory $cd,
        private bool $ownsSource,
    ) {
        $this->sheets = WorkbookResolver::resolve($source, $cd);
    }

    /**
     * Open a template from a path or any Source. A path is wrapped in a
     * LocalFileSource this instance owns and closes; a caller-supplied
     * Source stays the caller's to close.
     */
    public static function open(Source|string $pathOrSource): self
    {
        $ownsSource = \is_string($pathOrSource);
        $source = $ownsSource ? new LocalFileSource($pathOrSource) : $pathOrSource;

        return new self($source, ZipDirectory::fromSource($source), $ownsSource);
    }

    /**
     * Open a template that never touched disk — rendered in this process,
     * pulled from cache, or carried on a queue message.
     */
    public static function fromString(string $bytes): self
    {
        $source = new StringSource($bytes);

        return new self($source, ZipDirectory::fromSource($source), true);
    }

    /** @return list<string> sheet names in workbook order */
    public function sheetNames(): array
    {
        return array_column($this->sheets, 'name');
    }

    /** Zip entry backing a sheet name. */
    public function entryFor(string $name): string
    {
        foreach ($this->sheets as $sheet) {
            if ($sheet['name'] === $name) {
                return $sheet['entry'];
            }
        }

        throw XlsxStreamException::templateSheetNotFound($name, $this->sheetNames());
    }

    /** @return list<string> every zip entry, so untouched ones can be copied */
    public function entryNames(): array
    {
        return $this->cd->names();
    }

    public function hasEntry(string $name): bool
    {
        return $this->cd->has($name);
    }

    public function readEntry(string $name): string
    {
        return $this->cd->readEntry($this->source, $name);
    }

    /**
     * Cut one sheet at `<sheetData>`. Rows below `dataStartRow` stay as
     * headers; rows from it on are consumed as style-variant samples.
     * Memoised per (sheet, cut point).
     */
    public function sheet(string $name, int $dataStartRow): TemplateSheet
    {
        $key = $name.'|'.$dataStartRow;
        if (isset($this->sheetCache[$key])) {
            return $this->sheetCache[$key];
        }

        $entry = $this->entryFor($name);
        $xml = $this->sheetXml[$entry] ??= $this->readEntry($entry);

        return $this->sheetCache[$key] = TemplateSheet::parse($xml, $dataStartRow, $entry);
    }

    /** The underlying Source — the writer copies untouched entries from it. */
    public function source(): Source
    {
        return $this->source;
    }

    /** The parsed central directory — entry offsets, sizes and CRCs. */
    public function zip(): ZipDirectory
    {
        return $this->cd;
    }

    /** Release the Source when this instance opened it itself. */
    public function close(): void
    {
        if ($this->ownsSource) {
            $this->source->close();
        }
    }
}

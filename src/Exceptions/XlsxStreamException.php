<?php

namespace Kolay\XlsxStream\Exceptions;

/**
 * Base exception for all XLSX Stream related errors
 */
class XlsxStreamException extends \Exception
{
    /**
     * Create exception for sink write failure
     */
    public static function sinkWriteFailed(string $reason): self
    {
        return new self("Failed to write to sink: {$reason}");
    }

    /**
     * Create exception for sink close failure
     */
    public static function sinkCloseFailed(string $reason): self
    {
        return new self("Failed to close sink: {$reason}");
    }

    /**
     * Create exception for invalid compression level
     */
    public static function invalidCompressionLevel(int $level): self
    {
        return new self("Invalid compression level: {$level}. Must be between 1 and 9.");
    }

    /**
     * Create exception for invalid buffer size
     */
    public static function invalidBufferSize(int $size): self
    {
        return new self("Invalid buffer size: {$size}. Must be at least 1.");
    }

    /**
     * Create exception for writer already closed
     */
    public static function writerAlreadyClosed(): self
    {
        return new self('Cannot perform operation on closed writer.');
    }

    /**
     * Create exception for headers not set
     */
    public static function headersNotSet(): self
    {
        return new self('Headers must be set before writing rows. Call startFile() first.');
    }

    /**
     * Create exception for startFile called more than once
     */
    public static function alreadyStarted(): self
    {
        return new self('Writer has already been started. startFile() can only be called once.');
    }

    /**
     * Create exception for column count exceeding Excel's limit
     */
    public static function tooManyColumns(int $given, int $max): self
    {
        return new self("Column count {$given} exceeds Excel's maximum of {$max} columns per sheet.");
    }

    /**
     * Create exception for finalize-with-no-data
     */
    public static function emptyWorkbook(): self
    {
        return new self(
            'Cannot finalize an empty workbook. Write at least one row via writeRow() '.
            'or call newSheet() to create a sheet before finishFile().'
        );
    }

    /**
     * Create exception for setColumnFormat targeting a column past the header count.
     */
    public static function columnIndexOutOfRange(int $given, int $max): self
    {
        return new self(
            "Column index {$given} is out of range — the current header has {$max} columns. ".
            'Call setColumnFormat() with an index between 1 and the header count.'
        );
    }

    /**
     * Raised when an export approaches a 32-bit ZIP container limit.
     * Loud rejection beats silently truncating size fields and shipping
     * a corrupt archive — split the export across multiple files until
     * ZIP64 writer support lands.
     */
    public static function zip32LimitExceeded(string $detail): self
    {
        return new self(
            "ZIP32 limit exceeded: {$detail}. ".
            'Split the export across multiple files or sheets as a workaround. '.
            'ZIP64 writer support is tracked for a future release.'
        );
    }

    /**
     * Create exception for a template sheet name that is not in the workbook
     */
    public static function templateSheetNotFound(string $name, array $available): self
    {
        $list = $available === [] ? '(none)' : implode(', ', $available);

        return new self("Template has no sheet named '{$name}'. Available: {$list}.");
    }

    /**
     * Create exception for a template sheet with no sheetData element
     */
    public static function templateSheetDataMissing(string $entry): self
    {
        return new self("Template sheet '{$entry}' has no <sheetData> element to stream into.");
    }

    /**
     * Create exception for a merge that reaches into the streamed data region
     */
    public static function templateMergeInDataRegion(string $ref, int $dataStartRow): self
    {
        return new self(
            "Template merge '{$ref}' reaches row {$dataStartRow} or below, inside the data region. ".
            'Merges over streamed rows are not supported; keep merges above dataStartRow.'
        );
    }

    /**
     * Create exception for a template whose sample-row block is implausibly large
     */
    public static function templateTooManySampleRows(int $count, int $max): self
    {
        return new self(
            "Template declares {$count} sample rows (limit {$max}). ".
            'A template is a layout: put one row per style variant below dataStartRow, not a full report.'
        );
    }

    /**
     * Create exception for a sheet whose dialect the row builders cannot write into
     */
    public static function templateUnsupportedDialect(string $name): self
    {
        return new self(
            "Template sheet '{$name}' binds SpreadsheetML only to a namespace prefix, ".
            'never as the default namespace. Streamed rows are written unprefixed, so they '.
            'would land in no namespace at all and Excel would offer to repair the file. '.
            'Re-save the template from Excel or PhpSpreadsheet, which both declare the default namespace.'
        );
    }

    /**
     * Create exception for a template operation attempted before a sheet was chosen
     */
    public static function templateSheetNotSelected(): self
    {
        return new self(
            'No template sheet is being streamed. Call sheet() with the sheet name '.
            'and the row its data starts on before writing rows.'
        );
    }

    /**
     * Create exception for a template sheet chosen twice
     */
    public static function templateSheetAlreadyStreamed(string $name): self
    {
        return new self(
            "Template sheet '{$name}' has already been streamed. ".
            'Each sheet is cut once; write all of its rows before moving to the next.'
        );
    }

    /**
     * Create exception for a classic operation that template mode forbids
     */
    public static function templateModeForbids(string $operation, string $because): self
    {
        return new self("{$operation} is not available in template mode — {$because}.");
    }

    /**
     * Create exception for a template operation attempted on a classic writer
     */
    public static function templateModeRequired(string $operation): self
    {
        return new self("{$operation} requires template mode. Attach a template with useTemplate() first.");
    }

    /**
     * Create exception for a template attached to an already-configured writer
     */
    public static function templateConflictsWith(string $what): self
    {
        return new self(
            "A template cannot be attached after {$what} — the template owns the layout. ".
            'Call useTemplate() on a fresh writer.'
        );
    }

    /**
     * Create exception for a template sheet that ran past Excel's row limit
     */
    public static function templateSheetRowLimit(string $name, int $limit): self
    {
        return new self(
            "Template sheet '{$name}' reached row {$limit}, Excel's per-sheet limit. ".
            'Template mode cannot auto-split — the overflow sheet would have no layout to inherit. '.
            'Split the data across template sheets yourself.'
        );
    }
}

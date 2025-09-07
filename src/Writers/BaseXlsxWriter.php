<?php

namespace Kolay\XlsxStream\Writers;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;

/**
 * Base XLSX Writer - Core streaming functionality
 * 
 * Excel limits:
 * - Max rows per sheet: 1,048,576
 * - Max columns: 16,384 (XFD)
 */
abstract class BaseXlsxWriter
{
    protected array $centralDirectory = [];
    protected int $currentOffset = 0;
    
    // ZIP constants
    const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
    const CENTRAL_FILE_HEADER_SIGNATURE = 0x02014b50;
    const END_OF_CENTRAL_DIR_SIGNATURE = 0x06054b50;
    const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;
    
    // Compression methods
    const COMPRESSION_STORED = 0;
    const COMPRESSION_DEFLATED = 8;
    
    // Version info
    const VERSION_MADE_BY = 0x001E; // 3.0 UNIX
    const VERSION_NEEDED = 0x0014;  // 2.0
    
    // Excel limits
    const MAX_ROWS_PER_SHEET = 1048576; // Excel's hard limit
    const ROWS_PER_SHEET = 1048575; // MAX - 1 for header safety
    
    // Sheet management
    protected array $sheets = [];
    protected int $currentSheetIndex = 0;
    protected array $columns = [];
    
    // Current sheet streaming variables
    protected $deflateCtx = null;
    protected $crcContext = null;
    protected int $sheetCrc = 0;
    protected int $sheetUncompressedSize = 0;
    protected int $sheetCompressedSize = 0;
    protected int $sheetOffset = 0;
    protected int $currentSheetRow = 0;
    protected int $totalRows = 0;
    
    // Performance optimizations
    protected int $bufferFlushInterval = 1000;
    protected string $rowBuffer = '';
    protected int $rowBufferCount = 0;
    protected int $deflateLevel = 6;
    
    // Column letter cache for performance
    protected array $colLetterCache = [];
    
    /**
     * Write data to destination (must be implemented by child classes)
     */
    abstract protected function writeToDest(string $data): void;
    
    /**
     * Set deflate compression level (1-9)
     * 3 = fast (20-35% speed boost, larger files)
     * 6 = balanced (default)
     * 9 = best compression (slower)
     */
    public function setCompressionLevel(int $level): self
    {
        if ($level < 1 || $level > 9) {
            throw XlsxStreamException::invalidCompressionLevel($level);
        }
        $this->deflateLevel = $level;
        return $this;
    }
    
    /**
     * Set row buffer flush interval
     * Lower = more responsive streaming
     * Higher = better compression ratio
     */
    public function setBufferFlushInterval(int $rows): self
    {
        if ($rows < 1) {
            throw XlsxStreamException::invalidBufferSize($rows);
        }
        $this->bufferFlushInterval = $rows;
        return $this;
    }
    
    /**
     * Convert Unix timestamp to DOS time and date (separate fields)
     */
    protected function dosTimeParts(int $timestamp): array
    {
        $d = getdate($timestamp);
        
        // DOS Time: bits 15-11: hours, 10-5: minutes, 4-0: seconds/2
        $dosTime = (($d['hours'] & 0x1F) << 11) |
                   (($d['minutes'] & 0x3F) << 5) |
                   (($d['seconds'] >> 1) & 0x1F);
        
        // DOS Date: bits 15-9: year-1980, 8-5: month, 4-0: day
        $dosDate = ((($d['year'] - 1980) & 0x7F) << 9) |
                   (($d['mon'] & 0x0F) << 5) |
                   (($d['mday'] & 0x1F));
        
        return [$dosTime, $dosDate];
    }
    
    /**
     * Start XLSX file with headers and static files
     */
    public function startFile(array $headers): void
    {
        $this->columns = $headers;
        
        // Write only static files that don't depend on sheet count
        $this->writeStaticFile('_rels/.rels', $this->getRelsXml());
        $this->writeStaticFile('xl/styles.xml', $this->getStylesXml());
    }
    
    /**
     * Write a single row (handles multi-sheet automatically)
     */
    public function writeRow(array $row): void
    {
        // Check if we need to start a new sheet
        if ($this->currentSheetRow === 0 || $this->currentSheetRow >= self::ROWS_PER_SHEET) {
            if ($this->currentSheetRow > 0) {
                $this->flushRowBuffer();
                $this->finishCurrentSheet();
            }
            $this->currentSheetIndex++;
            $this->startNewSheet();
        }
        
        $this->currentSheetRow++;
        $this->totalRows++;
        
        // Build row XML and add to buffer
        $this->rowBuffer .= $this->buildRowXml($this->currentSheetRow, $row);
        $this->rowBufferCount++;
        
        // Flush buffer periodically for better streaming
        if ($this->rowBufferCount >= $this->bufferFlushInterval) {
            $this->flushRowBuffer();
        }
    }
    
    /**
     * Write multiple rows efficiently
     */
    public function writeRows(array $rows): void
    {
        foreach ($rows as $row) {
            $this->writeRow($row);
        }
    }
    
    /**
     * Flush row buffer to stream
     */
    protected function flushRowBuffer(): void
    {
        if ($this->rowBuffer !== '') {
            $this->writeSheetData($this->rowBuffer);
            $this->rowBuffer = '';
            $this->rowBufferCount = 0;
        }
    }
    
    /**
     * Sanitize sheet name for Excel compatibility
     */
    protected function sanitizeSheetName(string $name): string
    {
        $name = preg_replace('/[:*?\/\\\[\]]/', '_', $name);
        
        if (mb_strlen($name, 'UTF-8') > 31) {
            $name = mb_substr($name, 0, 31, 'UTF-8');
        }
        
        return $name === '' ? 'Sheet' : $name;
    }
    
    /**
     * Start a new sheet
     */
    protected function startNewSheet(): void
    {
        $sheetName = "Sheet{$this->currentSheetIndex}";
        if ($this->currentSheetIndex === 1) {
            $sheetName = "Report";
        }
        
        $sheetName = $this->sanitizeSheetName($sheetName);
        $filename = "xl/worksheets/sheet{$this->currentSheetIndex}.xml";
        
        $this->sheets[] = [
            'index' => $this->currentSheetIndex,
            'name' => $sheetName,
            'filename' => $filename,
            'rows' => 0
        ];
        
        $this->sheetOffset = $this->currentOffset;
        $this->currentSheetRow = 0;
        
        [$mtime, $mdate] = $this->dosTimeParts(time());
        
        // Write local file header
        $header = pack('V', self::LOCAL_FILE_HEADER_SIGNATURE);
        $header .= pack('v', self::VERSION_NEEDED);
        $header .= pack('v', 0x0008);
        $header .= pack('v', self::COMPRESSION_DEFLATED);
        $header .= pack('v', $mtime);
        $header .= pack('v', $mdate);
        $header .= pack('V', 0);
        $header .= pack('V', 0);
        $header .= pack('V', 0);
        $header .= pack('v', strlen($filename));
        $header .= pack('v', 0);
        $header .= $filename;
        
        $this->writeToDest($header);
        
        // Initialize deflate context
        $this->deflateCtx = deflate_init(ZLIB_ENCODING_RAW, ['level' => $this->deflateLevel]);
        $this->crcContext = hash_init('crc32b');
        $this->sheetCrc = 0;
        $this->sheetUncompressedSize = 0;
        $this->sheetCompressedSize = 0;
        $this->rowBuffer = '';
        $this->rowBufferCount = 0;
        
        // Write sheet header
        $sheetHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $sheetHeader .= '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
        $sheetHeader .= '<sheetData>';
        
        if (!empty($this->columns)) {
            $headerRow = '<row r="1">';
            foreach ($this->columns as $i => $header) {
                $cellRef = $this->getColumnLetter($i + 1) . '1';
                $escaped = $this->fastXmlEscape($header);
                $headerRow .= '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
            }
            $headerRow .= '</row>';
            $sheetHeader .= $headerRow;
            $this->currentSheetRow = 1;
        }
        
        $this->writeSheetData($sheetHeader);
    }
    
    /**
     * Write data to current sheet with streaming compression
     */
    protected function writeSheetData(string $data): void
    {
        hash_update($this->crcContext, $data);
        $this->sheetUncompressedSize += strlen($data);
        
        $compressed = deflate_add($this->deflateCtx, $data, ZLIB_NO_FLUSH);
        if ($compressed !== false && strlen($compressed) > 0) {
            $this->writeToDest($compressed);
            $this->sheetCompressedSize += strlen($compressed);
        }
    }
    
    /**
     * Finish current sheet
     */
    protected function finishCurrentSheet(): void
    {
        $this->flushRowBuffer();
        
        $sheetFooter = '</sheetData></worksheet>';
        
        hash_update($this->crcContext, $sheetFooter);
        $this->sheetUncompressedSize += strlen($sheetFooter);
        
        $this->sheetCrc = hexdec(hash_final($this->crcContext));
        
        $compressed = deflate_add($this->deflateCtx, $sheetFooter, ZLIB_FINISH);
        if ($compressed !== false) {
            $this->writeToDest($compressed);
            $this->sheetCompressedSize += strlen($compressed);
        }
        
        // Write data descriptor
        $descriptor = pack('V', self::DATA_DESCRIPTOR_SIGNATURE);
        $descriptor .= pack('V', $this->sheetCrc);
        $descriptor .= pack('V', $this->sheetCompressedSize);
        $descriptor .= pack('V', $this->sheetUncompressedSize);
        
        $this->writeToDest($descriptor);
        
        // Add to central directory
        $sheetInfo = end($this->sheets);
        $this->centralDirectory[] = [
            'filename' => $sheetInfo['filename'],
            'crc32' => $this->sheetCrc,
            'compressed_size' => $this->sheetCompressedSize,
            'uncompressed_size' => $this->sheetUncompressedSize,
            'offset' => $this->sheetOffset,
            'compression' => self::COMPRESSION_DEFLATED,
            'flags' => 0x0008,
            'timestamp' => time()
        ];
        
        $this->sheets[count($this->sheets) - 1]['rows'] = $this->currentSheetRow;
    }
    
    /**
     * Build row XML with optimization
     */
    protected function buildRowXml(int $rowIndex, array $data): string
    {
        $cells = [];
        $colIndex = 1;
        
        foreach ($data as $value) {
            $cellRef = $this->getColumnLetter($colIndex) . $rowIndex;
            $colIndex++;
            
            if ($value === null || $value === '') {
                $cells[] = '<c r="' . $cellRef . '"/>';
            } elseif (is_numeric($value)) {
                if (is_string($value) && strlen($value) > 1 && $value[0] === '0' && $value[1] !== '.') {
                    $escaped = $this->fastXmlEscape($value);
                    $cells[] = '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
                } else {
                    $cells[] = '<c r="' . $cellRef . '" t="n"><v>' . (0 + $value) . '</v></c>';
                }
            } else {
                $escaped = $this->fastXmlEscape((string)$value);
                
                if ($escaped !== '' && ($escaped[0] === ' ' || $escaped[0] === "\t" || 
                    $escaped[strlen($escaped) - 1] === ' ' || $escaped[strlen($escaped) - 1] === "\t")) {
                    $cells[] = '<c r="' . $cellRef . '" t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>';
                } else {
                    $cells[] = '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $escaped . '</t></is></c>';
                }
            }
        }
        
        return '<row r="' . $rowIndex . '">' . implode('', $cells) . '</row>';
    }
    
    /**
     * Get column letter with caching
     */
    protected function getColumnLetter(int $index): string
    {
        if (!isset($this->colLetterCache[$index])) {
            $n = $index;
            $s = '';
            while ($n > 0) {
                $n--;
                $s = chr(65 + ($n % 26)) . $s;
                $n = intdiv($n, 26);
            }
            $this->colLetterCache[$index] = $s;
        }
        return $this->colLetterCache[$index];
    }
    
    /**
     * Ultra-fast XML escaping
     */
    protected function fastXmlEscape(string $str): string
    {
        static $trans = [
            '&' => '&amp;',
            '<' => '&lt;',
            '>' => '&gt;',
            '"' => '&quot;',
            "'" => '&apos;'
        ];
        
        if (strpbrk($str, '&<>"\'\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0B\x0C\x0E\x0F\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F') === false) {
            return $str;
        }
        
        $str = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F]/', '', $str);
        
        return strtr($str, $trans);
    }
    
    /**
     * Write static ZIP entry
     */
    protected function writeStaticFile(string $filename, string $content): void
    {
        $uncompressedSize = strlen($content);
        $compressedContent = gzdeflate($content, $this->deflateLevel);
        $compressedSize = strlen($compressedContent);
        $crc32 = crc32($content);
        
        [$mtime, $mdate] = $this->dosTimeParts(time());
        
        $this->centralDirectory[] = [
            'filename' => $filename,
            'crc32' => $crc32,
            'compressed_size' => $compressedSize,
            'uncompressed_size' => $uncompressedSize,
            'offset' => $this->currentOffset,
            'compression' => self::COMPRESSION_DEFLATED,
            'flags' => 0x0000,
            'timestamp' => time()
        ];
        
        $header = pack('V', self::LOCAL_FILE_HEADER_SIGNATURE);
        $header .= pack('v', self::VERSION_NEEDED);
        $header .= pack('v', 0x0000);
        $header .= pack('v', self::COMPRESSION_DEFLATED);
        $header .= pack('v', $mtime);
        $header .= pack('v', $mdate);
        $header .= pack('V', $crc32);
        $header .= pack('V', $compressedSize);
        $header .= pack('V', $uncompressedSize);
        $header .= pack('v', strlen($filename));
        $header .= pack('v', 0);
        $header .= $filename;
        
        $this->writeToDest($header);
        $this->writeToDest($compressedContent);
    }
    
    /**
     * Write ZIP central directory
     */
    protected function writeCentralDirectory(): void
    {
        $centralDirStart = $this->currentOffset;
        $centralDirSize = 0;
        
        foreach ($this->centralDirectory as $entry) {
            [$mtime, $mdate] = $this->dosTimeParts($entry['timestamp']);
            
            $header = pack('V', self::CENTRAL_FILE_HEADER_SIGNATURE);
            $header .= pack('v', self::VERSION_MADE_BY);
            $header .= pack('v', self::VERSION_NEEDED);
            $header .= pack('v', $entry['flags']);
            $header .= pack('v', $entry['compression']);
            $header .= pack('v', $mtime);
            $header .= pack('v', $mdate);
            $header .= pack('V', $entry['crc32']);
            $header .= pack('V', $entry['compressed_size']);
            $header .= pack('V', $entry['uncompressed_size']);
            $header .= pack('v', strlen($entry['filename']));
            $header .= pack('v', 0);
            $header .= pack('v', 0);
            $header .= pack('v', 0);
            $header .= pack('v', 0);
            $header .= pack('V', 0x81A40000);
            $header .= pack('V', $entry['offset']);
            $header .= $entry['filename'];
            
            $this->writeToDest($header);
            $centralDirSize += strlen($header);
        }
        
        $endRecord = pack('V', self::END_OF_CENTRAL_DIR_SIGNATURE);
        $endRecord .= pack('v', 0);
        $endRecord .= pack('v', 0);
        $endRecord .= pack('v', count($this->centralDirectory));
        $endRecord .= pack('v', count($this->centralDirectory));
        $endRecord .= pack('V', $centralDirSize);
        $endRecord .= pack('V', $centralDirStart);
        $endRecord .= pack('v', 0);
        
        $this->writeToDest($endRecord);
    }
    
    /**
     * Finalize the XLSX file
     */
    public function finishFile(): array
    {
        if ($this->currentSheetRow > 0) {
            $this->flushRowBuffer();
            $this->finishCurrentSheet();
        }
        
        $this->writeStaticFile('xl/_rels/workbook.xml.rels', $this->getWorkbookRelsXml());
        $this->writeStaticFile('xl/workbook.xml', $this->getWorkbookXml());
        $this->writeStaticFile('[Content_Types].xml', $this->getContentTypesXml());
        
        $this->writeCentralDirectory();
        
        return [
            'bytes' => $this->currentOffset,
            'rows' => $this->totalRows,
            'sheets' => count($this->sheets),
            'sheet_details' => $this->sheets
        ];
    }
    
    // XLSX structure generators
    
    protected function getContentTypesXml(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>';
        
        for ($i = 1; $i <= count($this->sheets); $i++) {
            $xml .= "\n    " . '<Override PartName="/xl/worksheets/sheet' . $i . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        
        $xml .= '
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>';
        
        return $xml;
    }
    
    protected function getRelsXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
    }
    
    protected function getWorkbookRelsXml(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        
        for ($i = 1; $i <= count($this->sheets); $i++) {
            $xml .= "\n    " . '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $i . '.xml"/>';
        }
        
        $styleId = count($this->sheets) + 1;
        $xml .= "\n    " . '<Relationship Id="rId' . $styleId . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        
        $xml .= '
</Relationships>';
        
        return $xml;
    }
    
    protected function getWorkbookXml(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>';
        
        foreach ($this->sheets as $sheet) {
            $escapedName = $this->fastXmlEscape($sheet['name']);
            $xml .= "\n        " . '<sheet name="' . $escapedName . '" sheetId="' . $sheet['index'] . '" r:id="rId' . $sheet['index'] . '"/>';
        }
        
        $xml .= '
    </sheets>
</workbook>';
        
        return $xml;
    }
    
    protected function getStylesXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1">
        <font><sz val="11"/><name val="Calibri"/></font>
    </fonts>
    <fills count="2">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    </fills>
    <borders count="1">
        <border><left/><right/><top/><bottom/><diagonal/></border>
    </borders>
    <cellXfs count="1">
        <xf fontId="0" fillId="0" borderId="0"/>
    </cellXfs>
</styleSheet>';
    }
}
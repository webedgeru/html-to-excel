<?php

namespace sberdyug\htmltoexcel\services;

use DOMDocument;
use DOMElement;
use DOMNode;
use DOMXPath;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use yii\base\Component;
use yii\base\InvalidArgumentException;
use yii\helpers\ArrayHelper;

/**
 * HtmlTableToExcelParser
 *
 * Parser that converts HTML table(s) to a PhpSpreadsheet Spreadsheet instance.
 * Supports:
 *  - Multiple <table> elements (either split into separate sheets or stacked on one sheet)
 *  - <thead>, <tbody>, <tfoot>, <tr>, <th>, <td>
 *  - Rowspan and Colspan merging
 *  - Inline styles and attributes (colors, background, font-weight, font-size, alignments, borders)
 *  - Hyperlinks (<a href="...">)
 *  - Data types detection (numbers, formulas, strings, dates)
 *  - Custom attributes (data-type, data-format, data-sheet-name)
 *  - Auto column sizing
 */
class HtmlTableToExcelParser extends Component
{
    /**
     * @var bool Whether to create a new worksheet for each table found in HTML.
     * If false, all tables are written sequentially to the same worksheet.
     */
    public $separateSheetsForTables = true;

    /**
     * @var bool Whether to parse and apply inline CSS and HTML styling.
     */
    public $parseStyles = true;

    /**
     * @var bool Whether to auto-size columns based on content.
     */
    public $autoSizeColumns = true;

    /**
     * @var bool Whether to automatically detect and cast numeric values and formulas.
     */
    public $detectDataTypes = true;

    /**
     * @var string Default font family for the generated sheets.
     */
    public $defaultFontFamily = 'Calibri';

    /**
     * @var float Default font size in points.
     */
    public $defaultFontSize = 13.0;

    /**
     * @var int Number of blank rows between tables when separateSheetsForTables is false.
     */
    public $tableSpacing = 5;

    /**
     * @var Spreadsheet|null
     */
    protected $_spreadsheet;

    /**
     * @var array Standard HTML named colors map to Hex RGB
     */
    protected static $namedColors = [
        'black' => '000000',
        'white' => 'FFFFFF',
        'red' => 'FF0000',
        'green' => '008000',
        'blue' => '0000FF',
        'yellow' => 'FFFF00',
        'gray' => '808080',
        'grey' => '808080',
        'silver' => 'C0C0C0',
        'maroon' => '800000',
        'purple' => '800080',
        'fuchsia' => 'FF00FF',
        'lime' => '00FF00',
        'olive' => '808000',
        'navy' => '000080',
        'teal' => '008080',
        'aqua' => '00FFFF',
        'orange' => 'FFA500',
        'lightgray' => 'D3D3D3',
        'lightgrey' => 'D3D3D3',
        'darkgray' => 'A9A9A9',
        'darkgrey' => 'A9A9A9',
    ];

    /**
     * Parse HTML string and return a PhpSpreadsheet Spreadsheet instance.
     *
     * @param string $html HTML content containing one or more <table> tags.
     * @param Spreadsheet|null $existingSpreadsheet Optional existing spreadsheet to append to.
     * @return Spreadsheet
     */
    public function parse(string $html, ?Spreadsheet $existingSpreadsheet = null): Spreadsheet
    {
        $this->_spreadsheet = $existingSpreadsheet ?? new Spreadsheet();

        // Configure default styles
        $this->_spreadsheet->getDefaultStyle()->getFont()->setName($this->defaultFontFamily);
        $this->_spreadsheet->getDefaultStyle()->getFont()->setSize($this->defaultFontSize);

        if (empty(trim($html))) {
            return $this->_spreadsheet;
        }

        $dom = $this->createDomDocument($html);
        $xpath = new DOMXPath($dom);
        $tables = $xpath->query('//table');

        if ($tables === false || $tables->length === 0) {
            // If no <table> tag, return empty spreadsheet
            return $this->_spreadsheet;
        }

        $currentSheetIndex = 0;
        $currentRow = 1;

        foreach ($tables as $index => $tableNode) {
            /** @var DOMElement $tableNode */
            if ($this->separateSheetsForTables) {
                if ($index === 0) {
                    $sheet = $this->_spreadsheet->getActiveSheet();
                } else {
                    $sheet = $this->_spreadsheet->createSheet();
                }
                $sheetTitle = $this->determineSheetTitle($tableNode, $index + 1);
                $this->setSafeSheetTitle($sheet, $sheetTitle);
                $currentRow = 1;
            } else {
                $sheet = $this->_spreadsheet->getActiveSheet();
                if ($index > 0) {
                    $currentRow += $this->tableSpacing;
                }
            }

            $currentRow = $this->parseTable($tableNode, $sheet, $currentRow);

            if ($this->autoSizeColumns) {
                $this->applyAutoColumnSizes($sheet);
            }
        }

        return $this->_spreadsheet;
    }

    /**
     * Parse HTML file by path.
     *
     * @param string $filePath
     * @param Spreadsheet|null $existingSpreadsheet
     * @return Spreadsheet
     * @throws InvalidArgumentException
     */
    public function parseFile(string $filePath, ?Spreadsheet $existingSpreadsheet = null): Spreadsheet
    {
        if (!file_exists($filePath) || !is_readable($filePath)) {
            throw new InvalidArgumentException("File '{$filePath}' does not exist or is not readable.");
        }

        $content = file_get_contents($filePath);
        return $this->parse($content, $existingSpreadsheet);
    }

    /**
     * Parse a single <table> DOM element into the given Worksheet.
     *
     * @param DOMElement $tableNode
     * @param Worksheet $sheet
     * @param int $startRow
     * @return int Next available row number after this table
     */
    protected function parseTable(DOMElement $tableNode, Worksheet $sheet, int $startRow = 1): int
    {
        // Matrix to track occupied cells by row and col index (1-based)
        $grid = [];
        $currentRow = $startRow;

        $xpath = new DOMXPath($tableNode->ownerDocument);

        // Find all rows (tr) directly or within thead, tbody, tfoot
        // using relative xpath './tr | ./thead/tr | ./tbody/tr | ./tfoot/tr'
        $rows = $xpath->query('.//tr', $tableNode);

        if ($rows === false) {
            return $currentRow;
        }

        foreach ($rows as $rowNode) {
            /** @var DOMElement $rowNode */

            // Apply row height if present
            $this->applyRowStyle($rowNode, $sheet, $currentRow);

            $currentCol = 1;
            $cells = $xpath->query('./th | ./td', $rowNode);

            if ($cells === false) {
                $currentRow++;
                continue;
            }

            foreach ($cells as $cellNode) {
                /** @var DOMElement $cellNode */

                // Advance currentCol until an unoccupied cell is found in $grid
                while (isset($grid[$currentRow][$currentCol])) {
                    $currentCol++;
                }

                $rowspan = (int) $cellNode->getAttribute('rowspan');
                if ($rowspan < 1) {
                    $rowspan = 1;
                }

                $colspan = (int) $cellNode->getAttribute('colspan');
                if ($colspan < 1) {
                    $colspan = 1;
                }

                // Mark grid cells as occupied for this cell's span
                for ($r = 0; $r < $rowspan; $r++) {
                    for ($c = 0; $c < $colspan; $c++) {
                        $grid[$currentRow + $r][$currentCol + $c] = true;
                    }
                }

                // Parse and write cell data
                $this->writeCell($cellNode, $sheet, $currentCol, $currentRow, $colspan, $rowspan);

                // Advance column for next cell
                $currentCol += $colspan;
            }

            $currentRow++;
        }

        return $currentRow;
    }

    /**
     * Write content, data type, styles and merges to the cell.
     *
     * @param DOMElement $cellNode
     * @param Worksheet $sheet
     * @param int $col 1-based column index
     * @param int $row 1-based row index
     * @param int $colspan
     * @param int $rowspan
     */
    protected function writeCell(DOMElement $cellNode, Worksheet $sheet, int $col, int $row, int $colspan, int $rowspan): void
    {
        $cellCoordinate = Coordinate::stringFromColumnIndex($col) . $row;
        $cell = $sheet->getCell($cellCoordinate);

        // Handle cell text / value
        $rawValue = $this->extractCellText($cellNode);
        $explicitType = strtolower(trim($cellNode->getAttribute('data-type')));

        if ($explicitType !== '') {
            $this->setExplicitCellValue($cell, $rawValue, $explicitType);
        } elseif ($this->detectDataTypes) {
            $this->setAutoCellValue($cell, $rawValue);
        } else {
            $cell->setValueExplicit($rawValue, DataType::TYPE_STRING);
        }

        // Handle hyperlinks
        $linkNode = $cellNode->getElementsByTagName('a')->item(0);
        if ($linkNode instanceof DOMElement && $linkNode->hasAttribute('href')) {
            $href = trim($linkNode->getAttribute('href'));
            if ($href !== '') {
                $cell->getHyperlink()->setUrl($href);
                if ($linkNode->hasAttribute('title')) {
                    $cell->getHyperlink()->setTooltip($linkNode->getAttribute('title'));
                }
            }
        }

        // Apply styles (borders, font, background, alignment)
        if ($this->parseStyles) {
            $this->applyCellStyle($cellNode, $sheet, $col, $row, $colspan, $rowspan);
        }

        // Merge cells if colspan or rowspan > 1
        if ($colspan > 1 || $rowspan > 1) {
            $endCol = $col + $colspan - 1;
            $endRow = $row + $rowspan - 1;
            $endCoordinate = Coordinate::stringFromColumnIndex($endCol) . $endRow;
            $range = "{$cellCoordinate}:{$endCoordinate}";
            $sheet->mergeCells($range);
        }
    }

    /**
     * Extract cell text content, cleaning up whitespace while preserving text.
     *
     * @param DOMElement $cellNode
     * @return string
     */
    protected function extractCellText(DOMElement $cellNode): string
    {
        // Check if there are <br> tags and convert them to newlines
        $html = '';
        foreach ($cellNode->childNodes as $child) {
            if ($child->nodeType === XML_ELEMENT_NODE && strtolower($child->nodeName) === 'br') {
                $html .= "\n";
            } else {
                $html .= $child->textContent;
            }
        }

        $text = html_entity_decode($html, ENT_QUOTES | ENT_HTML5, 'UTF-8');
        // Normalize line breaks and spaces
        $text = preg_replace('/[ \t]+/', ' ', $text);
        $text = preg_replace('/\n\s+/', "\n", $text);
        return trim($text);
    }

    /**
     * Sets cell value according to an explicit data-type attribute.
     *
     * @param \PhpOffice\PhpSpreadsheet\Cell\Cell $cell
     * @param string $value
     * @param string $type ('string', 'number', 'numeric', 'formula', 'bool', 'boolean', 'date')
     */
    protected function setExplicitCellValue($cell, string $value, string $type): void
    {
        switch ($type) {
            case 'numeric':
            case 'number':
            case 'int':
            case 'float':
                $cleaned = str_replace([' ', ','], ['', '.'], $value);
                if (is_numeric($cleaned)) {
                    $cell->setValueExplicit((float) $cleaned, DataType::TYPE_NUMERIC);
                } else {
                    $cell->setValueExplicit($value, DataType::TYPE_STRING);
                }
                break;
            case 'formula':
                $cell->setValueExplicit($value, DataType::TYPE_FORMULA);
                break;
            case 'bool':
            case 'boolean':
                $boolVal = in_array(strtolower($value), ['1', 'true', 'yes', 'on', 'да', 'истина'], true);
                $cell->setValueExplicit($boolVal, DataType::TYPE_BOOL);
                break;
            case 'string':
            default:
                $cell->setValueExplicit($value, DataType::TYPE_STRING);
                break;
        }
    }

    /**
     * Automatically detect data type and set value on cell.
     *
     * @param \PhpOffice\PhpSpreadsheet\Cell\Cell $cell
     * @param string $value
     */
    protected function setAutoCellValue($cell, string $value): void
    {
        if ($value === '') {
            $cell->setValueExplicit('', DataType::TYPE_STRING);
            return;
        }

        // Formula check
        if (strpos($value, '=') === 0 && strlen($value) > 1) {
            $cell->setValueExplicit($value, DataType::TYPE_FORMULA);
            return;
        }

        // Pure integer or float (with optional dot or comma)
        // Avoid treating phone numbers or leading zeros as numbers (e.g., "01234")
        if (preg_match('/^-?\d+(\.\d+)?$/', $value)) {
            if (strlen($value) > 1 && $value[0] === '0' && $value[1] !== '.') {
                // Leading zero like '0123' -> string
                $cell->setValueExplicit($value, DataType::TYPE_STRING);
            } else {
                $cell->setValueExplicit((float) $value, DataType::TYPE_NUMERIC);
            }
            return;
        }

        // Check number with comma as decimal: e.g. "123,45"
        if (preg_match('/^-?\d+,\d+$/', $value)) {
            $numeric = (float) str_replace(',', '.', $value);
            $cell->setValueExplicit($numeric, DataType::TYPE_NUMERIC);
            return;
        }

        // Default as string
        $cell->setValueExplicit($value, DataType::TYPE_STRING);
    }

    /**
     * Parse and apply cell styles (font, fill, alignment, borders).
     *
     * @param DOMElement $cellNode
     * @param Worksheet $sheet
     * @param int $col
     * @param int $row
     * @param int $colspan
     * @param int $rowspan
     */
    protected function applyCellStyle(DOMElement $cellNode, Worksheet $sheet, int $col, int $row, int $colspan, int $rowspan): void
    {
        $startCoord = Coordinate::stringFromColumnIndex($col) . $row;
        $endCoord = Coordinate::stringFromColumnIndex($col + $colspan - 1) . ($row + $rowspan - 1);
        $range = ($colspan > 1 || $rowspan > 1) ? "{$startCoord}:{$endCoord}" : $startCoord;

        $style = $sheet->getStyle($range);

        $isHeader = strtolower($cellNode->nodeName) === 'th';
        if ($isHeader) {
            $style->getFont()->setBold(true);
            $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }

        // Check if inside <thead>
        $parent = $cellNode->parentNode;
        if ($parent && $parent->parentNode && strtolower($parent->parentNode->nodeName) === 'thead') {
            $style->getFont()->setBold(true);
        }

        // Inline CSS parsing
        $inlineStyle = $cellNode->getAttribute('style');
        $css = $this->parseCssString($inlineStyle);

        // HTML attributes fallback
        $align = $cellNode->getAttribute('align') ?: ($css['text-align'] ?? null);
        $valign = $cellNode->getAttribute('valign') ?: ($css['vertical-align'] ?? null);
        $bgcolor = $cellNode->getAttribute('bgcolor') ?: ($css['background-color'] ?? ($css['background'] ?? null));

        // Horizontal alignment
        if ($align) {
            switch (strtolower(trim($align))) {
                case 'center':
                    $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                    break;
                case 'right':
                    $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                    break;
                case 'left':
                    $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
                    break;
                case 'justify':
                    $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_JUSTIFY);
                    break;
            }
        }

        // Vertical alignment
        if ($valign) {
            switch (strtolower(trim($valign))) {
                case 'top':
                    $style->getAlignment()->setVertical(Alignment::VERTICAL_TOP);
                    break;
                case 'middle':
                case 'center':
                    $style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                    break;
                case 'bottom':
                    $style->getAlignment()->setVertical(Alignment::VERTICAL_BOTTOM);
                    break;
            }
        }

        // Background color
        if ($bgcolor) {
            $hexColor = $this->normalizeColorToHex($bgcolor);
            if ($hexColor) {
                $style->getFill()->setFillType(Fill::FILL_SOLID);
                $style->getFill()->getStartColor()->setARGB('FF' . $hexColor);
            }
        }

        // Font styling
        $fontColor = $css['color'] ?? $cellNode->getAttribute('color');
        if ($fontColor) {
            $hexColor = $this->normalizeColorToHex($fontColor);
            if ($hexColor) {
                $style->getFont()->getColor()->setARGB('FF' . $hexColor);
            }
        }

        if (isset($css['font-weight'])) {
            $fw = strtolower(trim($css['font-weight']));
            if ($fw === 'bold' || $fw === 'bolder' || (is_numeric($fw) && (int) $fw >= 600)) {
                $style->getFont()->setBold(true);
            } elseif ($fw === 'normal' || (is_numeric($fw) && (int) $fw < 600)) {
                $style->getFont()->setBold(false);
            }
        }

        if (isset($css['font-style'])) {
            $fs = strtolower(trim($css['font-style']));
            if ($fs === 'italic' || $fs === 'oblique') {
                $style->getFont()->setItalic(true);
            }
        }

        if (isset($css['text-decoration'])) {
            $td = strtolower(trim($css['text-decoration']));
            if (strpos($td, 'underline') !== false) {
                $style->getFont()->setUnderline(Font::UNDERLINE_SINGLE);
            }
            if (strpos($td, 'line-through') !== false) {
                $style->getFont()->setStrikethrough(true);
            }
        }

        if (isset($css['font-size'])) {
            $size = $this->parseFontSize($css['font-size']);
            if ($size !== null) {
                $style->getFont()->setSize($size);
            }
        }

        if (isset($css['font-family'])) {
            $family = trim(explode(',', $css['font-family'])[0], " '\"");
            if ($family !== '') {
                $style->getFont()->setName($family);
            }
        }

        // Text wrap
        if (isset($css['white-space']) && in_array(strtolower($css['white-space']), ['normal', 'pre-wrap', 'wrap'], true)) {
            $style->getAlignment()->setWrapText(true);
        }

        // Check child formatting tags (<b>, <strong>, <i>, <em>, <u>, <s>)
        if ($cellNode->getElementsByTagName('b')->length > 0 || $cellNode->getElementsByTagName('strong')->length > 0) {
            $style->getFont()->setBold(true);
        }
        if ($cellNode->getElementsByTagName('i')->length > 0 || $cellNode->getElementsByTagName('em')->length > 0) {
            $style->getFont()->setItalic(true);
        }
        if ($cellNode->getElementsByTagName('u')->length > 0 || $cellNode->getElementsByTagName('ins')->length > 0) {
            $style->getFont()->setUnderline(Font::UNDERLINE_SINGLE);
        }
        if ($cellNode->getElementsByTagName('s')->length > 0 || $cellNode->getElementsByTagName('del')->length > 0 || $cellNode->getElementsByTagName('strike')->length > 0) {
            $style->getFont()->setStrikethrough(true);
        }

        // Borders
        $this->applyBorders($cellNode, $css, $style);

        // Column width attribute/style
        $width = $cellNode->getAttribute('width') ?: ($css['width'] ?? null);
        if ($width && $colspan === 1) {
            $numWidth = (float) filter_var($width, FILTER_SANITIZE_NUMBER_FLOAT, FILTER_FLAG_ALLOW_FRACTION);
            if ($numWidth > 0) {
                // Approximate conversion px/points to Excel column width
                $excelWidth = (strpos($width, 'px') !== false) ? $numWidth / 7.5 : $numWidth;
                $colLetter = Coordinate::stringFromColumnIndex($col);
                $sheet->getColumnDimension($colLetter)->setWidth(max(8, min(100, $excelWidth)));
            }
        }
    }

    /**
     * Apply row styling such as height.
     *
     * @param DOMElement $rowNode
     * @param Worksheet $sheet
     * @param int $row
     */
    protected function applyRowStyle(DOMElement $rowNode, Worksheet $sheet, int $row): void
    {
        $inlineStyle = $rowNode->getAttribute('style');
        $css = $this->parseCssString($inlineStyle);
        $height = $rowNode->getAttribute('height') ?: ($css['height'] ?? null);

        if ($height) {
            $numHeight = (float) filter_var($height, FILTER_SANITIZE_NUMBER_FLOAT, FILTER_FLAG_ALLOW_FRACTION);
            if ($numHeight > 0) {
                $excelHeight = (strpos($height, 'px') !== false) ? $numHeight * 0.75 : $numHeight;
                $sheet->getRowDimension($row)->setRowHeight($excelHeight);
            }
        }
    }

    /**
     * Apply border styles from HTML attributes or CSS.
     *
     * @param DOMElement $cellNode
     * @param array $css
     * @param \PhpOffice\PhpSpreadsheet\Style\Style $style
     */
    protected function applyBorders(DOMElement $cellNode, array $css, $style): void
    {
        $tableNode = $cellNode->parentNode;
        while ($tableNode && strtolower($tableNode->nodeName) !== 'table') {
            $tableNode = $tableNode->parentNode;
        }

        $tableBorder = ($tableNode instanceof DOMElement) ? (int) $tableNode->getAttribute('border') : 0;
        $cellBorder = (int) $cellNode->getAttribute('border');

        if ($tableBorder > 0 || $cellBorder > 0) {
            $style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        }

        // CSS borders
        if (isset($css['border'])) {
            $this->parseAndSetBorder($css['border'], $style->getBorders()->getAllBorders());
        }
        if (isset($css['border-top'])) {
            $this->parseAndSetBorder($css['border-top'], $style->getBorders()->getTop());
        }
        if (isset($css['border-bottom'])) {
            $this->parseAndSetBorder($css['border-bottom'], $style->getBorders()->getBottom());
        }
        if (isset($css['border-left'])) {
            $this->parseAndSetBorder($css['border-left'], $style->getBorders()->getLeft());
        }
        if (isset($css['border-right'])) {
            $this->parseAndSetBorder($css['border-right'], $style->getBorders()->getRight());
        }
    }

    /**
     * Helper to parse CSS border rule and apply to PhpSpreadsheet border object.
     *
     * @param string $borderRule E.g. "1px solid #000" or "none"
     * @param Border $border
     */
    protected function parseAndSetBorder(string $borderRule, Border $border): void
    {
        $rule = strtolower(trim($borderRule));
        if ($rule === 'none' || $rule === '0' || $rule === '0px') {
            $border->setBorderStyle(Border::BORDER_NONE);
            return;
        }

        $parts = preg_split('/\s+/', $rule);
        $borderStyle = Border::BORDER_THIN;
        $color = '000000';

        foreach ($parts as $part) {
            if (in_array($part, ['solid', 'thin'], true)) {
                $borderStyle = Border::BORDER_THIN;
            } elseif (in_array($part, ['dashed'], true)) {
                $borderStyle = Border::BORDER_DASHED;
            } elseif (in_array($part, ['dotted'], true)) {
                $borderStyle = Border::BORDER_DOTTED;
            } elseif (in_array($part, ['double'], true)) {
                $borderStyle = Border::BORDER_DOUBLE;
            } elseif (in_array($part, ['thick', 'medium'], true)) {
                $borderStyle = Border::BORDER_MEDIUM;
            } else {
                $hex = $this->normalizeColorToHex($part);
                if ($hex) {
                    $color = $hex;
                }
            }
        }

        $border->setBorderStyle($borderStyle);
        $border->getColor()->setARGB('FF' . $color);
    }

    /**
     * Parse CSS font-size into points (pt).
     *
     * @param string $fontSize
     * @return float|null
     */
    protected function parseFontSize(string $fontSize): ?float
    {
        $fontSize = strtolower(trim($fontSize));
        if (strpos($fontSize, 'pt') !== false) {
            return (float) str_replace('pt', '', $fontSize);
        }
        if (strpos($fontSize, 'px') !== false) {
            // 1px approx = 0.75pt
            return (float) str_replace('px', '', $fontSize) * 0.75;
        }
        if (strpos($fontSize, 'em') !== false || strpos($fontSize, 'rem') !== false) {
            $scale = (float) str_replace(['rem', 'em'], '', $fontSize);
            return $this->defaultFontSize * $scale;
        }
        if (is_numeric($fontSize)) {
            return (float) $fontSize;
        }
        return null;
    }

    /**
     * Parse CSS style string into key-value array.
     *
     * @param string $styleString
     * @return array
     */
    protected function parseCssString(string $styleString): array
    {
        $styles = [];
        if (empty(trim($styleString))) {
            return $styles;
        }

        $rules = explode(';', $styleString);
        foreach ($rules as $rule) {
            if (strpos($rule, ':') !== false) {
                [$key, $value] = explode(':', $rule, 2);
                $styles[strtolower(trim($key))] = trim($value);
            }
        }
        return $styles;
    }

    /**
     * Normalize various color representations (hex, rgb, named color) to 6-char hex RGB string.
     *
     * @param string $color
     * @return string|null Hex representation like "FF0000" or null
     */
    protected function normalizeColorToHex(string $color): ?string
    {
        $color = strtolower(trim($color));
        if ($color === '' || $color === 'transparent' || $color === 'inherit') {
            return null;
        }

        // Named colors
        if (isset(self::$namedColors[$color])) {
            return self::$namedColors[$color];
        }

        // Hex color (#RGB or #RRGGBB)
        if (strpos($color, '#') === 0) {
            $hex = substr($color, 1);
            if (strlen($hex) === 3) {
                return $hex[0] . $hex[0] . $hex[1] . $hex[1] . $hex[2] . $hex[2];
            }
            if (strlen($hex) === 6) {
                return strtoupper($hex);
            }
            return null;
        }

        // rgb(r, g, b) or rgba(r, g, b, a)
        if (preg_match('/rgba?\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/i', $color, $matches)) {
            $r = sprintf('%02X', min(255, (int) $matches[1]));
            $g = sprintf('%02X', min(255, (int) $matches[2]));
            $b = sprintf('%02X', min(255, (int) $matches[3]));
            return $r . $g . $b;
        }

        return null;
    }

    /**
     * Determine a friendly title for worksheet based on table attributes.
     *
     * @param DOMElement $tableNode
     * @param int $tableIndex
     * @return string
     */
    protected function determineSheetTitle(DOMElement $tableNode, int $tableIndex): string
    {
        if ($tableNode->hasAttribute('data-sheet-name')) {
            $title = trim($tableNode->getAttribute('data-sheet-name'));
            if ($title !== '') {
                return $title;
            }
        }

        // Check for <caption> tag
        $captions = $tableNode->getElementsByTagName('caption');
        if ($captions->length > 0) {
            $captionText = trim($captions->item(0)->textContent);
            if ($captionText !== '') {
                return $captionText;
            }
        }

        if ($tableNode->hasAttribute('id')) {
            $id = trim($tableNode->getAttribute('id'));
            if ($id !== '') {
                return $id;
            }
        }

        if ($tableNode->hasAttribute('name')) {
            $name = trim($tableNode->getAttribute('name'));
            if ($name !== '') {
                return $name;
            }
        }

        return 'Table ' . $tableIndex;
    }

    /**
     * Set worksheet title safely according to Excel restrictions (max 31 chars, no invalid chars).
     *
     * @param Worksheet $sheet
     * @param string $title
     */
    protected function setSafeSheetTitle(Worksheet $sheet, string $title): void
    {
        // Replace invalid Excel sheet characters: \ / ? * : [ ]
        $safeTitle = preg_replace('/[\\\\\\/\?\*\:\[\]]/', '_', $title);
        $safeTitle = mb_substr($safeTitle, 0, 31, 'UTF-8');
        if (trim($safeTitle) === '') {
            $safeTitle = 'Sheet';
        }

        // Ensure uniqueness
        $parent = $sheet->getParent();
        if ($parent) {
            $existingTitles = [];
            foreach ($parent->getAllSheets() as $otherSheet) {
                if ($otherSheet !== $sheet) {
                    $existingTitles[] = $otherSheet->getTitle();
                }
            }

            $uniqueTitle = $safeTitle;
            $counter = 1;
            while (in_array($uniqueTitle, $existingTitles, true)) {
                $suffix = " ({$counter})";
                $maxLen = 31 - mb_strlen($suffix, 'UTF-8');
                $uniqueTitle = mb_substr($safeTitle, 0, $maxLen, 'UTF-8') . $suffix;
                $counter++;
            }
            $safeTitle = $uniqueTitle;
        }

        $sheet->setTitle($safeTitle);
    }

    /**
     * Apply auto-size to all used columns of a worksheet.
     *
     * @param Worksheet $sheet
     */
    protected function applyAutoColumnSizes(Worksheet $sheet): void
    {
        $highestColumn = $sheet->getHighestColumn();
        $highestColIndex = Coordinate::columnIndexFromString($highestColumn);

        for ($col = 1; $col <= $highestColIndex; $col++) {
            $colLetter = Coordinate::stringFromColumnIndex($col);
            $dim = $sheet->getColumnDimension($colLetter);
            // Only enable auto-size if custom width wasn't already explicitly set
            if ($dim->getWidth() < 0) {
                $dim->setAutoSize(true);
            }
        }
    }

    /**
     * Creates and loads DOMDocument handling UTF-8 encoding properly.
     *
     * @param string $html
     * @return DOMDocument
     */
    protected function createDomDocument(string $html): DOMDocument
    {
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->preserveWhiteSpace = false;
        $dom->formatOutput = true;

        // Suppress HTML5 / malformed HTML warnings
        $internalErrors = libxml_use_internal_errors(true);

        // Prepend meta charset or XML encoding tag to ensure UTF-8 handling in DOMDocument
        $htmlWithEncoding = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"/></head><body>' . $html . '</body></html>';

        $dom->loadHTML($htmlWithEncoding, LIBXML_HTML_NOIMPLIED | LIBXML_HTML_NODEFDTD);

        libxml_clear_errors();
        libxml_use_internal_errors($internalErrors);

        return $dom;
    }
}

<?php

namespace sberdyug\htmltoexcel\components;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use sberdyug\htmltoexcel\services\HtmlTableToExcelParser;
use yii\base\Component;
use yii\base\InvalidConfigException;
use yii\helpers\FileHelper;
use yii\web\Response;

/**
 * HtmlToExcel Application Component for Yii2
 * 
 * Usage example:
 * ```php
 * // In config/web.php components:
 * 'htmlToExcel' => [
 *     'class' => \yii2\htmltoexcel\components\HtmlToExcel::class,
 *     'parseStyles' => true,
 *     'autoSizeColumns' => true,
 * ],
 * 
 * // In controller or service:
 * $html = '<table><tr><th>Name</th><th>Email</th></tr><tr><td>John</td><td>john@example.com</td></tr></table>';
 * return Yii::$app->htmlToExcel->send($html, 'users.xlsx');
 * ```
 */
class HtmlToExcel extends Component
{
    /**
     * @var bool Whether to create a separate sheet for each table in HTML.
     */
    public $separateSheetsForTables = true;

    /**
     * @var bool Whether to parse and apply inline CSS styles.
     */
    public $parseStyles = true;

    /**
     * @var bool Whether to auto-size columns based on cell contents.
     */
    public $autoSizeColumns = true;

    /**
     * @var bool Whether to automatically detect numbers, formulas, and booleans.
     */
    public $detectDataTypes = true;

    /**
     * @var string Default font family.
     */
    public $defaultFontFamily = 'Calibri';

    /**
     * @var float Default font size in points.
     */
    public $defaultFontSize = 11.0;

    /**
     * @var string Default output writer type ('Xlsx', 'Xls', 'Csv', 'Ods', 'Html', 'Pdf').
     */
    public $defaultWriterType = 'Xlsx';

    /**
     * @var array Content type map for file downloads
     */
    protected static $mimeMap = [
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'xls' => 'application/vnd.ms-excel',
        'csv' => 'text/csv',
        'ods' => 'application/vnd.oasis.opendocument.spreadsheet',
        'html' => 'text/html',
        'pdf' => 'application/pdf',
    ];

    /**
     * Create a new instance of HtmlTableToExcelParser with current component configuration.
     *
     * @return HtmlTableToExcelParser
     */
    public function createParser(): HtmlTableToExcelParser
    {
        return new HtmlTableToExcelParser([
            'separateSheetsForTables' => $this->separateSheetsForTables,
            'parseStyles' => $this->parseStyles,
            'autoSizeColumns' => $this->autoSizeColumns,
            'detectDataTypes' => $this->detectDataTypes,
            'defaultFontFamily' => $this->defaultFontFamily,
            'defaultFontSize' => $this->defaultFontSize,
        ]);
    }

    /**
     * Parse HTML string to PhpSpreadsheet Spreadsheet.
     *
     * @param string $html
     * @param Spreadsheet|null $existingSpreadsheet
     * @return Spreadsheet
     */
    public function parse(string $html, ?Spreadsheet $existingSpreadsheet = null): Spreadsheet
    {
        return $this->createParser()->parse($html, $existingSpreadsheet);
    }

    /**
     * Parse HTML file to PhpSpreadsheet Spreadsheet.
     *
     * @param string $filePath
     * @param Spreadsheet|null $existingSpreadsheet
     * @return Spreadsheet
     */
    public function parseFile(string $filePath, ?Spreadsheet $existingSpreadsheet = null): Spreadsheet
    {
        return $this->createParser()->parseFile($filePath, $existingSpreadsheet);
    }

    /**
     * Convert HTML table to an Excel file and save to disk.
     *
     * @param string $html HTML containing table(s)
     * @param string $outputPath Target file path (e.g. '@runtime/export.xlsx')
     * @param string|null $writerType 'Xlsx', 'Xls', 'Csv', 'Ods', etc. (default is component's defaultWriterType)
     * @return bool
     */
    public function save(string $html, string $outputPath, ?string $writerType = null): bool
    {
        $realPath = \Yii::getAlias($outputPath);
        $directory = dirname($realPath);
        
        if (!is_dir($directory)) {
            FileHelper::createDirectory($directory);
        }

        $type = $writerType ?? $this->detectWriterTypeFromPath($realPath);
        $spreadsheet = $this->parse($html);
        $writer = IOFactory::createWriter($spreadsheet, $type);
        $writer->save($realPath);

        return file_exists($realPath);
    }

    /**
     * Convert HTML table to Excel and send directly as download response in Yii2.
     *
     * @param string $html HTML containing table(s)
     * @param string $filename Download filename (e.g. 'report.xlsx')
     * @param string|null $writerType 'Xlsx', 'Xls', 'Csv', 'Ods', etc.
     * @param Response|null $response Optional Yii response object
     * @return Response
     */
    public function send(string $html, string $filename = 'export.xlsx', ?string $writerType = null, ?Response $response = null): Response
    {
        $response = $response ?? \Yii::$app->response;
        $type = $writerType ?? $this->detectWriterTypeFromPath($filename);
        $extension = strtolower(pathinfo($filename, PATHINFO_EXTENSION)) ?: 'xlsx';
        $mimeType = self::$mimeMap[$extension] ?? 'application/octet-stream';

        $spreadsheet = $this->parse($html);

        // Create in-memory output stream
        $tempStream = fopen('php://temp', 'r+');
        $writer = IOFactory::createWriter($spreadsheet, $type);
        $writer->save($tempStream);
        rewind($tempStream);

        return $response->sendStreamAsFile($tempStream, $filename, [
            'mimeType' => $mimeType,
            'inline' => false,
        ]);
    }

    /**
     * Detect Writer type from filename extension or default.
     *
     * @param string $filename
     * @return string
     */
    public function detectWriterTypeFromPath(string $filename): string
    {
        $ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        switch ($ext) {
            case 'xls':
                return 'Xls';
            case 'csv':
                return 'Csv';
            case 'ods':
                return 'Ods';
            case 'html':
            case 'htm':
                return 'Html';
            case 'pdf':
                return 'Pdf';
            case 'xlsx':
            default:
                return $this->defaultWriterType;
        }
    }
}

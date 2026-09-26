<?php

namespace sberdyug\htmltoexcel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use sberdyug\htmltoexcel\components\HtmlToExcel;
use sberdyug\htmltoexcel\services\HtmlTableToExcelParser;
use yii\base\Module as BaseModule;
use yii\web\Response;

/**
 * HtmlToExcel Module for Yii2
 * 
 * To use this module, add it to your application configuration:
 * 
 * ```php
 * 'modules' => [
 *     'htmltoexcel' => [
 *         'class' => \yii2\htmltoexcel\Module::class,
 *         'separateSheetsForTables' => true,
 *         'parseStyles' => true,
 *         'autoSizeColumns' => true,
 *         'detectDataTypes' => true,
 *     ],
 * ],
 * ```
 */
class Module extends BaseModule
{
    /**
     * @inheritdoc
     */
    public $controllerNamespace = 'yii2\htmltoexcel\controllers';

    /**
     * @var bool Whether to create a separate sheet for each table in HTML.
     */
    public $separateSheetsForTables = true;

    /**
     * @var bool Whether to parse and apply inline CSS and HTML styles.
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
     * @var HtmlToExcel|null
     */
    protected $_component;

    /**
     * @inheritdoc
     */
    public function init()
    {
        parent::init();

        $this->_component = new HtmlToExcel([
            'separateSheetsForTables' => $this->separateSheetsForTables,
            'parseStyles' => $this->parseStyles,
            'autoSizeColumns' => $this->autoSizeColumns,
            'detectDataTypes' => $this->detectDataTypes,
            'defaultFontFamily' => $this->defaultFontFamily,
            'defaultFontSize' => $this->defaultFontSize,
            'defaultWriterType' => $this->defaultWriterType,
        ]);
    }

    /**
     * Parse HTML string to PhpSpreadsheet Spreadsheet.
     *
     * @param string $html
     * @return Spreadsheet
     */
    public function parse(string $html): Spreadsheet
    {
        return $this->_component->parse($html);
    }

    /**
     * Parse HTML file to PhpSpreadsheet Spreadsheet.
     *
     * @param string $filePath
     * @return Spreadsheet
     */
    public function parseFile(string $filePath): Spreadsheet
    {
        return $this->_component->parseFile($filePath);
    }

    /**
     * Save HTML table(s) to Excel file.
     *
     * @param string $html
     * @param string $outputPath
     * @param string|null $writerType
     * @return bool
     */
    public function save(string $html, string $outputPath, ?string $writerType = null): bool
    {
        return $this->_component->save($html, $outputPath, $writerType);
    }

    /**
     * Send HTML table(s) as Excel download response.
     *
     * @param string $html
     * @param string $filename
     * @param string|null $writerType
     * @param Response|null $response
     * @return Response
     */
    public function send(string $html, string $filename = 'export.xlsx', ?string $writerType = null, ?Response $response = null): Response
    {
        return $this->_component->send($html, $filename, $writerType, $response);
    }

    /**
     * Get underlying HtmlToExcel component instance.
     *
     * @return HtmlToExcel
     */
    public function getComponent(): HtmlToExcel
    {
        return $this->_component;
    }
}

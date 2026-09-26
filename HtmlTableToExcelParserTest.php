<?php

namespace sberdyug\htmltoexcel\tests;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PHPUnit\Framework\TestCase;
use sberdyug\htmltoexcel\components\HtmlToExcel;
use sberdyug\htmltoexcel\Module;
use sberdyug\htmltoexcel\services\HtmlTableToExcelParser;

class HtmlTableToExcelParserTest extends TestCase
{
    protected function tearDown(): void
    {
        parent::tearDown();
        $runtime = \Yii::getAlias('@runtime');
        if (is_dir($runtime)) {
            $files = glob($runtime . '/*');
            foreach ($files as $file) {
                if (is_file($file)) {
                    unlink($file);
                }
            }
        }
    }

    public function testBasicTableParsing()
    {
        $html = <<<HTML
<table>
    <thead>
        <tr>
            <th>Product</th>
            <th>Price</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>Apple</td>
            <td>100</td>
        </tr>
        <tr>
            <td>Banana</td>
            <td>200.5</td>
        </tr>
    </tbody>
</table>
HTML;

        $parser = new HtmlTableToExcelParser();
        $spreadsheet = $parser->parse($html);
        $sheet = $spreadsheet->getActiveSheet();

        $this->assertEquals('Product', $sheet->getCell('A1')->getValue());
        $this->assertEquals('Price', $sheet->getCell('B1')->getValue());
        $this->assertTrue($sheet->getStyle('A1')->getFont()->getBold());
        $this->assertTrue($sheet->getStyle('B1')->getFont()->getBold());

        $this->assertEquals('Apple', $sheet->getCell('A2')->getValue());
        $this->assertEquals(100, $sheet->getCell('B2')->getValue());
        $this->assertEquals(DataType::TYPE_NUMERIC, $sheet->getCell('B2')->getDataType());

        $this->assertEquals('Banana', $sheet->getCell('A3')->getValue());
        $this->assertEquals(200.5, $sheet->getCell('B3')->getValue());
    }

    public function testColspanAndRowspan()
    {
        $html = <<<HTML
<table>
    <tr>
        <th rowspan="2">Group</th>
        <th colspan="2">Details</th>
    </tr>
    <tr>
        <th>Sub 1</th>
        <th>Sub 2</th>
    </tr>
    <tr>
        <td>A</td>
        <td>10</td>
        <td>20</td>
    </tr>
</table>
HTML;

        $parser = new HtmlTableToExcelParser();
        $spreadsheet = $parser->parse($html);
        $sheet = $spreadsheet->getActiveSheet();

        // Row 1: A1 (rowspan 2) -> spans A1:A2, B1 (colspan 2) -> spans B1:C1
        $this->assertEquals('Group', $sheet->getCell('A1')->getValue());
        $this->assertEquals('Details', $sheet->getCell('B1')->getValue());

        // Merges check
        $merges = $sheet->getMergeCells();
        $this->assertArrayHasKey('A1:A2', $merges);
        $this->assertArrayHasKey('B1:C1', $merges);

        // Row 2: Sub 1 must be placed at B2, Sub 2 at C2 (since A2 is occupied by A1's rowspan)
        $this->assertEquals('Sub 1', $sheet->getCell('B2')->getValue());
        $this->assertEquals('Sub 2', $sheet->getCell('C2')->getValue());

        // Row 3: A3, B3, C3
        $this->assertEquals('A', $sheet->getCell('A3')->getValue());
        $this->assertEquals(10, $sheet->getCell('B3')->getValue());
        $this->assertEquals(20, $sheet->getCell('C3')->getValue());
    }

    public function testMultipleTablesSeparateSheets()
    {
        $html = <<<HTML
<table data-sheet-name="Users">
    <tr><th>User</th></tr>
    <tr><td>Alice</td></tr>
</table>
<table data-sheet-name="Orders">
    <tr><th>Order ID</th></tr>
    <tr><td>#1001</td></tr>
</table>
HTML;

        $parser = new HtmlTableToExcelParser(['separateSheetsForTables' => true]);
        $spreadsheet = $parser->parse($html);

        $this->assertEquals(2, $spreadsheet->getSheetCount());
        $this->assertEquals('Users', $spreadsheet->getSheet(0)->getTitle());
        $this->assertEquals('Orders', $spreadsheet->getSheet(1)->getTitle());

        $this->assertEquals('Alice', $spreadsheet->getSheet(0)->getCell('A2')->getValue());
        $this->assertEquals('#1001', $spreadsheet->getSheet(1)->getCell('A2')->getValue());
    }

    public function testMultipleTablesStackedOnSingleSheet()
    {
        $html = <<<HTML
<table>
    <tr><th>Table 1</th></tr>
    <tr><td>Data 1</td></tr>
</table>
<table>
    <tr><th>Table 2</th></tr>
    <tr><td>Data 2</td></tr>
</table>
HTML;

        $parser = new HtmlTableToExcelParser(['separateSheetsForTables' => false, 'tableSpacing' => 2]);
        $spreadsheet = $parser->parse($html);

        $this->assertEquals(1, $spreadsheet->getSheetCount());
        $sheet = $spreadsheet->getActiveSheet();

        // Table 1 at row 1-2
        $this->assertEquals('Table 1', $sheet->getCell('A1')->getValue());
        $this->assertEquals('Data 1', $sheet->getCell('A2')->getValue());

        // Row 3 and 4 are blank, Table 2 at row 5-6
        $this->assertEquals('Table 2', $sheet->getCell('A5')->getValue());
        $this->assertEquals('Data 2', $sheet->getCell('A6')->getValue());
    }

    public function testInlineStylesAndColors()
    {
        $html = <<<HTML
<table>
    <tr>
        <td style="background-color: #FF0000; color: #FFFFFF; font-weight: bold; text-align: center; font-style: italic;">
            Styled Cell
        </td>
        <td align="right" valign="top" bgcolor="#00FF00">
            Attr Cell
        </td>
    </tr>
</table>
HTML;

        $parser = new HtmlTableToExcelParser();
        $spreadsheet = $parser->parse($html);
        $sheet = $spreadsheet->getActiveSheet();

        $style1 = $sheet->getStyle('A1');
        $this->assertTrue($style1->getFont()->getBold());
        $this->assertTrue($style1->getFont()->getItalic());
        $this->assertEquals(Alignment::HORIZONTAL_CENTER, $style1->getAlignment()->getHorizontal());
        $this->assertEquals('FFFF0000', $style1->getFill()->getStartColor()->getARGB());
        $this->assertEquals('FFFFFFFF', $style1->getFont()->getColor()->getARGB());

        $style2 = $sheet->getStyle('B1');
        $this->assertEquals(Alignment::HORIZONTAL_RIGHT, $style2->getAlignment()->getHorizontal());
        $this->assertEquals(Alignment::VERTICAL_TOP, $style2->getAlignment()->getVertical());
        $this->assertEquals('FF00FF00', $style2->getFill()->getStartColor()->getARGB());
    }

    public function testFormulasAndLinks()
    {
        $html = <<<HTML
<table>
    <tr>
        <td>10</td>
        <td>20</td>
        <td>=SUM(A1:B1)</td>
        <td><a href="https://yiiframework.com" title="Yii2 Official">Yii Framework</a></td>
    </tr>
</table>
HTML;

        $parser = new HtmlTableToExcelParser();
        $spreadsheet = $parser->parse($html);
        $sheet = $spreadsheet->getActiveSheet();

        $this->assertEquals(10, $sheet->getCell('A1')->getValue());
        $this->assertEquals(20, $sheet->getCell('B1')->getValue());
        $this->assertEquals('=SUM(A1:B1)', $sheet->getCell('C1')->getValue());
        $this->assertEquals(DataType::TYPE_FORMULA, $sheet->getCell('C1')->getDataType());

        $cellD1 = $sheet->getCell('D1');
        $this->assertEquals('Yii Framework', $cellD1->getValue());
        $this->assertEquals('https://yiiframework.com', $cellD1->getHyperlink()->getUrl());
        $this->assertEquals('Yii2 Official', $cellD1->getHyperlink()->getTooltip());
    }

    public function testSaveToFile()
    {
        $html = '<table><tr><th>ID</th><th>Name</th></tr><tr><td>1</td><td>Test</td></tr></table>';
        $component = new HtmlToExcel();

        $targetPath = \Yii::getAlias('@runtime/test_output.xlsx');
        $saved = $component->save($html, $targetPath);

        $this->assertTrue($saved);
        $this->assertFileExists($targetPath);
        $this->assertGreaterThan(100, filesize($targetPath));
    }

    public function testModuleExecution()
    {
        $module = new Module('htmltoexcel');
        $html = '<table><tr><td>Cell 1</td></tr></table>';
        $spreadsheet = $module->parse($html);
        $sheet = $spreadsheet->getActiveSheet();

        $this->assertEquals('Cell 1', $sheet->getCell('A1')->getValue());
    }
}

<?php

namespace sberdyug\htmltoexcel\widgets;

use yii\base\Widget;
use yii\helpers\ArrayHelper;
use yii\helpers\Html;
use yii\helpers\Json;
use yii\helpers\Url;

/**
 * TableExportWidget renders a button that grabs HTML table content from the page
 * and sends it to the server to download as an Excel file.
 * 
 * Usage:
 * ```php
 * <?= \yii2\htmltoexcel\widgets\TableExportWidget::widget([
 *     'tableSelector' => '#sales-table', // or '.grid-view table'
 *     'filename' => 'sales_report.xlsx',
 *     'buttonLabel' => '📥 Экспорт в Excel',
 *     'buttonOptions' => ['class' => 'btn btn-success'],
 * ]) ?>
 * ```
 */
class TableExportWidget extends Widget
{
    /**
     * @var string CSS selector of the table to export (e.g. '#my-table', '.grid-view table').
     */
    public $tableSelector;

    /**
     * @var string Output filename.
     */
    public $filename = 'export.xlsx';

    /**
     * @var string Export format ('Xlsx', 'Xls', 'Csv', 'Ods', 'Html').
     */
    public $format = 'Xlsx';

    /**
     * @var string|array Target export route.
     */
    public $exportRoute = ['/htmltoexcel/export/html'];

    /**
     * @var string Button label.
     */
    public $buttonLabel = 'Экспорт в Excel';

    /**
     * @var array HTML options for the button.
     */
    public $buttonOptions = ['class' => 'btn btn-success'];

    /**
     * @inheritdoc
     */
    public function run()
    {
        $id = $this->getId();
        $buttonOptions = $this->buttonOptions;
        $buttonOptions['id'] = $id . '-btn';

        $exportUrl = Url::to($this->exportRoute);
        $csrfParam = \Yii::$app->request->csrfParam;
        $csrfToken = \Yii::$app->request->getCsrfToken();

        $config = Json::encode([
            'buttonId' => $buttonOptions['id'],
            'tableSelector' => $this->tableSelector,
            'filename' => $this->filename,
            'format' => $this->format,
            'exportUrl' => $exportUrl,
            'csrfParam' => $csrfParam,
            'csrfToken' => $csrfToken,
        ]);

        $this->registerClientScript($config);

        return Html::button($this->buttonLabel, $buttonOptions);
    }

    /**
     * Register JavaScript for table grab & download.
     *
     * @param string $configJson
     */
    protected function registerClientScript(string $configJson): void
    {
        $view = $this->getView();
        $js = <<<JS
(function(cfg) {
    var btn = document.getElementById(cfg.buttonId);
    if (!btn) return;

    btn.addEventListener('click', function(e) {
        e.preventDefault();
        var tableHtml = '';
        if (cfg.tableSelector) {
            var tables = document.querySelectorAll(cfg.tableSelector);
            if (tables.length === 0) {
                alert('Table element not found for selector: ' + cfg.tableSelector);
                return;
            }
            tables.forEach(function(tbl) {
                tableHtml += tbl.outerHTML;
            });
        }

        // Dynamically create a hidden form and submit
        var form = document.createElement('form');
        form.method = 'POST';
        form.action = cfg.exportUrl;
        form.style.display = 'none';

        var appendField = function(name, val) {
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = name;
            input.value = val;
            form.appendChild(input);
        };

        if (cfg.csrfParam && cfg.csrfToken) {
            appendField(cfg.csrfParam, cfg.csrfToken);
        }
        appendField('html', tableHtml);
        appendField('filename', cfg.filename);
        appendField('format', cfg.format);

        document.body.appendChild(form);
        form.submit();
        setTimeout(function() {
            document.body.removeChild(form);
        }, 1000);
    });
})({$configJson});
JS;
        $view->registerJs($js);
    }
}

<?php

use yii\helpers\Html;

/** @var yii\web\View $this */
/** @var string $html */
/** @var string $filename */
/** @var string $format */

$this->title = 'HTML Table to Excel Export Demo';
?>
<div class="html-to-excel-default-index" style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 1200px; margin: 20px auto; padding: 20px;">
    <h1><?= Html::encode($this->title) ?></h1>
    <p>Модуль преобразует любые HTML-таблицы в документы Excel (с сохранением стилей, цветов, границ, объединений colspan/rowspan и формул) с помощью <strong>PhpSpreadsheet</strong>.</p>

    <?= Html::beginForm(['index'], 'post', ['id' => 'export-form']) ?>
        <div style="display: flex; gap: 20px; margin-bottom: 15px;">
            <div style="flex: 2;">
                <label for="filename" style="font-weight: bold; display: block; margin-bottom: 5px;">Имя файла:</label>
                <?= Html::input('text', 'filename', $filename, ['id' => 'filename', 'class' => 'form-control', 'style' => 'width: 100%; padding: 8px; box-sizing: border-box;']) ?>
            </div>
            <div style="flex: 1;">
                <label for="format" style="font-weight: bold; display: block; margin-bottom: 5px;">Формат:</label>
                <?= Html::dropDownList('format', $format, [
                    'Xlsx' => 'Excel (.xlsx)',
                    'Xls' => 'Excel 97-2003 (.xls)',
                    'Csv' => 'CSV (.csv)',
                    'Ods' => 'OpenDocument (.ods)',
                    'Html' => 'HTML (.html)',
                ], ['id' => 'format', 'class' => 'form-control', 'style' => 'width: 100%; padding: 8px; box-sizing: border-box;']) ?>
            </div>
        </div>

        <div style="margin-bottom: 15px;">
            <label for="html-input" style="font-weight: bold; display: block; margin-bottom: 5px;">HTML код таблицы (или нескольких таблиц):</label>
            <?= Html::textarea('html', $html, [
                'id' => 'html-input',
                'rows' => 14,
                'class' => 'form-control',
                'style' => 'width: 100%; font-family: monospace; font-size: 13px; padding: 10px; box-sizing: border-box;',
            ]) ?>
        </div>

        <div style="margin-bottom: 25px;">
            <?= Html::hiddenInput('submit_action', 'export', ['id' => 'submit_action']) ?>
            <?= Html::submitButton('📥 Скачать Excel', [
                'class' => 'btn btn-primary',
                'style' => 'background-color: #28a745; color: white; padding: 10px 20px; border: none; border-radius: 4px; font-size: 16px; cursor: pointer;',
            ]) ?>
        </div>
    <?= Html::endForm() ?>

    <hr style="margin: 30px 0; border: 0; border-top: 1px solid #eee;">

    <h3>Предпросмотр HTML:</h3>
    <div style="overflow-x: auto; border: 1px solid #ddd; padding: 15px; border-radius: 4px; background: #fafafa;">
        <?= $html ?>
    </div>
</div>

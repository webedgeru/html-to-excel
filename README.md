# Yii2 HTML Table to Excel (PhpSpreadsheet)

Модуль и компонент для **Yii2 Framework**, предназначенный для быстрого и точного парсинга **HTML-таблиц в таблицы Excel** (`.xlsx`, `.xls`, `.csv`, `.ods`, `.html`, `.pdf`) с использованием библиотеки **PhpSpreadsheet**.

---

## 🚀 Основные возможности

- 📊 **Парсинг любых HTML-таблиц**: `<table>`, `<thead>`, `<tbody>`, `<tfoot>`, `<tr>`, `<th>`, `<td>`.
- 🧩 **Поддержка объединения ячеек**: корректная обработка атрибутов `colspan` и `rowspan` без смещения последующих ячеек.
- 🎨 **Сохранение стилей и оформления**:
  - Жирный, курсив, подчеркивание, зачеркивание (`<b>`, `<strong>`, `<i>`, `<em>`, `<u>`, `<s>`, CSS `font-weight`, `font-style`, `text-decoration`).
  - Цвета текста и фона (CSS `color`, `background-color`, `background`, атрибут `bgcolor` — в форматах HEX, RGB, RGBA и названий цветов).
  - Выравнивание по горизонтали и вертикали (`text-align`, `vertical-align`, `align`, `valign`).
  - Границы ячеек (CSS `border`, `border-top`, `border-left`..., атрибут `border="1"`).
  - Размеры шрифтов и семейство шрифтов (`font-size`, `font-family`).
  - Перенос строк (`<br>` и `white-space: wrap`).
  - Настройка ширины колонок и высоты строк (атрибуты `width`, `height` и CSS-стили).
- 🔢 **Автоопределение типов данных**:
  - Числа (`int`, `float`), формулы (начинающиеся с `=`), логические значения, строки.
  - Поддержка явного указания типа через атрибут `data-type="numeric|formula|string|bool"`.
  - Гиперссылки (`<a href="..." title="...">`).
- 📑 **Несколько таблиц**:
  - Экспорт каждой таблицы на **отдельный лист** с автоматическим или кастомным именем (через `data-sheet-name`, `<caption>` или `id`).
  - Либо размещение всех таблиц последовательно на **одном листе**.
- 📥 **Гибкие сценарии использования**:
  - Мгновенная отдача файла на скачивание пользователю в браузер (`send()`).
  - Сохранение в файл на сервере (`save()`).
  - Получение объекта `PhpOffice\PhpSpreadsheet\Spreadsheet` для дальнейшей ручной модификации (`parse()`).
  - Готовый виджет `TableExportWidget` для экспорта таблиц прямо со страниц сайта в один клик.

---

## 📦 Требования

- PHP `>= 7.4` или `>= 8.0` (с расширениями `dom`, `libxml`, `mbstring`, `zip`, `gd`)
- Yii2 Framework `>= 2.0.0`
- PhpSpreadsheet `^1.18 || ^2.0 || ^3.0`

---

## 🛠 Установка

Установите пакет через Composer:

```bash
composer require yii2mod/yii2-html-to-excel
```

Или добавьте в ваш `composer.json`:

```json
"require": {
    "yii2mod/yii2-html-to-excel": "^1.0"
}
```

---

## ⚙️ Настройка в Yii2

### 1. Как модуль (`config/web.php` или `config/main.php`)

```php
'modules' => [
    'htmltoexcel' => [
        'class' => \yii2\htmltoexcel\Module::class,
        'separateSheetsForTables' => true, // Разделять таблицы по отдельным листам
        'parseStyles' => true,             // Парсить CSS и HTML стили
        'autoSizeColumns' => true,         // Автоподбор ширины колонок
        'detectDataTypes' => true,         // Автоопределение чисел и формул
        'defaultFontFamily' => 'Calibri',  // Шрифт по умолчанию
        'defaultFontSize' => 11.0,         // Размер шрифта по умолчанию
        'defaultWriterType' => 'Xlsx',     // Формат по умолчанию ('Xlsx', 'Xls', 'Csv', 'Ods', 'Html')
    ],
],
```

После подключения модуля доступна страница интерактивного тестирования и демо по адресу:
`http://your-app.test/htmltoexcel` или `http://your-app.test/index.php?r=htmltoexcel/default/index`.

### 2. Как компонент приложения (`config/web.php` или `config/main.php`)

```php
'components' => [
    'htmlToExcel' => [
        'class' => \yii2\htmltoexcel\components\HtmlToExcel::class,
        'separateSheetsForTables' => true,
        'parseStyles' => true,
        'autoSizeColumns' => true,
        'detectDataTypes' => true,
    ],
],
```

---

## 💡 Примеры использования

### 1. Отдача файла на скачивание в браузере (в Controller)

```php
namespace app\controllers;

use Yii;
use yii\web\Controller;

class ReportController extends Controller
{
    public function actionExport()
    {
        $html = <<<HTML
        <table border="1" data-sheet-name="Отчет">
            <thead>
                <tr style="background-color: #2F4F4F; color: #FFFFFF;">
                    <th>ID</th>
                    <th>Клиент</th>
                    <th>Сумма</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>1</td>
                    <td><b>ООО "Компания"</b></td>
                    <td style="text-align: right;" data-type="numeric">15000.50</td>
                </tr>
            </tbody>
        </table>
        HTML;

        // Отправить пользователю на скачивание (через компонент)
        return Yii::$app->htmlToExcel->send($html, 'report.xlsx');
        
        // Либо через модуль:
        // return Yii::$app->getModule('htmltoexcel')->send($html, 'report.xlsx');
    }
}
```

### 2. Сохранение файла на сервере

```php
$html = '<table><tr><td>Данные</td></tr></table>';

// Сохранить в @runtime/exports/my_table.xlsx
Yii::$app->htmlToExcel->save($html, '@runtime/exports/my_table.xlsx');

// Экспорт в формате CSV
Yii::$app->htmlToExcel->save($html, '@runtime/exports/my_table.csv', 'Csv');
```

### 3. Получение объекта `Spreadsheet` для дальнейшей работы

```php
use yii2\htmltoexcel\services\HtmlTableToExcelParser;

$parser = new HtmlTableToExcelParser([
    'parseStyles' => true,
    'autoSizeColumns' => true,
]);

/** @var \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet */
$spreadsheet = $parser->parse($htmlContent);

// Дополнительная кастомизация через PhpSpreadsheet API:
$sheet = $spreadsheet->getActiveSheet();
$sheet->getStyle('A1')->getFont()->setSize(16);

// Сохранение любым писателем PhpSpreadsheet:
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('custom_report.xlsx');
```

### 4. Экспорт с `colspan`, `rowspan`, стилями и формулами

Модуль корректно строит сложную сетку ячеек:

```html
<table border="1" style="border-collapse: collapse;" data-sheet-name="Сводный отчет">
    <thead>
        <tr style="background-color: #337ab7; color: #ffffff;">
            <th rowspan="2" style="text-align: center; vertical-align: middle;">Категория</th>
            <th colspan="2" style="text-align: center;">2025 год</th>
            <th colspan="2" style="text-align: center;">2026 год</th>
        </tr>
        <tr style="background-color: #5bc0de; color: #ffffff;">
            <th>План</th>
            <th>Факт</th>
            <th>План</th>
            <th>Факт</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>Товары</td>
            <td data-type="numeric">1000</td>
            <td data-type="numeric">1200</td>
            <td data-type="numeric">1500</td>
            <td data-type="numeric">1650</td>
        </tr>
        <tr>
            <td>Услуги</td>
            <td data-type="numeric">500</td>
            <td data-type="numeric">480</td>
            <td data-type="numeric">600</td>
            <td data-type="numeric">710</td>
        </tr>
    </tbody>
    <tfoot>
        <tr style="font-weight: bold; background-color: #f5f5f5;">
            <td>ИТОГО:</td>
            <td data-type="formula">=SUM(B3:B4)</td>
            <td data-type="formula">=SUM(C3:C4)</td>
            <td data-type="formula">=SUM(D3:D4)</td>
            <td data-type="formula">=SUM(E3:E4)</td>
        </tr>
    </tfoot>
</table>
```

### 5. Использование виджета на ст��анице (`TableExportWidget`)

Позволяет добавить кнопку экспорта к любой таблице или `GridView`:

```php
use yii2\htmltoexcel\widgets\TableExportWidget;

// Кнопка экспорта таблицы с id="sales-grid"
<?= TableExportWidget::widget([
    'tableSelector' => '#sales-grid table',
    'filename' => 'sales_export.xlsx',
    'format' => 'Xlsx',
    'buttonLabel' => '📥 Экспорт таблицы в Excel',
    'buttonOptions' => ['class' => 'btn btn-success'],
]) ?>

// Сама таблица (GridView или стандартная HTML таблица)
<?= \yii\grid\GridView::widget([
    'id' => 'sales-grid',
    'dataProvider' => $dataProvider,
    'columns' => [
        'id',
        'name',
        'created_at:datetime',
        'amount:currency',
    ],
]) ?>
```

---

## ⚙️ Параметры конфигурации

| Параметр | Тип | По умолчанию | Описание |
|---|---|---|---|
| `separateSheetsForTables` | `bool` | `true` | Если `true`, каждая `<table>` в HTML помещается на отдельный лист Excel. Если `false`, таблицы размещаются друг под другом на одном листе. |
| `parseStyles` | `bool` | `true` | Парсинг inline CSS (`style="..."`) и HTML-атрибутов (`bgcolor`, `align`, `valign`, `border`, `width`, `height`, тегов `<b>`, `<i>`, `<u>`, `<s>`). |
| `autoSizeColumns` | `bool` | `true` | Автоматический подбор ширины колонок под содержимое. |
| `detectDataTypes` | `bool` | `true` | Автоопределение числовых значений и формул (начинающихся с `=`). |
| `defaultFontFamily` | `string` | `'Calibri'` | Семейство шрифтов по умолчанию. |
| `defaultFontSize` | `float` | `11.0` | Размер шрифта по умолчанию (в pt). |
| `defaultWriterType` | `string` | `'Xlsx'` | Формат экспорта по умолчанию (`Xlsx`, `Xls`, `Csv`, `Ods`, `Html`, `Pdf`). |
| `tableSpacing` | `int` | `2` | Количество пустых строк между таблицами при `separateSheetsForTables = false`. |

---

## 🧪 Запуск тестов

Для запуска тестов используйте PHPUnit:

```bash
vendor/bin/phpunit
```

---

## 📄 Лицензия

MIT License.

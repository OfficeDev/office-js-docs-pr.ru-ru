---
title: Работа с диаграммами с использованием API JavaScript для Excel
description: 'Примеры кода, демонстрирующие задачи диаграммы с Excel API JavaScript.'
ms.date: 11/29/2021
ms.localizationpriority: medium
---

# <a name="work-with-charts-using-the-excel-javascript-api"></a>Работа с диаграммами с использованием API JavaScript для Excel

В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для диаграмм с использованием API JavaScript для Excel.
Полный список свойств и методов, поддерживаемых объектами, см. в таблице [Chart Object (API JavaScript для Excel)](/javascript/api/excel/excel.chart) и Объект коллекции диаграмм [(API JavaScript для Excel)](/javascript/api/excel/excel.chartcollection).`Chart` `ChartCollection`

## <a name="create-a-chart"></a>Создание диаграммы

В примере кода ниже показано, как создать диаграмму на листе **Sample** (Пример). Диаграмма представляет собой **график**, построенный на основе данных из диапазона **A1:B13**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var dataRange = sheet.getRange("A1:B13");
    var chart = sheet.charts.add("Line", dataRange, "auto");

    chart.title.text = "Sales Data";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Новый график**

![Новая диаграмма строки в Excel.](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a>Добавление ряда данных в диаграмму

В примере кода ниже показано, как добавить ряд данных в первую диаграмму на листе. Новый ряд данных соответствует столбцу **2016** и основан на данных из диапазона **D2:D5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var chart = sheet.charts.getItemAt(0);
    var dataRange = sheet.getRange("D2:D5");

    var newSeries = chart.series.add("2016");
    newSeries.setValues(dataRange);

    return context.sync();
}).catch(errorHandlerFunction);
```

**Диаграмма перед добавлением ряда данных 2016**

![Диаграмма в Excel до добавленной серии данных за 2016 г.](../images/excel-charts-data-series-before.png)

**Диаграмма после добавления ряда данных 2016**

![Диаграмма в Excel после добавленной серии данных 2016 г.](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a>Задание названия диаграммы

В примере ниже показано, как задать название **Sales Data by Year** (Данные продаж по годам) для первой диаграммы на листе.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Диаграмма после задания заголовка**

![Диаграмма с заголовком в Excel.](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a>Задание свойств оси диаграммы

Диаграммы, в которых используется [декартова система координат](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), например гистограммы, линейчатые и точечные диаграммы, содержат ось категорий и ось значений. В примерах ниже показано, как задать название и отобразить единицу измерения по оси для диаграммы.

### <a name="set-axis-title"></a>Задание названия оси

В примере кода ниже показано, как задать название **Product** (Продукт) для оси категорий первой диаграммы на листе.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Диаграмма после задания названия оси категорий**

![Диаграмма с названием оси в Excel.](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a>Задание отображаемой единицы измерения оси

В примере ниже показано, как задать отображаемую единицу измерения **Hundreds** (Сотни) для оси значений первой диаграммы на листе.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Диаграмма после задания единицы измерения оси значений**

![Диаграмма с блоком отображения оси в Excel.](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>Настройка видимости линий сетки на диаграмме

В примере ниже показано, как скрыть основные линии сетки для оси значений первой диаграммы на листе. Основные линии сетки для оси значения диаграммы можно показать, `chart.axes.valueAxis.majorGridlines.visible` установив значение `true`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Диаграмма со скрытыми линиями сетки**

![Диаграмма с сетками, скрытыми в Excel.](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a>Линии трендов диаграммы

### <a name="add-a-trendline"></a>Добавление линии тренда

В примере кода ниже показано, как добавить линию тренда "скользящее среднее" в первый ряд первой диаграммы на листе **Sample** (Пример). Линия тренда отображает "скользящее среднее" за 5 периодов.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Диаграмма с линией тренда "скользящее среднее"**

![Диаграмма с скользящего среднего тренда в Excel.](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a>Изменение линии тренда

Следующий пример кода задает линию тренда `Linear` для введите для первой серии в первой диаграмме в таблице с именем **Sample**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    var series = seriesCollection.getItemAt(0);
    series.trendlines.getItem(0).type = "Linear";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Диаграмма с линейной линией тренда**

![Диаграмма с линейным трендом в Excel.](../images/excel-charts-trendline-linear.png)

## <a name="add-and-format-a-chart-data-table"></a>Добавление и формат таблицы данных диаграммы

С помощью метода можно получить доступ к элементу таблицы данных диаграммы [`Chart.getDataTableOrNullObject`](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatableornullobject-member(1)) . Этот метод возвращает объект [`ChartDataTable`](/javascript/api/excel/excel.chartdatatable) . Объект `ChartDataTable` имеет свойства форматирования boolean, такие как `visible`, и `showLegendKey``showHorizontalBorder`.

Свойство `ChartDataTable.format` возвращает объект [`ChartDataTableFormat`](/javascript/api/excel/excel.chartdatatableformat) , что позволяет далее форматирование и стиль таблицы данных. Объект `ChartDataTableFormat` предлагает `border`и `fill`свойства `font` .

В следующем примере кода показано, как добавить таблицу данных в диаграмму, а затем форматировать эту таблицу данных с помощью объектов `ChartDataTable` и объектов `ChartDataTableFormat` .

```js
// This code sample adds a data table to a chart that already exists on the worksheet, 
// and then adjusts the display and format of that data table.
Excel.run(function (context) {
    // Retrieve the chart on the "Sample" worksheet.
    var chart = context.workbook.worksheets.getItem("Sample").charts.getItemAt(0);

    // Get the chart data table object and load its properties.
    var chartDataTable = chart.getDataTableOrNullObject();
    chartDataTable.load();

    // Set the display properties of the chart data table.
    chartDataTable.visible = true;
    chartDataTable.showLegendKey = true;
    chartDataTable.showHorizontalBorder = false;
    chartDataTable.showVerticalBorder = true;
    chartDataTable.showOutlineBorder = true;

    // Retrieve the chart data table format object and set font and border properties. 
    var chartDataTableFormat = chartDataTable.format;
    chartDataTableFormat.font.color = "#B76E79";
    chartDataTableFormat.font.name = "Comic Sans";
    chartDataTableFormat.border.color = "blue";

    return context.sync();
}).catch(errorHandlerFunction);
```

На следующем скриншоте показана таблица данных, которую создает предыдущий пример кода.

![Диаграмма со таблицей данных, демонстрация настраиваемого форматирования таблицы данных.](../images/excel-charts-data-table.png)

## <a name="export-a-chart-as-an-image"></a>Экспорт диаграммы как изображения

Диаграммы можно отображать как изображения за пределами Excel. Метод `Chart.getImage` возвращает диаграмму в виде строки в кодировке base64, представляющей диаграмму в формате изображения JPEG. В приведенном ниже коде показано, как получить строку изображения и записать ее в консоли.

```js
Excel.run(function (ctx) {
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    var imageAsString = chart.getImage();
    return context.sync().then(function () {
        console.log(imageAsString.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

Метод `Chart.getImage` использует три дополнительных параметра: ширина, высота и режим подгонки.

```typescript
getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>;
```

Эти параметры определяют размер изображения. Изображения всегда масштабируются пропорционально. Параметры ширины и высоты устанавливают верхние или нижние границы для масштабированного изображения. `ImageFittingMode` имеет три значения со следующими действиями.

- `Fill`: Минимальная высота или ширина изображения — указанная высота или ширина (в зависимости от того, достигается ли она сначала при масштабирования изображения). Это поведение по умолчанию, если не задан параметр режима подгонки.
- `Fit`: Максимальная высота или ширина изображения — указанная высота или ширина (в зависимости от того, достигается ли она сначала при масштабирования изображения).
- `FitAndCenter`: Максимальная высота или ширина изображения — указанная высота или ширина (в зависимости от того, достигается ли она сначала при масштабирования изображения). Получившееся изображение выравнивается по центру относительно другого измерения.

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)

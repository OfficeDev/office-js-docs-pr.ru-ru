---
title: Excel API JavaScript установлено 1.1
description: Сведения о наборе требований ExcelApi 1.1.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 45061afc7e401e18a67377bf88fa1670bb7a8ece
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745957"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel API JavaScript установлено 1.1

Excel API JavaScript 1.1 — это первая версия API. Это единственный набор Excel, поддерживаемый Excel 2016.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.1. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.1 см. Excel API в наборе [требований 1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel. CalculationType)](/javascript/api/excel/excel.application#excel-excel-application-calculate-member(1))|Пересчитывает данные во всех открытых в текущий момент книгах Excel.|
||[calculationMode](/javascript/api/excel/excel.application#excel-excel-application-calculationmode-member)|Возвращает режим вычисления, используемый в книге, как это определено константами в `Excel.CalculationMode`.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#excel-excel-binding-getrange-member(1))|Возвращает представленный привязкой диапазон.|
||[getTable()](/javascript/api/excel/excel.binding#excel-excel-binding-gettable-member(1))|Возвращает представленную привязкой таблицу.|
||[getText()](/javascript/api/excel/excel.binding#excel-excel-binding-gettext-member(1))|Возвращает представленный привязкой текст.|
||[id](/javascript/api/excel/excel.binding#excel-excel-binding-id-member)|Представляет идентификатор привязки.|
||[тип](/javascript/api/excel/excel.binding#excel-excel-binding-type-member)|Возвращает тип привязки.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[count](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-count-member)|Возвращает число привязок в коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitem-member(1))|Возвращает объект привязки по идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemat-member(1))|Возвращает объект привязки с учетом его положения в массиве элементов.|
||[items](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Chart](/javascript/api/excel/excel.chart)|[axes](/javascript/api/excel/excel.chart#excel-excel-chart-axes-member)|Представляет оси диаграммы.|
||[dataLabels](/javascript/api/excel/excel.chart#excel-excel-chart-datalabels-member)|Представляет метки данных на диаграмме.|
||[delete()](/javascript/api/excel/excel.chart#excel-excel-chart-delete-member(1))|Удаляет объект диаграммы.|
||[format](/javascript/api/excel/excel.chart#excel-excel-chart-format-member)|Инкапсулирует свойства формата для области диаграммы.|
||[height](/javascript/api/excel/excel.chart#excel-excel-chart-height-member)|Указывает высоту в точках объекта диаграммы.|
||[left](/javascript/api/excel/excel.chart#excel-excel-chart-left-member)|Расстояние в пунктах от левого края диаграммы до начала листа.|
||[legend](/javascript/api/excel/excel.chart#excel-excel-chart-legend-member)|Представляет условные обозначения для диаграммы.|
||[name](/javascript/api/excel/excel.chart#excel-excel-chart-name-member)|Указывает имя объекта диаграммы.|
||[series](/javascript/api/excel/excel.chart#excel-excel-chart-series-member)|Представляет один ряд данных или коллекцию рядов данных в диаграмме.|
||[setData(sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#excel-excel-chart-setdata-member(1))|Сбрасывает исходные данные для диаграммы.|
||[setPosition (startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#excel-excel-chart-setposition-member(1))|Располагает диаграмму относительно ячеек на листе.|
||[заголовок](/javascript/api/excel/excel.chart#excel-excel-chart-title-member)|Представляет заголовок указанной диаграммы, включая его текст, видимость, положение и форматирование.|
||[top](/javascript/api/excel/excel.chart#excel-excel-chart-top-member)|Указывает расстояние в точках от верхнего края объекта до верхней строки 1 (на таблице) или верхней части области диаграммы (на диаграмме).|
||[width](/javascript/api/excel/excel.chart#excel-excel-chart-width-member)|Указывает ширину объекта диаграммы в точках.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-fill-member)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-font-member)|Представляет атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для текущего объекта.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-categoryaxis-member)|Представляет ось категорий на диаграмме.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-seriesaxis-member)|Представляет ось серии 3-D диаграммы.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-valueaxis-member)|Представляет ось значений для оси.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[format](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-format-member)|Представляет форматирование объекта диаграммы, в том числе форматирование линий и шрифта.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorgridlines-member)|Возвращает объект, который представляет основные сетки для указанной оси.|
||[majorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorunit-member)|Обозначает интервал между двумя основными делениями.|
||[maximum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-maximum-member)|Представляет максимальное значение на оси значений.|
||[minimum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minimum-member)|Представляет минимальное значение на оси значений.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorgridlines-member)|Возвращает объект, который представляет второстепенные сетки для указанной оси.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorunit-member)|Представляет интервал между двумя промежуточными делениями.|
||[заголовок](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-title-member)|Обозначает название оси.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-font-member)|Указывает атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для элемента оси диаграммы.|
||[line](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-line-member)|Указывает форматирование строки диаграммы.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-format-member)|Указывает форматирование названия оси диаграммы.|
||[text](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-text-member)|Указывает заголовок оси.|
||[visible](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-visible-member)|Указывает, является ли заголовок оси visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-font-member)|Указывает атрибуты шрифта заголовок оси диаграммы, такие как имя шрифта, размер шрифта или цвет объекта заголовок оси диаграммы.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel. ChartType, sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-add-member(1))|Создает диаграмму.|
||[count](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-count-member)|Возвращает количество диаграмм на листе.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitem-member(1))|Возвращает диаграмму по ее имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemat-member(1))|Возвращает диаграмму с учетом ее положения в коллекции.|
||[items](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-fill-member)|Представляет формат заливки для текущей метки данных диаграммы.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-font-member)|Представляет атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для метки данных диаграммы.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[format](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-format-member)|Указывает формат меток данных диаграммы, включающий заполнение и форматирование шрифтов.|
||[position](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-position-member)|Значение, которое представляет положение метки данных.|
||[сепаратор](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-separator-member)|Строка, представляющая разделитель, который используется для меток данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showbubblesize-member)|Указывает, виден ли размер пузыря метки данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showcategoryname-member)|Указывает, отображается ли имя категории метки данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showlegendkey-member)|Указывает, виден ли ключ легенды метки данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showpercentage-member)|Указывает, виден ли процент метки данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showseriesname-member)|Указывает, отображается ли имя серии меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showvalue-member)|Указывает, отображается ли значение метки данных.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-clear-member(1))|Очищает цвет заполнения элемента диаграммы.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-setsolidcolor-member(1))|Устанавливает форматирование заливки элемента диаграммы на единый цвет.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-bold-member)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-color-member)|Представление цветового кода HTML текстового цвета (например, #FF0000 представляет красный).|
||[italic](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-italic-member)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-name-member)|Имя шрифта (например, "Calibri")|
||[размер](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-size-member)|Размер шрифта (например, 11)|
||[underline](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-underline-member)|Тип подчеркивания, применяемый для шрифта.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-format-member)|Представляет форматирование линий сетки диаграммы.|
||[visible](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-visible-member)|Указывает, видны ли линии сетки оси.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#excel-excel-chartgridlinesformat-line-member)|Представляет форматирование линий диаграммы.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[format](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-format-member)|Представляет форматирование легенды диаграммы, включая заливку и шрифт.|
||[наложение](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-overlay-member)|Указывает, должна ли легенда диаграммы перекрываться с основным телом диаграммы.|
||[position](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-position-member)|Указывает положение легенды на диаграмме.|
||[visible](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-visible-member)|Указывает, видна ли легенда диаграммы.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-fill-member)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-font-member)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта и цвет легенды диаграммы.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-clear-member(1))|Очищает формат строки элемента диаграммы.|
||[color](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-color-member)|HTML-код цвета, представляющий цвет линий в диаграмме.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-format-member)|Инкапсулирует свойства формата точки диаграммы.|
||[value](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-value-member)|Возвращает значение точки диаграммы.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-fill-member)|Представляет формат заполнения диаграммы, включающий сведения о формате фона.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[count](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-count-member)|Возвращает количество точек диаграммы в ряду.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getitemat-member(1))|Получение точки на основании ее положения в ряду.|
||[items](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[format](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-format-member)|Представляет форматирование ряда диаграммы, включая формат заливки и линий.|
||[name](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-name-member)|Указывает имя серии на диаграмме.|
||[точки](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-points-member)|Возвращает коллекцию всех точек в серии.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[count](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-count-member)|Возвращает количество рядов в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getitemat-member(1))|Возвращает ряд в зависимости от его позиции в коллекции.|
||[items](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-fill-member)|Представляет формат заливки ряда диаграммы, включая сведения о форматировании фона.|
||[line](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-line-member)|Представляет форматирование линий.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[format](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-format-member)|Представляет форматирование названия диаграммы, включая формат заливки и шрифта.|
||[наложение](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-overlay-member)|Указывает, будет ли заголовок диаграммы наложением диаграммы.|
||[text](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-text-member)|Указывает текст заголовка диаграммы.|
||[visible](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-visible-member)|Указывает, является ли заголовок диаграммы visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-fill-member)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-font-member)|Представляет атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для объекта.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrange-member(1))|Возвращает объект Range, сопоставленный с именем.|
||[name](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-name-member)|Имя объекта.|
||[тип](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-type-member)|Указывает тип значения, возвращаемого по формуле имени.|
||[value](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-value-member)|Представляет значение, вычисленное по формуле имени.|
||[visible](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-visible-member)|Указывает, виден ли объект.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitem-member(1))|Получает объект `NamedItem` с его именем.|
||[items](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/excel/excel.range)|[address](/javascript/api/excel/excel.range#excel-excel-range-address-member)|Указывает ссылку диапазона в стиле A1.|
||[addressLocal](/javascript/api/excel/excel.range#excel-excel-range-addresslocal-member)|Представляет ссылку диапазона для указанного диапазона на языке пользователя.|
||[cellCount](/javascript/api/excel/excel.range#excel-excel-range-cellcount-member)|Указывает количество ячеек в диапазоне.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#excel-excel-range-clear-member(1))|Очищает значения, формат, заливку, границу диапазона и т. д.|
||[columnCount](/javascript/api/excel/excel.range#excel-excel-range-columncount-member)|Указывает общее количество столбцов в диапазоне.|
||[columnIndex](/javascript/api/excel/excel.range#excel-excel-range-columnindex-member)|Указывает номер столбца первой ячейки в диапазоне.|
||[delete (shift: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-delete-member(1))|Удаляет ячейки, связанные с диапазоном.|
||[format](/javascript/api/excel/excel.range#excel-excel-range-format-member)|Возвращает объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона.|
||[formulas](/javascript/api/excel/excel.range#excel-excel-range-formulas-member)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.range#excel-excel-range-formulaslocal-member)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом.|
||[getBoundingRect (anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getboundingrect-member(1))|Возвращает наименьший объект диапазона, включающий в себя заданные диапазоны.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcell-member(1))|Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcolumn-member(1))|Возвращает столбец в диапазоне.|
||[getEntireColumn()](/javascript/api/excel/excel.range#excel-excel-range-getentirecolumn-member(1))|Получает объект, который представляет весь столбец диапазона (например, если текущий диапазон представляет ячейки "B4:E11", `getEntireColumn` он представляет столбцы "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#excel-excel-range-getentirerow-member(1))|Получает объект, который представляет весь ряд диапазона (например, если текущий диапазон представляет ячейки "B4:E11", `GetEntireRow` это диапазон, который представляет строки "4:11").|
||[getIntersection (anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersection-member(1))|Возвращает объект диапазона, представляющий собой прямоугольное пересечение заданных диапазонов.|
||[getLastCell()](/javascript/api/excel/excel.range#excel-excel-range-getlastcell-member(1))|Возвращает последнюю ячейку в диапазоне.|
||[getLastColumn()](/javascript/api/excel/excel.range#excel-excel-range-getlastcolumn-member(1))|Возвращает последний столбец в диапазоне.|
||[getLastRow()](/javascript/api/excel/excel.range#excel-excel-range-getlastrow-member(1))|Возвращает последнюю строку в диапазоне.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1))|Возвращает объект, представляющий диапазон, который смещен от указанного диапазона.|
||[getRow(row: number)](/javascript/api/excel/excel.range#excel-excel-range-getrow-member(1))|Возвращает строку из диапазона.|
||[insert(shift: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1))|Вставляет ячейку или диапазон ячеек на лист вместо этого диапазона, а также сдвигает другие ячейки, чтобы освободить место.|
||[numberFormat](/javascript/api/excel/excel.range#excel-excel-range-numberformat-member)|Представляет Excel формата номеров для данного диапазона.|
||[rowCount](/javascript/api/excel/excel.range#excel-excel-range-rowcount-member)|Возвращает общее количество строк в диапазоне.|
||[rowIndex](/javascript/api/excel/excel.range#excel-excel-range-rowindex-member)|Возвращает номер строки первой ячейки диапазона.|
||[select()](/javascript/api/excel/excel.range#excel-excel-range-select-member(1))|Выбирает указанный диапазон в пользовательском интерфейсе Excel.|
||[text](/javascript/api/excel/excel.range#excel-excel-range-text-member)|Текстовые значения указанного диапазона.|
||[valueTypes](/javascript/api/excel/excel.range#excel-excel-range-valuetypes-member)|Указывает тип данных в каждой ячейке.|
||[values](/javascript/api/excel/excel.range#excel-excel-range-values-member)|Представляет необработанные значения указанного диапазона.|
||[worksheet](/javascript/api/excel/excel.range#excel-excel-range-worksheet-member)|Лист, содержащий текущий диапазон.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-color-member)|ЦВЕТОВой код HTML, представляющий цвет пограничной строки, в форме #RRGGBB (например, "FFA500"), или в виде имени HTML-цвета (например, "оранжевый").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-sideindex-member)|Постоянное значение, указывающее определенную сторону границы.|
||[style](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-style-member)|Одна из констант стиля линии, определяющая стиль линии границы.|
||[weight](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-weight-member)|Определяет толщину границы вокруг диапазона.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[count](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-count-member)|Количество объектов границы в коллекции.|
||[getItem(index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitem-member(1))|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitemat-member(1))|Возвращает объект границы по его индексу.|
||[items](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-clear-member(1))|Сброс фона диапазона.|
||[color](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-color-member)|ЦВЕТОВой код HTML, представляющий цвет фона, в форме #RRGGBB (например, "FFA500"), или в виде имени HTML-цвета (например, "оранжевый")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-bold-member)|Представляет смелый статус шрифта.|
||[color](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-color-member)|Представление цветового кода HTML текстового цвета (например, #FF0000 представляет красный).|
||[italic](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-italic-member)|Указывает italic состояние шрифта.|
||[name](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-name-member)|Имя шрифта (например, "Калибри").|
||[размер](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-size-member)|размер шрифта|
||[underline](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-underline-member)|Тип подчеркивания, применяемый для шрифта.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[borders](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-borders-member)|Коллекция объектов границ, которые применяются ко всему диапазону.|
||[fill](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-fill-member)|Возвращает объект заливки, определенный для всего диапазона.|
||[font](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-font-member)|Возвращает объект шрифта, определенный для всего диапазона.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-horizontalalignment-member)|Представляет выравнивание по горизонтали для указанного объекта.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-verticalalignment-member)|Представляет выравнивание по вертикали для указанного объекта.|
||[wrapText](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-wraptext-member)|Указывает, Excel обертывание текста в объекте.|
|[Table](/javascript/api/excel/excel.table)|[columns](/javascript/api/excel/excel.table#excel-excel-table-columns-member)|Представляет коллекцию всех столбцов в таблице.|
||[delete()](/javascript/api/excel/excel.table#excel-excel-table-delete-member(1))|Удаляет таблицу.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#excel-excel-table-getdatabodyrange-member(1))|Получает объект диапазона, связанный с телом данных таблицы.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#excel-excel-table-getheaderrowrange-member(1))|Получает объект диапазона, связанный со строкой заголовка таблицы.|
||[getRange()](/javascript/api/excel/excel.table#excel-excel-table-getrange-member(1))|Получает объект диапазона, связанный со всей таблицей.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#excel-excel-table-gettotalrowrange-member(1))|Получает объект диапазона, связанный со строкой итогов таблицы.|
||[id](/javascript/api/excel/excel.table#excel-excel-table-id-member)|Возвращает значение, однозначно идентифицирующее таблицу в данной книге.|
||[name](/javascript/api/excel/excel.table#excel-excel-table-name-member)|Имя таблицы.|
||[строки](/javascript/api/excel/excel.table#excel-excel-table-rows-member)|Представляет коллекцию всех строк в таблице.|
||[showHeaders](/javascript/api/excel/excel.table#excel-excel-table-showheaders-member)|Указывает, видна ли строка заглавной строки.|
||[showTotals](/javascript/api/excel/excel.table#excel-excel-table-showtotals-member)|Указывает, видна ли общая строка.|
||[style](/javascript/api/excel/excel.table#excel-excel-table-style-member)|Постоянное значение, представляю которое представляет стиль таблицы.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-add-member(1))|Создает таблицу.|
||[count](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-count-member)|Возвращает количество таблиц в книге.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitem-member(1))|Получает таблицу по имени или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemat-member(1))|Получает таблицу на основании ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-delete-member(1))|Удаляет столбец из таблицы.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getdatabodyrange-member(1))|Получает объект диапазона, связанный с текстом данных столбца.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getheaderrowrange-member(1))|Получает объект диапазона, связанный со строкой заголовков столбца.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getrange-member(1))|Получает объект диапазона, связанный со всем столбцом.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-gettotalrowrange-member(1))|Получает объект диапазона, связанный со строкой итогов столбца.|
||[id](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-id-member)|Возвращает уникальный ключ, идентифицирующий столбец в таблице.|
||[индекс](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-index-member)|Возвращает номер индекса столбца в коллекции столбцов таблицы.|
||[name](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-name-member)|Указывает имя столбца таблицы.|
||[values](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-values-member)|Представляет необработанные значения указанного диапазона.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-add-member(1))|Добавляет новый столбец в таблицу.|
||[count](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-count-member)|Возвращает количество столбцов в таблице.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitem-member(1))|Возвращает объект столбца по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemat-member(1))|Возвращает столбец на основании его позиции в коллекции.|
||[items](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-delete-member(1))|Удаляет строку из таблицы.|
||[getRange()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-getrange-member(1))|Получает объект диапазона, связанный со всей строкой.|
||[индекс](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-index-member)|Возвращает номер индекса строки в коллекции строк таблицы.|
||[values](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-values-member)|Представляет необработанные значения указанного диапазона.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, alwaysInsert?: boolean)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))|Добавляет одну или несколько строк в таблицу.|
||[count](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-count-member)|Возвращает количество строк в таблице.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getitemat-member(1))|Получает строку на основании ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Workbook](/javascript/api/excel/excel.workbook)|[application](/javascript/api/excel/excel.workbook#excel-excel-workbook-application-member)|Представляет экземпляр Excel приложения, содержащий эту книгу.|
||[bindings](/javascript/api/excel/excel.workbook#excel-excel-workbook-bindings-member)|Представляет коллекцию привязок, включенных в книгу.|
||[getSelectedRange()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1))|Получает выбранный в настоящее время отдельный диапазон из книги.|
||[имена](/javascript/api/excel/excel.workbook#excel-excel-workbook-names-member)|Представляет коллекцию именных элементов с именами книг (именуемого диапазона и констант).|
||[таблицы](/javascript/api/excel/excel.workbook#excel-excel-workbook-tables-member)|Представляет коллекцию таблиц, сопоставленных с книгой.|
||[таблицы](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member)|Представляет коллекцию листов, сопоставленных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-activate-member(1))|Активация листа в пользовательском интерфейсе Excel.|
||[диаграммы](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-charts-member)|Возвращает коллекцию диаграмм, которые являются частью таблицы.|
||[delete()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-delete-member(1))|Удаляет лист из книги.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getcell-member(1))|Получает объект `Range` , содержащий одну ячейку на основе номеров строки и столбцов.|
||[getRange (адрес?: строка)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1))|Получает объект `Range` , представляющий один прямоугольный блок ячеек, указанный адресом или именем.|
||[id](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-id-member)|Возвращает значение, однозначно идентифицирующее лист в данной книге.|
||[name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member)|Отображаемое имя листа.|
||[position](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-position-member)|Положение листа (начиная с нуля) в книге.|
||[таблицы](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tables-member)|Коллекция таблиц, имеющихся на листе.|
||[visibility](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-visibility-member)|Видимость листа.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-add-member(1))|Добавляет новый лист в книгу.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getactiveworksheet-member(1))|Получает текущий активный лист в книге.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitem-member(1))|Получает объект листа по его имени или ИД.|
||[items](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

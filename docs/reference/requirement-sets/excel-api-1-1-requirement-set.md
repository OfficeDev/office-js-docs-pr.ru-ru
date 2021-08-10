---
title: Excel Набор API JavaScript 1.1
description: Сведения о наборе требований ExcelApi 1.1.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: db8754e793d86fbc1c85bae85a1ce1f925504c649b1694659896ba567dc4e478
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093825"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel Набор API JavaScript 1.1

Excel API JavaScript 1.1 — это первая версия API. Это единственный набор Excel, поддерживаемый Excel 2016.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.1. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.1 см. в Excel API в наборе [требований 1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel. CalculationType)](/javascript/api/excel/excel.application#calculate_calculationType_)|Пересчитывает данные во всех открытых в текущий момент книгах Excel.|
||[calculationMode](/javascript/api/excel/excel.application#calculationMode)|Возвращает режим вычисления, используемый в книге, как это определено константами в `Excel.CalculationMode` .|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getRange__)|Возвращает представленный привязкой диапазон.|
||[getTable()](/javascript/api/excel/excel.binding#getTable__)|Возвращает представленную привязкой таблицу.|
||[getText()](/javascript/api/excel/excel.binding#getText__)|Возвращает представленный привязкой текст.|
||[id](/javascript/api/excel/excel.binding#id)|Представляет идентификатор привязки.|
||[type](/javascript/api/excel/excel.binding#type)|Возвращает тип привязки.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getItem_id_)|Возвращает объект привязки по идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getItemAt_index_)|Возвращает объект привязки с учетом его положения в массиве элементов.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Возвращает число привязок в коллекции.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete__)|Удаляет объект диаграммы.|
||[height](/javascript/api/excel/excel.chart#height)|Указывает высоту в точках объекта диаграммы.|
||[left](/javascript/api/excel/excel.chart#left)|Расстояние в пунктах от левого края диаграммы до начала листа.|
||[name](/javascript/api/excel/excel.chart#name)|Указывает имя объекта диаграммы.|
||[axes](/javascript/api/excel/excel.chart#axes)|Представляет оси диаграммы.|
||[dataLabels](/javascript/api/excel/excel.chart#dataLabels)|Представляет метки данных на диаграмме.|
||[format](/javascript/api/excel/excel.chart#format)|Инкапсулирует свойства формата для области диаграммы.|
||[legend](/javascript/api/excel/excel.chart#legend)|Представляет условные обозначения для диаграммы.|
||[series](/javascript/api/excel/excel.chart#series)|Представляет один ряд данных или коллекцию рядов данных в диаграмме.|
||[заголовок](/javascript/api/excel/excel.chart#title)|Представляет заголовок указанной диаграммы, включая его текст, видимость, положение и форматирование.|
||[setData(sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#setData_sourceData__seriesBy_)|Сбрасывает исходные данные для диаграммы.|
||[setPosition (startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#setPosition_startCell__endCell_)|Располагает диаграмму относительно ячеек на листе.|
||[top](/javascript/api/excel/excel.chart#top)|Указывает расстояние в точках от верхнего края объекта до верхней строки 1 (на таблице) или верхней части области диаграммы (на диаграмме).|
||[width](/javascript/api/excel/excel.chart#width)|Указывает ширину объекта диаграммы в точках.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Представляет атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для текущего объекта.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryAxis)|Представляет ось категорий на диаграмме.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesAxis)|Представляет ось серии 3-D диаграммы.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueAxis)|Представляет ось значений для оси.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorUnit)|Обозначает интервал между двумя основными делениями.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Представляет максимальное значение на оси значений.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Представляет минимальное значение на оси значений.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorUnit)|Представляет интервал между двумя промежуточными делениями.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Представляет форматирование объекта диаграммы, в том числе форматирование линий и шрифта.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorGridlines)|Возвращает объект, который представляет основные сетки для указанной оси.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorGridlines)|Возвращает объект, который представляет второстепенные сетки для указанной оси.|
||[заголовок](/javascript/api/excel/excel.chartaxis#title)|Обозначает название оси.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Указывает атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для элемента оси диаграммы.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Указывает форматирование строки диаграммы.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Указывает форматирование названия оси диаграммы.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Указывает заголовок оси.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Указывает, является ли заголовок оси visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Указывает атрибуты шрифта заголовок оси диаграммы, такие как имя шрифта, размер шрифта или цвет объекта заголовок оси диаграммы.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel. ChartType, sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add_type__sourceData__seriesBy_)|Создает диаграмму.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getItem_name_)|Возвращает диаграмму по ее имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getItemAt_index_)|Возвращает диаграмму с учетом ее положения в коллекции.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Возвращает количество диаграмм на листе.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Представляет формат заливки для текущей метки данных диаграммы.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Представляет атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для метки данных диаграммы.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Значение, которое представляет положение метки данных.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Указывает формат меток данных диаграммы, включающий заполнение и форматирование шрифтов.|
||[сепаратор](/javascript/api/excel/excel.chartdatalabels#separator)|Строка, представляющая разделитель, который используется для меток данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showBubbleSize)|Указывает, виден ли размер пузыря метки данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showCategoryName)|Указывает, отображается ли имя категории метки данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showLegendKey)|Указывает, виден ли ключ легенды метки данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showPercentage)|Указывает, виден ли процент метки данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showSeriesName)|Указывает, отображается ли имя серии меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showValue)|Указывает, отображается ли значение метки данных.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear__)|Очищает цвет заполнения элемента диаграммы.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setSolidColor_color_)|Устанавливает форматирование заливки элемента диаграммы на единый цвет.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.chartfont#color)|Представление цветового кода HTML текстового цвета (например, #FF0000 представляет красный цвет).|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.chartfont#name)|Имя шрифта (например, "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|Размер шрифта (например, 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Тип подчеркивания, применяемый для шрифта.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Представляет форматирование линий сетки диаграммы.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Указывает, видны ли линии сетки оси.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Представляет форматирование линий диаграммы.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[наложение](/javascript/api/excel/excel.chartlegend#overlay)|Указывает, должна ли легенда диаграммы перекрываться с основным телом диаграммы.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Указывает положение легенды на диаграмме.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Представляет форматирование легенды диаграммы, включая заливку и шрифт.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Указывает, видна ли легенда диаграммы.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта и цвет легенды диаграммы.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear__)|Очищает формат строки элемента диаграммы.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|HTML-код цвета, представляющий цвет линий в диаграмме.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Инкапсулирует свойства формата точки диаграммы.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Возвращает значение точки диаграммы.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Представляет формат заполнения диаграммы, включающий сведения о формате фона.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getItemAt_index_)|Получение точки на основании ее положения в ряду.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Возвращает количество точек диаграммы в ряду.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Указывает имя серии на диаграмме.|
||[format](/javascript/api/excel/excel.chartseries#format)|Представляет форматирование ряда диаграммы, включая формат заливки и линий.|
||[точки](/javascript/api/excel/excel.chartseries#points)|Возвращает коллекцию всех точек в серии.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getItemAt_index_)|Возвращает ряд в зависимости от его позиции в коллекции.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Возвращает количество рядов в коллекции.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Представляет формат заливки ряда диаграммы, включая сведения о форматировании фона.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Представляет форматирование линий.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[наложение](/javascript/api/excel/excel.charttitle#overlay)|Указывает, будет ли заголовок диаграммы наложением диаграммы.|
||[format](/javascript/api/excel/excel.charttitle#format)|Представляет форматирование названия диаграммы, включая формат заливки и шрифта.|
||[text](/javascript/api/excel/excel.charttitle#text)|Указывает текст заголовка диаграммы.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Указывает, является ли заголовок диаграммы visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Представляет атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для объекта.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getRange__)|Возвращает объект Range, сопоставленный с именем.|
||[name](/javascript/api/excel/excel.nameditem#name)|Имя объекта.|
||[type](/javascript/api/excel/excel.nameditem#type)|Указывает тип значения, возвращаемого по формуле имени.|
||[value](/javascript/api/excel/excel.nameditem#value)|Представляет значение, вычисленное по формуле имени.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Указывает, виден ли объект.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getItem_name_)|Получает объект `NamedItem` с его именем.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear_applyTo_)|Очищает значения, формат, заливку, границу диапазона и т. д.|
||[delete (shift: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#delete_shift_)|Удаляет ячейки, связанные с диапазоном.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulasLocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом.|
||[getBoundingRect (anotherRange: Range \| string)](/javascript/api/excel/excel.range#getBoundingRect_anotherRange_)|Возвращает наименьший объект диапазона, включающий в себя заданные диапазоны.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getCell_row__column_)|Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getColumn_column_)|Возвращает столбец в диапазоне.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getEntireColumn__)|Получает объект, который представляет весь столбец диапазона (например, если текущий диапазон представляет ячейки "B4:E11", он представляет столбцы `getEntireColumn` "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#getEntireRow__)|Получает объект, который представляет весь ряд диапазона (например, если текущий диапазон представляет ячейки "B4:E11", это диапазон, который представляет строки `GetEntireRow` "4:11").|
||[getIntersection (anotherRange: Range \| string)](/javascript/api/excel/excel.range#getIntersection_anotherRange_)|Возвращает объект диапазона, представляющий собой прямоугольное пересечение заданных диапазонов.|
||[getLastCell()](/javascript/api/excel/excel.range#getLastCell__)|Возвращает последнюю ячейку в диапазоне.|
||[getLastColumn()](/javascript/api/excel/excel.range#getLastColumn__)|Возвращает последний столбец в диапазоне.|
||[getLastRow()](/javascript/api/excel/excel.range#getLastRow__)|Возвращает последнюю строку в диапазоне.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getOffsetRange_rowOffset__columnOffset_)|Возвращает объект, представляющий диапазон, который смещен от указанного диапазона.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getRow_row_)|Возвращает строку из диапазона.|
||[insert(shift: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#insert_shift_)|Вставляет ячейку или диапазон ячеек на лист вместо этого диапазона, а также сдвигает другие ячейки, чтобы освободить место.|
||[numberFormat](/javascript/api/excel/excel.range#numberFormat)|Представляет Excel формата номеров для данного диапазона.|
||[address](/javascript/api/excel/excel.range#address)|Указывает ссылку диапазона в стиле A1.|
||[addressLocal](/javascript/api/excel/excel.range#addressLocal)|Представляет ссылку диапазона для указанного диапазона на языке пользователя.|
||[cellCount](/javascript/api/excel/excel.range#cellCount)|Указывает количество ячеек в диапазоне.|
||[columnCount](/javascript/api/excel/excel.range#columnCount)|Указывает общее количество столбцов в диапазоне.|
||[columnIndex](/javascript/api/excel/excel.range#columnIndex)|Указывает номер столбца первой ячейки в диапазоне.|
||[format](/javascript/api/excel/excel.range#format)|Возвращает объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона.|
||[rowCount](/javascript/api/excel/excel.range#rowCount)|Возвращает общее количество строк в диапазоне.|
||[rowIndex](/javascript/api/excel/excel.range#rowIndex)|Возвращает номер строки первой ячейки диапазона.|
||[text](/javascript/api/excel/excel.range#text)|Текстовые значения указанного диапазона.|
||[valueTypes](/javascript/api/excel/excel.range#valueTypes)|Указывает тип данных в каждой ячейке.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|Лист, содержащий текущий диапазон.|
||[select()](/javascript/api/excel/excel.range#select__)|Выбирает указанный диапазон в пользовательском интерфейсе Excel.|
||[values](/javascript/api/excel/excel.range#values)|Представляет необработанные значения указанного диапазона.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|ЦВЕТОВой код HTML, представляющий цвет пограничной строки, в форме #RRGGBB (например, "FFA500"), или в виде имени HTML-цвета (например, "оранжевый").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideIndex)|Постоянное значение, указывающее определенную сторону границы.|
||[style](/javascript/api/excel/excel.rangeborder#style)|Одна из констант стиля линии, определяющая стиль линии границы.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Определяет толщину границы вокруг диапазона.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem(index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getItem_index_)|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getItemAt_index_)|Возвращает объект границы по его индексу.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Количество объектов границы в коллекции.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear__)|Сброс фона диапазона.|
||[color](/javascript/api/excel/excel.rangefill#color)|ЦВЕТОВой код HTML, представляющий цвет фона, в форме #RRGGBB (например, "FFA500"), или в виде имени HTML-цвета (например, "оранжевый")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Представляет смелый статус шрифта.|
||[color](/javascript/api/excel/excel.rangefont#color)|Представление цветового кода HTML текстового цвета (например, #FF0000 представляет красный цвет).|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Указывает italic состояние шрифта.|
||[name](/javascript/api/excel/excel.rangefont#name)|Имя шрифта (например, "Калибри").|
||[size](/javascript/api/excel/excel.rangefont#size)|размер шрифта|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Тип подчеркивания, применяемый для шрифта.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalAlignment)|Представляет выравнивание по горизонтали для указанного объекта.|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|Коллекция объектов границ, которые применяются ко всему диапазону.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Возвращает объект заливки, определенный для всего диапазона.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Возвращает объект шрифта, определенный для всего диапазона.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalAlignment)|Представляет выравнивание по вертикали для указанного объекта.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wrapText)|Указывает, Excel обертывание текста в объекте.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete__)|Удаляет таблицу.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getDataBodyRange__)|Получает объект диапазона, связанный с телом данных таблицы.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getHeaderRowRange__)|Получает объект диапазона, связанный со строкой заголовка таблицы.|
||[getRange()](/javascript/api/excel/excel.table#getRange__)|Получает объект диапазона, связанный со всей таблицей.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#getTotalRowRange__)|Получает объект диапазона, связанный со строкой итогов таблицы.|
||[name](/javascript/api/excel/excel.table#name)|Имя таблицы.|
||[columns](/javascript/api/excel/excel.table#columns)|Представляет коллекцию всех столбцов в таблице.|
||[id](/javascript/api/excel/excel.table#id)|Возвращает значение, однозначно идентифицирующее таблицу в данной книге.|
||[строки](/javascript/api/excel/excel.table#rows)|Представляет коллекцию всех строк в таблице.|
||[showHeaders](/javascript/api/excel/excel.table#showHeaders)|Указывает, видна ли строка заглавной строки.|
||[showTotals](/javascript/api/excel/excel.table#showTotals)|Указывает, видна ли общая строка.|
||[style](/javascript/api/excel/excel.table#style)|Постоянное значение, представляю которое представляет стиль таблицы.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add_address__hasHeaders_)|Создает таблицу.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getItem_key_)|Получает таблицу по имени или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getItemAt_index_)|Получает таблицу на основании ее позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Возвращает количество таблиц в книге.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete__)|Удаляет столбец из таблицы.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getDataBodyRange__)|Получает объект диапазона, связанный с текстом данных столбца.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getHeaderRowRange__)|Получает объект диапазона, связанный со строкой заголовков столбца.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getRange__)|Получает объект диапазона, связанный со всем столбцом.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#getTotalRowRange__)|Получает объект диапазона, связанный со строкой итогов столбца.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Указывает имя столбца таблицы.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Возвращает уникальный ключ, идентифицирующий столбец в таблице.|
||[index](/javascript/api/excel/excel.tablecolumn#index)|Возвращает номер индекса столбца в коллекции столбцов таблицы.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Представляет необработанные значения указанного диапазона.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add_index__values__name_)|Добавляет новый столбец в таблицу.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getItem_key_)|Возвращает объект столбца по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getItemAt_index_)|Возвращает столбец на основании его позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Возвращает количество столбцов в таблице.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete__)|Удаляет строку из таблицы.|
||[getRange()](/javascript/api/excel/excel.tablerow#getRange__)|Получает объект диапазона, связанный со всей строкой.|
||[index](/javascript/api/excel/excel.tablerow#index)|Возвращает номер индекса строки в коллекции строк таблицы.|
||[values](/javascript/api/excel/excel.tablerow#values)|Представляет необработанные значения указанного диапазона.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number)](/javascript/api/excel/excel.tablerowcollection#add_index__values_)|Добавляет одну или несколько строк в таблицу.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getItemAt_index_)|Получает строку на основании ее позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Возвращает количество строк в таблице.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange()](/javascript/api/excel/excel.workbook#getSelectedRange__)|Получает выбранный в настоящее время отдельный диапазон из книги.|
||[application](/javascript/api/excel/excel.workbook#application)|Представляет экземпляр Excel, содержащий эту книгу.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Представляет коллекцию привязок, включенных в книгу.|
||[имена](/javascript/api/excel/excel.workbook#names)|Представляет коллекцию именных элементов с именами книг (именуемого диапазона и констант).|
||[таблицы](/javascript/api/excel/excel.workbook#tables)|Представляет коллекцию таблиц, сопоставленных с книгой.|
||[таблицы](/javascript/api/excel/excel.workbook#worksheets)|Представляет коллекцию листов, сопоставленных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate__)|Активация листа в пользовательском интерфейсе Excel.|
||[delete()](/javascript/api/excel/excel.worksheet#delete__)|Удаляет лист из книги.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getCell_row__column_)|Получает `Range` объект, содержащий одну ячейку на основе номеров строки и столбцов.|
||[getRange (адрес?: строка)](/javascript/api/excel/excel.worksheet#getRange_address_)|Получает `Range` объект, представляющий один прямоугольный блок ячеек, указанный адресом или именем.|
||[name](/javascript/api/excel/excel.worksheet#name)|Отображаемое имя листа.|
||[position](/javascript/api/excel/excel.worksheet#position)|Положение листа (начиная с нуля) в книге.|
||[диаграммы](/javascript/api/excel/excel.worksheet#charts)|Возвращает коллекцию диаграмм, которые являются частью таблицы.|
||[id](/javascript/api/excel/excel.worksheet#id)|Возвращает значение, однозначно идентифицирующее лист в данной книге.|
||[таблицы](/javascript/api/excel/excel.worksheet#tables)|Коллекция таблиц, имеющихся на листе.|
||[видимость](/javascript/api/excel/excel.worksheet#visibility)|Видимость листа.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#add_name_)|Добавляет новый лист в книгу.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getActiveWorksheet__)|Получает текущий активный лист в книге.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getItem_key_)|Получает объект листа по его имени или ИД.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

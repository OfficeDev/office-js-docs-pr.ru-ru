---
title: Набор обязательных элементов API JavaScript для Excel 1,1
description: Сведения о наборе требований ExcelApi 1,1
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 90d7ee7cef2e8c48e458b2e14893ba9c13c68a30
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940789"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Набор обязательных элементов API JavaScript для Excel 1,1

API JavaScript для Excel 1,1 — это первая версия API. Это единственный набор обязательных элементов Excel, поддерживаемый Excel 2016.

## <a name="api-list"></a>Список API

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[Calculate (Калкулатионтипе: Excel. Калкулатионтипе)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Пересчитывает данные во всех открытых в текущий момент книгах Excel.|
||[Калкулатионмоде](/javascript/api/excel/excel.application#calculationmode)|Возвращает режим вычислений, используемый в книге в соответствии с константами в Excel. Калкулатионмоде. Возможные значения: `Automatic`, где Excel управляет пересчетом; `AutomaticExceptTables`, где Excel контролирует пересчет, но игнорирует изменения в таблицах; `Manual`, где выполняется расчет, когда пользователь запрашивает его.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Возвращает представленный привязкой диапазон. Если тип привязки неправильный, выдается ошибка.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Возвращает представленную привязкой таблицу. Если тип привязки неправильный, выдается ошибка.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Возвращает представленный привязкой текст. Если тип привязки неправильный, выдается ошибка.|
||[id](/javascript/api/excel/excel.binding#id)|Представляет идентификатор привязки. Только для чтения.|
||[type](/javascript/api/excel/excel.binding#type)|Возвращает тип привязки. Дополнительные сведения см. в статье Excel. BindingType. Только для чтения.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Возвращает объект привязки по идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Возвращает объект привязки с учетом его положения в массиве элементов.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Возвращает число привязок в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Удаляет объект диаграммы.|
||[height](/javascript/api/excel/excel.chart#height)|Обозначает высоту объекта диаграммы (в пунктах).|
||[left](/javascript/api/excel/excel.chart#left)|Расстояние в пунктах от левого края диаграммы до начала листа.|
||[name](/javascript/api/excel/excel.chart#name)|Обозначает имя объекта диаграммы.|
||[Axes](/javascript/api/excel/excel.chart#axes)|Представляет оси диаграммы. Только для чтения.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Представляет метки данных на диаграмме. Только для чтения.|
||[format](/javascript/api/excel/excel.chart#format)|Инкапсулирует свойства формата для области диаграммы. Только для чтения.|
||[списком](/javascript/api/excel/excel.chart#legend)|Представляет условные обозначения для диаграммы. Только для чтения.|
||[series](/javascript/api/excel/excel.chart#series)|Представляет один ряд данных или коллекцию рядов данных в диаграмме. Только для чтения.|
||[заголовок](/javascript/api/excel/excel.chart#title)|Представляет заголовок указанной диаграммы, включая его текст, видимость, положение и форматирование. Только для чтения.|
||[setData (sourceData: Range, seriesBy?: Excel. Чартсериесби)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Сбрасывает исходные данные для диаграммы.|
||[setPosition (startCell: строка \| диапазона, endCell?: строка \| диапазона)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Располагает диаграмму относительно ячеек на листе.|
||[top](/javascript/api/excel/excel.chart#top)|Представляет расстояние в пунктах от верхнего края объекта до верхнего края первой строки (на листе) или до верхнего края области диаграммы (на диаграмме).|
||[width](/javascript/api/excel/excel.chart#width)|Представляет ширину объекта диаграммы (в пунктах).|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона. Только для чтения.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для текущего объекта. Только для чтения.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[Категоряксис](/javascript/api/excel/excel.chartaxes#categoryaxis)|Представляет ось категорий на диаграмме. Только для чтения.|
||[Сериесаксис](/javascript/api/excel/excel.chartaxes#seriesaxis)|Представляет ось ряда данных для объемной диаграммы. Только для чтения.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Представляет ось значений для оси. Только для чтения.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Обозначает интервал между двумя основными делениями. Можно указать в виде числового значения или пустой строки.  Возвращаемое значение всегда является числом.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Представляет максимальное значение на оси значений.  Можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси).  Возвращаемое значение всегда является числом.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Представляет минимальное значение на оси значений. Ему можно присвоить числовое значение или пустую строку (для автоматически заданных значений оси). Всегда возвращает числовое значение.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Представляет интервал между двумя промежуточными делениями. Его можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси). Возвращаемое значение всегда является числом.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Представляет форматирование объекта диаграммы, в том числе форматирование линий и шрифта. Только для чтения.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Возвращает объект линии сетки, который представляет основные линии сетки для указанной оси. Только для чтения.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Возвращает объект Gridlines, который представляет вспомогательные линии сетки для указанной оси. Только для чтения.|
||[заголовок](/javascript/api/excel/excel.chartaxis#title)|Обозначает название оси. Только для чтения.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для элемента оси диаграммы. Только для чтения.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Представляет форматирование линий диаграммы. Только для чтения.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Представляет форматирование для названия оси диаграммы. Только для чтения.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Обозначает название оси.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Логическое значение, которое определяет видимость названия оси.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д. объект заголовка оси диаграммы. Только для чтения.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[Добавить (тип: Excel. ChartType, sourceData: Range, seriesBy?: Excel. Чартсериесби)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Создает диаграмму.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Возвращает диаграмму по ее имени. Если одно и то же имя принадлежит нескольким диаграммам, будет возвращена первая из них.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Возвращает диаграмму с учетом ее положения в коллекции.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Возвращает количество диаграмм на листе. Только для чтения.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Представляет формат заливки для текущей метки данных диаграммы. Только для чтения.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Представляет атрибуты шрифта (имя, размер шрифта, цвет и т. д.) для подписи данных диаграммы. Только для чтения.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Значение DataLabelPosition, которое представляет положение метки данных. Дополнительные сведения см. в статье Excel. Чартдаталабелпоситион.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Представляет формат меток данных диаграммы, включая форматирование заливки и шрифтов. Только для чтения.|
||[символ](/javascript/api/excel/excel.chartdatalabels#separator)|Строка, представляющая разделитель, который используется для меток данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Логическое значение, которое указывает, отображается ли значение метки данных.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Очищает цвет заливки элемента диаграммы.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Устанавливает форматирование заливки элемента диаграммы на единый цвет.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.chartfont#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.chartfont#name)|Имя шрифта (например, Calibri)|
||[size](/javascript/api/excel/excel.chartfont#size)|Размер шрифта (например, 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Чартундерлинестиле.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Представляет форматирование линий сетки диаграммы. Только для чтения.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Логическое значение, определяющее, отображаются ли линии сетки оси.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Представляет форматирование линий диаграммы. Только для чтения.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[накладывающиеся](/javascript/api/excel/excel.chartlegend#overlay)|Логическое значение, определяющее, должна ли легенда диаграммы перекрываться с основной частью диаграммы.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Представляет расположение легенды на диаграмме. Дополнительные сведения см. в статье Excel. Чартлежендпоситион.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Представляет форматирование легенды диаграммы, включая заливку и шрифт. Только для чтения.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Логическое значение, представляющее видимость объекта ChartLegend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона. Только для чтения.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта, цвет и т. д., в условных обозначениях диаграммы. Только для чтения.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Очищает формат линий элемента диаграммы.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|HTML-код цвета, представляющий цвет линий в диаграмме.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Инкапсулирует свойства формата точки диаграммы. Только для чтения.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Возвращает значение точки диаграммы. Только для чтения.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Представляет формат заливки диаграммы, включающий сведения о форматировании фона. Только для чтения.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Получение точки на основании ее положения в ряду.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Возвращает количество точек диаграммы в ряду. Только для чтения.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Представляет имя ряда в диаграмме.|
||[format](/javascript/api/excel/excel.chartseries#format)|Представляет форматирование ряда диаграммы, включая формат заливки и линий. Только для чтения.|
||[этапах](/javascript/api/excel/excel.chartseries#points)|Представляет коллекцию всех точек в ряду. Только для чтения.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Возвращает ряд в зависимости от его позиции в коллекции.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Возвращает число рядов в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Представляет формат заливки ряда диаграммы, включая сведения о форматировании фона. Только для чтения.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Представляет форматирование линий. Только для чтения.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[накладывающиеся](/javascript/api/excel/excel.charttitle#overlay)|Логическое значение, определяющее, отображается ли заголовок диаграммы поверх нее.|
||[format](/javascript/api/excel/excel.charttitle#format)|Представляет форматирование названия диаграммы, включая формат заливки и шрифта. Только для чтения.|
||[text](/javascript/api/excel/excel.charttitle#text)|Представляет текст заголовка диаграммы.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Логическое значение, представляющее видимость объекта заголовка диаграммы.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона. Только для чтения.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Представляет атрибуты шрифта (имя шрифта, размер шрифта, цвет и т. д.) для объекта. Только для чтения.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Возвращает объект диапазона, связанный с именем. Выдает ошибку, если именованный элемент не является диапазоном.|
||[name](/javascript/api/excel/excel.nameditem#name)|Имя объекта. Только для чтения.|
||[type](/javascript/api/excel/excel.nameditem#type)|Указывает тип значения, возвращаемый формулой имени. Дополнительные сведения см. в статье Excel. Намедитемтипе. Только для чтения.|
||[value](/javascript/api/excel/excel.nameditem#value)|Представляет значение, вычисленное по формуле имени. Если задан именованный диапазон, возвращается адрес диапазона. Только для чтения.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Определяет, является ли объект видимым.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Возвращает объект NamedItem, используя его имя.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Очищает значения, формат, заливку, границу диапазона и т. д.|
||[Delete (Shift: Excel. Делетешифтдиректион)](/javascript/api/excel/excel.range#delete-shift-)|Удаляет ячейки, связанные с диапазоном.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[getBoundingRect (anotherRange: строка \| Range)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Возвращает наименьший объект диапазона, включающий в себя заданные диапазоны. Например, GetBoundingRect для "B2:C5" и "D10:E15" возвращает значение "B2:E15".|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца. Ячейка может находиться вне границ родительского диапазона, пока она остается в сетке листа. Возвращаемая ячейка располагается относительно верхней левой ячейки диапазона.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Возвращает столбец в диапазоне.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Получает объект, представляющий весь столбец диапазона (например, если текущий диапазон представляет ячейки "B4: E11", а `getEntireColumn` — диапазон, представляющий столбцы "б:е").|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Получает объект, представляющий всю строку диапазона (например, если текущий диапазон представляет ячейки "B4: E11", а `GetEntireRow` — диапазон, представляющий строки "4:11").|
||[пересечение (anotherRange: строка \| Range)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Возвращает объект диапазона, представляющий собой прямоугольное пересечение заданных диапазонов.|
||[Жетластцелл ()](/javascript/api/excel/excel.range#getlastcell--)|Возвращает последнюю ячейку в диапазоне. Например, последняя ячейка диапазона B2:D5 — D5.|
||[Жетластколумн ()](/javascript/api/excel/excel.range#getlastcolumn--)|Возвращает последний столбец в диапазоне. Например, последний столбец диапазона B2:D5 — D2:D5.|
||[Жетластров ()](/javascript/api/excel/excel.range#getlastrow--)|Возвращает последнюю строку в диапазоне. Например, последняя строка в диапазоне "B2:D5" — "B5:D5".|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Возвращает объект, представляющий диапазон, который смещен от указанного диапазона. Измерение возвращаемого диапазона будет соответствовать этому диапазону. Если результирующий диапазон выходит за пределы таблицы листа, возникнет ошибка.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Возвращает строку из диапазона.|
||[INSERT (Shift: Excel. Инсертшифтдиректион)](/javascript/api/excel/excel.range#insert-shift-)|Вставляет ячейку или диапазон ячеек на лист вместо этого диапазона, а также сдвигает другие ячейки, чтобы освободить место. Возвращает новый объект Range в пустом месте.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Представляет код числового формата Excel для заданного диапазона.|
||[address](/javascript/api/excel/excel.range#address)|Представляет ссылку на диапазон в стиле A1. Значение Address будет содержать ссылку на лист (например, "Лист1! A1: B4). Только для чтения.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Представляет ссылку на указанный диапазон на языке пользователя. Только для чтения.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Количество ячеек в диапазоне. Этот API возвращает значение -1, если количество ячеек превышает 2^31-1 (2,147,483,647). Только для чтения.|
||[Число](/javascript/api/excel/excel.range#columncount)|Представляет общее количество столбцов в диапазоне. Только для чтения.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Представляет номер столбца первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
||[format](/javascript/api/excel/excel.range#format)|Возвращает объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона. Только для чтения.|
||[Стро](/javascript/api/excel/excel.range#rowcount)|Возвращает общее количество строк в диапазоне. Только для чтения.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Возвращает номер строки первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
||[text](/javascript/api/excel/excel.range#text)|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Представляет тип данных каждой ячейки. Только для чтения.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|Лист, содержащий текущий диапазон. Только для чтения.|
||[select()](/javascript/api/excel/excel.range#select--)|Выбирает указанный диапазон в пользовательском интерфейсе Excel.|
||[values](/javascript/api/excel/excel.range#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[Сидеиндекс](/javascript/api/excel/excel.rangeborder#sideindex)|Постоянное значение, указывающее определенную сторону границы. Дополнительные сведения см. в статье Excel. Бордериндекс. Только для чтения.|
||[style](/javascript/api/excel/excel.rangeborder#style)|Одна из констант стиля линии, определяющая стиль линии границы. Дополнительные сведения см. в статье Excel. Бордерлинестиле.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Определяет толщину границы вокруг диапазона. Дополнительные сведения см. в статье Excel. Бордервеигхт.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[GetItem (index: Excel. Бордериндекс)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Возвращает объект границы по его индексу.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Количество объектов границы в коллекции. Только для чтения.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Сброс фона диапазона.|
||[color](/javascript/api/excel/excel.rangefill#color)|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова в HTML (например, orange).|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.rangefont#color)|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.rangefont#name)|Имя шрифта (например, Calibri)|
||[size](/javascript/api/excel/excel.rangefont#size)|размер шрифта|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Тип подчеркивания, применяемый для шрифта. Дополнительные сведения см. в статье Excel. Ранжеундерлинестиле.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Представляет выравнивание по горизонтали для указанного объекта. Дополнительные сведения см. в статье Excel. HorizontalAlignment.|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|Коллекция объектов границ, которые применяются ко всему диапазону. Только для чтения.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Возвращает объект заливки, определенный для всего диапазона. Только для чтения.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Возвращает объект шрифта, определенный для всего диапазона. Только для чтения.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Представляет выравнивание по вертикали для указанного объекта. Дополнительные сведения см. в статье Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Указывает, использует ли Excel обтекание текстом для объекта. Значение null указывает, что для диапазона в целом не применяется согласованный параметр обтекания.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Удаляет таблицу.|
||[Жетдатабодиранже ()](/javascript/api/excel/excel.table#getdatabodyrange--)|Получает объект диапазона, связанный с телом данных таблицы.|
||[Жесеадерровранже ()](/javascript/api/excel/excel.table#getheaderrowrange--)|Получает объект диапазона, связанный со строкой заголовков таблицы.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Получает объект диапазона, связанный со всей таблицей.|
||[Жеттоталровранже ()](/javascript/api/excel/excel.table#gettotalrowrange--)|Получает объект диапазона, связанный со строкой итогов таблицы.|
||[name](/javascript/api/excel/excel.table#name)|Имя таблицы.|
||[столбцы](/javascript/api/excel/excel.table#columns)|Представляет коллекцию всех столбцов в таблице. Только для чтения.|
||[id](/javascript/api/excel/excel.table#id)|Возвращает значение, однозначно идентифицирующее таблицу в данной книге. Значение идентификатора остается прежним, даже если переименовать таблицу. Только для чтения.|
||[строки](/javascript/api/excel/excel.table#rows)|Представляет коллекцию всех строк в таблице. Только для чтения.|
||[Шовхеадерс](/javascript/api/excel/excel.table#showheaders)|Указывает, отображается ли строка заголовков. Можно задать это значение, чтобы отобразить или скрыть строку заголовков.|
||[Шовтоталс](/javascript/api/excel/excel.table#showtotals)|Указывает, отображается ли строка итогов. Можно задать это значение, чтобы отобразить или скрыть строку итогов.|
||[style](/javascript/api/excel/excel.table#style)|Постоянное значение, представляющее стиль таблицы. Возможные значения: от TableStyleLight1 до TableStyleLight21, от TableStyleMedium1 до TableStyleMedium28, от TableStyleStyleDark1 до TableStyleStyleDark11. Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[Add (Address: строка \| диапазона, hasHeaders: Boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Создание таблицы. Объект или исходный адрес диапазона определяет лист, на который будет добавлена таблица. Если добавить таблицу не удается (например, если адрес недействителен или одна таблица будет перекрываться другой), выводится сообщение об ошибке.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Получает таблицу по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Получает таблицу на основании ее позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Возвращает количество таблиц в книге. Только для чтения.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Удаляет столбец из таблицы.|
||[Жетдатабодиранже ()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Получает объект диапазона, связанный с текстом данных столбца.|
||[Жесеадерровранже ()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Получает объект диапазона, связанный со строкой заголовков столбца.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Получает объект диапазона, связанный со всем столбцом.|
||[Жеттоталровранже ()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Получает объект диапазона, связанный со строкой итогов столбца.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Представляет имя столбца таблицы.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Возвращает уникальный ключ, идентифицирующий столбец в таблице. Только для чтения.|
||[индекс](/javascript/api/excel/excel.tablecolumn#index)|Возвращает номер индекса столбца в коллекции столбцов таблицы. Используется нулевой индекс. Только для чтения.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[Add (index?: число, Values?: массив<массив<логический \| номер \| строки>> \| логический \| номер \| строки, Name?: строка)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Добавляет новый столбец в таблицу.|
||[GetItem (ключ: число \| строка)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Возвращает объект column по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Возвращает столбец на основании его позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Возвращает количество столбцов в таблице. Только для чтения.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Удаляет строку из таблицы.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Получает объект диапазона, связанный со всей строкой.|
||[индекс](/javascript/api/excel/excel.tablerow#index)|Возвращает номер индекса строки в коллекции строк таблицы. Используется нулевой индекс. Только для чтения.|
||[values](/javascript/api/excel/excel.tablerow#values)|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[Add (index?: число, Values?: массив<массив<логический \| номер \| строки>> \| логический \| номер \| строки)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Добавляет одну или несколько строк в таблицу. Возвращается объект, находящийся над новыми строками.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Получает строку на основании ее позиции в коллекции.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Возвращает количество строк в таблице. Только для чтения.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Функцией getselectedrange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|Получает текущий выделенный диапазон из книги. Если выбрано несколько диапазонов, этот метод выдаст ошибку.|
||[application](/javascript/api/excel/excel.workbook#application)|Представляет экземпляр приложения Excel, который содержит эту книгу. Только для чтения.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Представляет коллекцию привязок, включенных в книгу. Только для чтения.|
||[names](/javascript/api/excel/excel.workbook#names)|Представляет коллекцию именованных элементов в книге (именованные диапазоны и константы). Только для чтения.|
||[Table](/javascript/api/excel/excel.workbook#tables)|Представляет коллекцию таблиц, сопоставленных с книгой. Только для чтения.|
||[листов](/javascript/api/excel/excel.workbook#worksheets)|Представляет коллекцию листов, сопоставленных с книгой. Только для чтения.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Активация листа в пользовательском интерфейсе Excel.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Удаляет лист из книги. Обратите внимание, что если для отображения листа задано значение "Верихидден", операция удаления завершится с помощью GeneralException.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца. Ячейка может находиться вне границ родительского диапазона, пока она остается в сетке листа.|
||[GetString (Address?: строка)](/javascript/api/excel/excel.worksheet#getrange-address-)|Получает объект Range, представляющий отдельный прямоугольный блок ячеек, заданный по адресу или имени.|
||[name](/javascript/api/excel/excel.worksheet#name)|Отображаемое имя листа.|
||[position](/javascript/api/excel/excel.worksheet#position)|Положение листа (начиная с нуля) в книге.|
||[темп](/javascript/api/excel/excel.worksheet#charts)|Возвращает коллекцию диаграмм, имеющихся на листе. Только для чтения.|
||[id](/javascript/api/excel/excel.worksheet#id)|Возвращает значение, однозначно идентифицирующее лист в данной книге. Значение идентификатора остается прежним, даже если переименовать или переместить лист. Только для чтения.|
||[Table](/javascript/api/excel/excel.worksheet#tables)|Коллекция таблиц, имеющихся на листе. Только для чтения.|
||[доступности](/javascript/api/excel/excel.worksheet#visibility)|Видимость листа.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[Добавить (имя?: строка)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Добавляет новый лист в книгу. Лист будет добавлен в конец набора имеющихся листов. Если вы хотите активировать только что добавленный лист, вызовите команду .activate().|
||[Жетактивеворкшит ()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Получает текущий активный лист в книге.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Получает объект листа по его имени или ИД.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)

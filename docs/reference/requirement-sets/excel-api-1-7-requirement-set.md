---
title: Excel Набор API JavaScript 1.7
description: Сведения о наборе требований ExcelApi 1.7.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 81ae4b7ec9180ebb14bdf3b0e19d6dc2a9e997cf
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153969"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Новые возможности API JavaScript для Excel 1.7

Функции набора обязательных элементов API JavaScript для Excel 1.7 включают API для диаграмм, событий, рабочих листов, диапазонов, свойств документа, именованных элементов, параметров защиты и стилей.

## <a name="customize-charts"></a>Настройка диаграмм

С помощью новых API диаграмм можно создавать дополнительные типы диаграмм, добавлять ряды данных в диаграмму, задавать заголовок диаграммы, добавлять заголовок оси, добавлять отображаемые единицы, добавлять линию тренда со скользящей средней, менять линию тренда на линейную и многое другое. Ниже приведены некоторые примеры.

- Ось диаграммы — получайте, задавайте, форматируйте и удаляйте единицу измерения, метку и заголовок оси на диаграмме.
- Ряды диаграммы — добавляйте, задавайте и удаляйте ряды на диаграмме.  Изменяйте маркеры рядов, порядок и размер построения.
- Линии трендов диаграммы — добавляйте, получайте и форматируйте линии тренда на диаграмме.
- Условные обозначения диаграммы — форматируйте шрифт условных обозначений на диаграмме.
- Точка диаграммы — задавайте цвет точки диаграммы.
- Подстрока заголовка диаграммы — получайте и задавайте подстроку заголовка для диаграммы.
- Тип диаграммы — параметр для создания дополнительных типов диаграмм.

## <a name="events"></a>События

API событий Excel предоставляют разнообразные обработчики событий, которые позволяют вашей надстройке автоматически запускать назначенную функцию при возникновении определенного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. Список доступных событий см. в статье [Работа с событиями с помощью API JavaScript для Excel](../../excel/excel-add-ins-events.md).

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Настройка внешнего вида листов и диапазонов

С помощью новых интерфейсов API можно настроить внешний вид листов несколькими способами:

- Закрепляйте области, чтобы отображать отдельные строки или столбцы при прокрутке листа. Например, если первая строка на вашем листе содержит заголовки, вы можете закрепить эту строку, чтобы заголовки столбцов оставались видимыми при прокрутке листа.
- Изменяйте цвета вкладки листа.
- Добавляйте заголовки листов.

Внешний вид диапазонов можно настроить несколькими способами:

- Задавайте стиль ячейки для диапазона, чтобы обеспечить для всех ячеек в диапазоне единообразное форматирование. Стиль ячейки — определенный набор параметров форматирования, таких как шрифты и размеры шрифтов, форматы чисел, границы ячейки и заливка ячеек. Используйте любой из встроенных стилей ячеек Excel или создайте свой собственный стиль ячейки.
- Настройте ориентацию текста для диапазона.
- Добавляйте или изменяйте гиперссылку в диапазоне, ведущую в другое место в рабочей книге или на внешнее расположение.

## <a name="manage-document-properties"></a>Управление свойствами документа

С помощью API свойств документа можно получить доступ к встроенным свойствам документа, а также создавать и управлять настраиваемыми свойствами документа для хранения состояния книги и управления рабочим процессом и бизнес-логикой.

## <a name="copy-worksheets"></a>Копирование листов

С помощью API копирования листа вы можете копировать данные и формат с одного листа на новый рабочий лист в пределах одной книги и уменьшить объем необходимой передачи данных.

## <a name="handle-ranges-with-ease"></a>Удобная обработка диапазонов

С помощью различных API-интерфейсов диапазона можно выполнять такие действия, как получение окружающей области, получение диапазона с измененными размерами и многое другое.  Эти API позволят намного эффективнее выполнять задачи обработки и адресации диапазонов.

Дополнительно:

- Параметры защиты книги и листа — используйте эти API для защиты данных на листе и в структуре книги.
- Обновление именованного элемента — используйте этот API для обновления именованного элемента.
- Получение активной ячейки — используйте этот API для получения активной ячейки книги.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, установленный 1.7. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.7 или ранее, см. в Excel API в наборе требований [1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#chartType)|Указывает тип диаграммы.|
||[id](/javascript/api/excel/excel.chart#id)|Уникальный идентификатор диаграммы.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showAllFieldButtons)|Указывает, следует ли отображать все кнопки поля на сводная диаграмма.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[граница](/javascript/api/excel/excel.chartareaformat#border)|Представляет пограничный формат области диаграммы, включаю в себя цвет, литейный стиль и вес.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (тип: Excel. ChartAxisType, группа?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getItem_type__group_)|Возвращает указанную ось, определенную по типу и группе.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#baseTimeUnit)|Указывает базовый блок для оси указанной категории.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categoryType)|Указывает тип оси категории.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayUnit)|Представляет отображаемую единицу измерения оси.|
||[logBase](/javascript/api/excel/excel.chartaxis#logBase)|Указывает базу логарифма при использовании логарифмических масштабов.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majorTickMark)|Указывает тип основных меток для указанной оси.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majorTimeUnitScale)|Указывает главное значение масштабирования единицы для оси категории при `categoryType` заданном свойстве `dateAxis` .|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minorTickMark)|Указывает тип незначительной метки галочки для указанной оси.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minorTimeUnitScale)|Указывает незначительное значение масштабирования единицы для оси категории при заданном `categoryType` свойстве `dateAxis` .|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisGroup)|Указывает группу для указанной оси.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customDisplayUnit)|Указывает пользовательское значение блока отображения оси.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Указывает высоту оси диаграммы в точках.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Указывает расстояние в точках от левого края оси до левой области диаграммы.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Указывает расстояние в точках от верхнего края оси до верхней области диаграммы.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Указывает тип оси.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Указывает ширину оси диаграммы в точках.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reversePlotOrder)|Указывает, Excel заданы точки данных с последнего до первого.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaleType)|Указывает тип шкалы оси значения.|
||[setCategoryNames (sourceData: Range)](/javascript/api/excel/excel.chartaxis#setCategoryNames_sourceData_)|Устанавливает все имена категорий для указанной оси.|
||[setCustomDisplayUnit (значение: номер)](/javascript/api/excel/excel.chartaxis#setCustomDisplayUnit_value_)|Задает отображаемую единицу измерения оси в виде настраиваемого значения.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showDisplayUnitLabel)|Указывает, видна ли метка блока отображения оси.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#tickLabelPosition)|Указывает положение меток меток на указанной оси.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#tickLabelSpacing)|Указывает количество категорий или рядов между меткими метами.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickMarkSpacing)|Указывает количество категорий или рядов между метками галочки.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Указывает, видна ли ось.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|HTML-код цвета, представляющий цвет границ в диаграмме.|
||[lineStyle](/javascript/api/excel/excel.chartborder#lineStyle)|Представляет тип линии границы.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Представляет толщину границы (в пунктах).|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Значение, которое представляет положение метки данных.|
||[сепаратор](/javascript/api/excel/excel.chartdatalabel#separator)|Строка, представляющая разделитель для метки данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showBubbleSize)|Указывает, виден ли размер пузыря метки данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showCategoryName)|Указывает, отображается ли имя категории метки данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showLegendKey)|Указывает, виден ли ключ легенды метки данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showPercentage)|Указывает, виден ли процент метки данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showSeriesName)|Указывает, отображается ли имя серии меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showValue)|Указывает, отображается ли значение метки данных.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта и цвет объекта символов диаграммы.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Указывает высоту в точках легенды на диаграмме.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Указывает левое значение в точках легенды на диаграмме.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendEntries)|Представляет коллекцию объектов legendEntries в условных обозначениях.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showShadow)|Указывает, имеет ли легенда тень на диаграмме.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Указывает верхнюю часть легенды диаграммы.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Указывает ширину в точках легенды на диаграмме.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Представляет видимость записи легенды диаграммы.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getCount__)|Возвращает количество записей легенды в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getItemAt_index_)|Возвращает запись легенды в заданный индекс.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#lineStyle)|Представляет стиль строки.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Представляет толщину линии (в пунктах).|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasDataLabel)|Представляет, имеет ли точка данных метку данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerBackgroundColor)|Представление цветового кода HTML фонового цвета маркера точки данных (например, #FF0000 представляет красный цвет).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerForegroundColor)|Представление цветового кода HTML маркера переднего плана точки данных (например, #FF0000 представляет красный цвет).|
||[markerSize](/javascript/api/excel/excel.chartpoint#markerSize)|Представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerStyle)|Представляет стиль маркера точки данных диаграммы.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#dataLabel)|Возвращает метку данных точки диаграммы.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[граница](/javascript/api/excel/excel.chartpointformat#border)|Представляет пограничный формат точки данных диаграммы, которая включает сведения о цвете, стиле и весе.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#chartType)|Представляет тип диаграммы для ряда.|
||[delete()](/javascript/api/excel/excel.chartseries#delete__)|Удаляет ряд диаграммы.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutHoleSize)|Представляет размер отверстия ряда кольцевой диаграммы.|
||[отфильтрованный](/javascript/api/excel/excel.chartseries#filtered)|Указывает, фильтруется ли серия.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapWidth)|Представляет ширину разрывов рядов диаграммы.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasDataLabels)|Указывает, есть ли в серии метки данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerBackgroundColor)|Указывает фоновый цвет маркера серии диаграмм.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerForegroundColor)|Указывает цвет маркера переднего плана серии диаграмм.|
||[markerSize](/javascript/api/excel/excel.chartseries#markerSize)|Указывает размер маркера серии диаграмм.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerStyle)|Указывает стиль маркера серии диаграмм.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotOrder)|Указывает порядок сюжета серии диаграмм в группе диаграмм.|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|Коллекция трендовых линий в серии.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#setBubbleSizes_sourceData_)|Задает размеры пузыря для серии диаграмм.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setValues_sourceData_)|Задает значения для серии диаграмм.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setXAxisValues_sourceData_)|Задает значения x-axis для серии диаграмм.|
||[showShadow](/javascript/api/excel/excel.chartseries#showShadow)|Указывает, есть ли в серии тень.|
||[гладкая](/javascript/api/excel/excel.chartseries#smooth)|Указывает, является ли серия гладкой.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#add_name__index_)|Добавляет новый ряд в коллекцию.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#getSubstring_start__length_)|Получите подстройку заголовка диаграммы.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalAlignment)|Указывает горизонтальное выравнивание для заголовка диаграммы.|
||[left](/javascript/api/excel/excel.charttitle#left)|Указывает расстояние в точках от левого края заголовка диаграммы до левого края области диаграммы.|
||[position](/javascript/api/excel/excel.charttitle#position)|Представляет положение заголовка диаграммы.|
||[height](/javascript/api/excel/excel.charttitle#height)|Возвращает высоту заголовка диаграммы (в пунктах).|
||[width](/javascript/api/excel/excel.charttitle#width)|Указывает ширину в точках заголовка диаграммы.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#setFormula_formula_)|Задает строковое значение, представляющее формулу заголовка диаграммы с использованием нотации стиля A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#showShadow)|Представляет логическое значение, которое определяет, имеет ли заголовок диаграммы тень.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textOrientation)|Указывает угол, на который ориентирован текст для заголовка диаграммы.|
||[top](/javascript/api/excel/excel.charttitle#top)|Указывает расстояние в точках от верхнего края заголовка диаграммы до верхней части области диаграммы.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalAlignment)|Указывает вертикальное выравнивание заголовка диаграммы.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[граница](/javascript/api/excel/excel.charttitleformat#border)|Представляет пограничный формат заголовка диаграммы, который включает цвет, линия и вес.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete__)|Удаляет объект линии тренда.|
||[перехват](/javascript/api/excel/excel.charttrendline#intercept)|Представляет значение отсекаемого отрезка линии тренда.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingAveragePeriod)|Представляет период трендовой линии диаграммы.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Представляет имя линии тренда.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialOrder)|Представляет порядок трендовой линии диаграммы.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Представляет форматирование линии тренда диаграммы.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Представляет тип линии тренда диаграммы.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add_type_)|Добавляет новую линию тренда в коллекцию линий тренда.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getCount__)|Возвращает количество линий тренда в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getItem_index_)|Получает объект trendline по индексу, который является порядком вставки в массиве элементов.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Представляет форматирование линий диаграммы.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete__)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.customproperty#key)|Ключ настраиваемого свойства.|
||[type](/javascript/api/excel/excel.customproperty#type)|Тип значения, используемого для настраиваемого свойства.|
||[value](/javascript/api/excel/excel.customproperty#value)|Значение настраиваемого свойства.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add_key__value_)|Создает или задает настраиваемое свойство.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#deleteAll__)|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getCount__)|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getItem_key_)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getItemOrNullObject_key_)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#refreshAll__)|Обновляет все подключения к данным в коллекции.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[автор](/javascript/api/excel/excel.documentproperties#author)|Автор книги.|
||[категория](/javascript/api/excel/excel.documentproperties#category)|Категория книги.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Комментарии книги.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Компания книги.|
||[ключевые слова](/javascript/api/excel/excel.documentproperties#keywords)|Ключевые слова книги.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|Менеджер книги.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationDate)|Получает дату создания книги.|
||[настраиваемый](/javascript/api/excel/excel.documentproperties#custom)|Получает коллекцию настраиваемых свойств книги.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastAuthor)|Получает последнего автора книги.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionNumber)|Получает номер редакции книги.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Тема книги.|
||[заголовок](/javascript/api/excel/excel.documentproperties#title)|Название книги.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Формула названного элемента.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayValues)|Возвращает объект, содержащий значения и типы именованного элемента.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Представляет типы для каждого элемента в массиве именуемого элемента|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Представляет значения каждого элемента в массиве именованных элементов.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: number, numColumns: number)](/javascript/api/excel/excel.range#getAbsoluteResizedRange_numRows__numColumns_)|Получает объект с той же верхней левой ячейкой, что и текущий объект, но с указанным числом `Range` `Range` строк и столбцов.|
||[getImage()](/javascript/api/excel/excel.range#getImage__)|Отрисовка диапазона в качестве изображения png с кодом base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getSurroundingRegion__)|Возвращает `Range` объект, который представляет окружающий регион для верхнего левого элемента в этом диапазоне.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Представляет гиперссылку для текущего диапазона.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberFormatLocal)|Представляет Excel формата номера для данного диапазона в зависимости от языковых параметров пользователя.|
||[isEntireColumn](/javascript/api/excel/excel.range#isEntireColumn)|Указывает, является ли текущий диапазон целым столбцом.|
||[isEntireRow](/javascript/api/excel/excel.range#isEntireRow)|Указывает, является ли текущий диапазон целой строкой.|
||[showCard()](/javascript/api/excel/excel.range#showCard__)|Отображает карточку для активной ячейки, если она имеет содержимое c форматированным значением.|
||[style](/javascript/api/excel/excel.range#style)|Представляет стиль текущего диапазона.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textOrientation)|Текстовая ориентация всех ячеек в диапазоне.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#useStandardHeight)|Определяет, равна ли высота строки объекта `Range` стандартной высоте листа.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#useStandardWidth)|Указывает, равна ли ширина столбца объекту `Range` стандартную ширину листа.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Представляет url-адрес для гиперссылки.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentReference)|Представляет адресную цель документа для гиперссылки.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#screenTip)|Представляет строку, отображаемую при наведении указателя на гиперссылку.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#textToDisplay)|Представляет строку, отображаемую в верхней левой ячейке диапазона.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete__)|Удаляет этот стиль.|
||[formulaHidden](/javascript/api/excel/excel.style#formulaHidden)|Указывает, будет ли формула скрыта при защите таблицы.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalAlignment)|Представляет горизонтальное выравнивание для стиля.|
||[includeAlignment](/javascript/api/excel/excel.style#includeAlignment)|Указывает, включает ли стиль свойства автоотступа, горизонтальное выравнивание, вертикальное выравнивание, текст упаковки, уровень отступа и свойства ориентации текста.|
||[includeBorder](/javascript/api/excel/excel.style#includeBorder)|Указывает, включает ли стиль свойства цвета, индекса цвета, стиля строки и весовых границ.|
||[includeFont](/javascript/api/excel/excel.style#includeFont)|Указывает, включает ли стиль фон, жирный цвет, цвет, индекс цвета, стиль шрифта, italic, имя, размер, strikethrough, subscript, superscript и underline font properties.|
||[includeNumber](/javascript/api/excel/excel.style#includeNumber)|Указывает, включает ли стиль свойство формата номеров.|
||[includePatterns](/javascript/api/excel/excel.style#includePatterns)|Указывает, включает ли стиль свойства цвета, индекса цвета, инверта, если отрицательный, шаблон, цвет шаблона и свойства индекса цвета шаблона.|
||[includeProtection](/javascript/api/excel/excel.style#includeProtection)|Указывает, включает ли стиль скрытые и заблокированные свойства защиты формулы.|
||[indentLevel](/javascript/api/excel/excel.style#indentLevel)|Целое число от 0 до 250, указывающее уровень отступа для стиля.|
||[locked](/javascript/api/excel/excel.style#locked)|Указывает, заблокирован ли объект при защите таблицы.|
||[numberFormat](/javascript/api/excel/excel.style#numberFormat)|Код числового формата для стиля.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberFormatLocal)|Локализованный код числового формата для стиля.|
||[readingOrder](/javascript/api/excel/excel.style#readingOrder)|Направление чтения для стиля.|
||[borders](/javascript/api/excel/excel.style#borders)|Коллекция из четырех пограничных объектов, которые представляют стиль четырех границ.|
||[builtIn](/javascript/api/excel/excel.style#builtIn)|Указывает, является ли стиль встроенным.|
||[fill](/javascript/api/excel/excel.style#fill)|Заполнение стиля.|
||[font](/javascript/api/excel/excel.style#font)|Объект, `Font` который представляет шрифт стиля.|
||[name](/javascript/api/excel/excel.style#name)|Имя стиля.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinkToFit)|Указывает, если текст автоматически сокращается, чтобы соответствовать ширине доступных столбцов.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalAlignment)|Указывает вертикальное выравнивание для стиля.|
||[wrapText](/javascript/api/excel/excel.style#wrapText)|Указывает, Excel обертывание текста в объекте.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add_name_)|Добавляет новый стиль в коллекцию.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getItem_name_)|Получает `Style` имя.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onChanged)|Происходит, когда данные в ячейках меняются на определенной таблице.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onSelectionChanged)|Происходит, когда выбор изменяется на определенной таблице.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Получает адрес, представляющий измененную область таблицы на конкретном листе.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changeType)|Получает тип изменений, который представляет, как запускается измененное событие.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableId)|Получает ID таблицы, в которой изменились данные.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetId)|Получает ID таблицы, в которой изменились данные.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onChanged)|Происходит, когда данные меняются на любой таблице в книге или в таблице.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Получает адрес диапазона, представляющий выбранную область таблицы на конкретном листе.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isInsideTable)|Указывает, находится ли выбор внутри таблицы.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableId)|Получает ID таблицы, в которой изменился выбор.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetId)|Получает ID таблицы, в которой изменен выбор.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getActiveCell__)|Получает текущую активную ячейку из книги.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataConnections)|Представляет все подключения к данным в книге.|
||[name](/javascript/api/excel/excel.workbook#name)|Получает имя книги.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Получает свойства книги.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Возвращает объект защиты для книги.|
||[стили](/javascript/api/excel/excel.workbook#styles)|Представляет коллекцию стилей, связанных с книгой.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#protect_password_)|Защищает книгу.|
||[защищена](/javascript/api/excel/excel.workbookprotection#protected)|Указывает, защищена ли книга.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect_password_)|Снимает защиту с книги.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Таблица)](/javascript/api/excel/excel.worksheet#copy_positionType__relativeTo_)|Копирует таблицу и помещает ее в указанное положение.|
||[getRangeByIndexes (startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getRangeByIndexes_startRow__startColumn__rowCount__columnCount_)|Получает объект, начиная с определенного индекса строки и индекса столбцов, и охватывает определенное количество `Range` строк и столбцов.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezePanes)|Получает объект, который можно использовать для управления замороженными стемнами на таблице.|
||[onActivated](/javascript/api/excel/excel.worksheet#onActivated)|Возникает при активации таблицы.|
||[onChanged](/javascript/api/excel/excel.worksheet#onChanged)|Происходит при изменениях данных в определенном таблице.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#onDeactivated)|Происходит при отключке таблицы.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onSelectionChanged)|Происходит, когда выбор изменяется на определенном таблице.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardHeight)|Возвращает стандартную (по умолчанию) высоту всех строк на листе (в пунктах).|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardWidth)|Указывает стандартную (по умолчанию) ширину всех столбцов в таблице.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabColor)|Цвет таблицы вкладок.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetId)|Получает ID активированного таблицы.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetId)|Получает ID таблицы, добавляемой в книгу.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changeType)|Получает тип изменений, который представляет, как запускается измененное событие.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetId)|Получает ID таблицы, в которой изменились данные.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onActivated)|Возникает при активации любого таблицы в книге.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onAdded)|Возникает при добавлении нового таблицы в книгу.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#onDeactivated)|Происходит при отключке любой таблицы в книге.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#onDeleted)|Возникает при удалении таблицы из книги.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetId)|Получает ID деактивированной таблицы.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetId)|Получает ID таблицы, удаляемой из книги.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#freezeAt_frozenRange_)|Задает закрепленные ячейки в представлении активного листа.|
||[freezeColumns (count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezeColumns_count_)|Замораживание первого столбца или столбцов таблицы на месте.|
||[freezeRows (count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezeRows_count_)|Замораживание верхней строки или строки таблицы на месте.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getLocation__)|Получает диапазон, описывающий закрепленные ячейки в представлении активного листа.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getLocationOrNullObject__)|Получает диапазон, описывающий закрепленные ячейки в представлении активного листа.|
||[разморозка()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze__)|Удаляет все закрепленные области в листе.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect_password_)|Снимает защиту с листа.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#allowEditObjects)|Представляет параметр защиты таблиц, позволяющий изменять объекты.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#allowEditScenarios)|Представляет параметр защиты таблицы, разрешающий редактирование сценариев.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionMode)|Представляет параметр защиты рабочего листа для режима выделения.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Получает адрес диапазона, представляющий выделенную область конкретного листа.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetId)|Получает ID таблицы, в которой изменен выбор.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

---
title: Excel API JavaScript установлено 1.7
description: Сведения о наборе требований ExcelApi 1.7.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
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

В следующей таблице перечислены API в Excel API JavaScript, установленный 1.7. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.7 или ранее, см. Excel API в наборе требований [1.7 или ранее](/javascript/api/excel?view=excel-js-1.7&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#excel-excel-chart-charttype-member)|Указывает тип диаграммы.|
||[id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member)|Уникальный идентификатор диаграммы.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#excel-excel-chart-showallfieldbuttons-member)|Указывает, следует ли отображать все кнопки поля на сводная диаграмма.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[граница](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-border-member)|Представляет пограничный формат области диаграммы, включаю в себя цвет, литейный стиль и вес.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (тип: Excel. ChartAxisType, группа?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-getitem-member(1))|Возвращает указанную ось, определенную по типу и группе.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[axisGroup](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-axisgroup-member)|Указывает группу для указанной оси.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-basetimeunit-member)|Указывает базовый блок для оси указанной категории.|
||[categoryType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-categorytype-member)|Указывает тип оси категории.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-customdisplayunit-member)|Указывает пользовательское значение блока отображения оси.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-displayunit-member)|Представляет отображаемую единицу измерения оси.|
||[height](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-height-member)|Указывает высоту оси диаграммы в точках.|
||[left](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-left-member)|Указывает расстояние в точках от левого края оси до левой области диаграммы.|
||[logBase](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-logbase-member)|Указывает базу логарифма при использовании логарифмических масштабов.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortickmark-member)|Указывает тип основных меток для указанной оси.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortimeunitscale-member)|Указывает главное значение масштабирования единицы для оси категории при `categoryType` заданном свойстве `dateAxis`.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortickmark-member)|Указывает тип незначительной метки галочки для указанной оси.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortimeunitscale-member)|Указывает незначительное значение масштабирования единицы для оси категории при `categoryType` заданном свойстве `dateAxis`.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-reverseplotorder-member)|Указывает, Excel заданы точки данных с последнего до первого.|
||[scaleType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-scaletype-member)|Указывает тип шкалы оси значения.|
||[setCategoryNames (sourceData: Range)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcategorynames-member(1))|Устанавливает все имена категорий для указанной оси.|
||[setCustomDisplayUnit (значение: номер)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcustomdisplayunit-member(1))|Задает отображаемую единицу измерения оси в виде настраиваемого значения.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-showdisplayunitlabel-member)|Указывает, видна ли метка блока отображения оси.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelposition-member)|Указывает положение меток меток на указанной оси.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelspacing-member)|Указывает количество категорий или рядов между меткими метами.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-tickmarkspacing-member)|Указывает количество категорий или рядов между метками галочки.|
||[top](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-top-member)|Указывает расстояние в точках от верхнего края оси до верхней области диаграммы.|
||[type](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-type-member)|Указывает тип оси.|
||[visible](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-visible-member)|Указывает, видна ли ось.|
||[width](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-width-member)|Указывает ширину оси диаграммы в точках.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-color-member)|HTML-код цвета, представляющий цвет границ в диаграмме.|
||[lineStyle](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-linestyle-member)|Представляет тип линии границы.|
||[weight](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-weight-member)|Представляет толщину границы (в пунктах).|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-position-member)|Значение, которое представляет положение метки данных.|
||[сепаратор](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-separator-member)|Строка, представляющая разделитель для метки данных на диаграмме.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showbubblesize-member)|Указывает, виден ли размер пузыря метки данных.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showcategoryname-member)|Указывает, отображается ли имя категории метки данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showlegendkey-member)|Указывает, виден ли ключ легенды метки данных.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showpercentage-member)|Указывает, виден ли процент метки данных.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showseriesname-member)|Указывает, отображается ли имя серии меток данных.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showvalue-member)|Указывает, отображается ли значение метки данных.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#excel-excel-chartformatstring-font-member)|Представляет атрибуты шрифта, такие как имя шрифта, размер шрифта и цвет объекта символов диаграммы.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-height-member)|Указывает высоту в точках легенды на диаграмме.|
||[left](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-left-member)|Указывает левое значение в точках легенды на диаграмме.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-legendentries-member)|Представляет коллекцию объектов legendEntries в условных обозначениях.|
||[showShadow](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-showshadow-member)|Указывает, имеет ли легенда тень на диаграмме.|
||[top](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-top-member)|Указывает верхнюю часть легенды диаграммы.|
||[width](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-width-member)|Указывает ширину в точках легенды на диаграмме.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-visible-member)|Представляет видимость записи легенды диаграммы.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getcount-member(1))|Возвращает количество записей легенды в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getitemat-member(1))|Возвращает запись легенды в заданный индекс.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-linestyle-member)|Представляет стиль строки.|
||[weight](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-weight-member)|Представляет толщину линии (в пунктах).|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[dataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-datalabel-member)|Возвращает метку данных точки диаграммы.|
||[hasDataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-hasdatalabel-member)|Представляет, имеет ли точка данных метку данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerbackgroundcolor-member)|Представление цветового кода HTML фонового цвета маркера точки данных (например, #FF0000 представляет красный цвет).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerforegroundcolor-member)|Представление цветового кода HTML маркера переднего плана точки данных (например, #FF0000 представляет красный цвет).|
||[markerSize](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markersize-member)|Представляет размер маркера точки данных.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerstyle-member)|Представляет стиль маркера точки данных диаграммы.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[граница](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-border-member)|Представляет пограничный формат точки данных диаграммы, которая включает сведения о цвете, стиле и весе.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-charttype-member)|Представляет тип диаграммы для ряда.|
||[delete()](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-delete-member(1))|Удаляет ряд диаграммы.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-doughnutholesize-member)|Представляет размер отверстия ряда кольцевой диаграммы.|
||[отфильтрованный](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-filtered-member)|Указывает, фильтруется ли серия.|
||[gapWidth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gapwidth-member)|Представляет ширину разрывов рядов диаграммы.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-hasdatalabels-member)|Указывает, есть ли в серии метки данных.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerbackgroundcolor-member)|Указывает фоновый цвет маркера серии диаграмм.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerforegroundcolor-member)|Указывает цвет маркера переднего плана серии диаграмм.|
||[markerSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markersize-member)|Указывает размер маркера серии диаграмм.|
||[markerStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerstyle-member)|Указывает стиль маркера серии диаграмм.|
||[plotOrder](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-plotorder-member)|Указывает порядок сюжета серии диаграмм в группе диаграмм.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setbubblesizes-member(1))|Задает размеры пузыря для серии диаграмм.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setvalues-member(1))|Задает значения для серии диаграмм.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setxaxisvalues-member(1))|Задает значения x-axis для серии диаграмм.|
||[showShadow](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showshadow-member)|Указывает, есть ли в серии тень.|
||[гладкая](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-smooth-member)|Указывает, является ли серия гладкой.|
||[trendlines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-trendlines-member)|Коллекция трендовых линий в серии.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-add-member(1))|Добавляет новый ряд в коллекцию.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-getsubstring-member(1))|Получите подстройку заголовка диаграммы.|
||[height](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-height-member)|Возвращает высоту заголовка диаграммы (в пунктах).|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-horizontalalignment-member)|Указывает горизонтальное выравнивание для заголовка диаграммы.|
||[left](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-left-member)|Указывает расстояние в точках от левого края заголовка диаграммы до левого края области диаграммы.|
||[position](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-position-member)|Представляет положение заголовка диаграммы.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-setformula-member(1))|Задает строковое значение, представляющее формулу заголовка диаграммы с использованием нотации стиля A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-showshadow-member)|Представляет логическое значение, которое определяет, имеет ли заголовок диаграммы тень.|
||[textOrientation](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-textorientation-member)|Указывает угол, на который ориентирован текст для заголовка диаграммы.|
||[top](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-top-member)|Указывает расстояние в точках от верхнего края заголовка диаграммы до верхней части области диаграммы.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-verticalalignment-member)|Указывает вертикальное выравнивание заголовка диаграммы.|
||[width](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-width-member)|Указывает ширину в точках заголовка диаграммы.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[граница](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-border-member)|Представляет пограничный формат заголовка диаграммы, который включает цвет, линия и вес.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-delete-member(1))|Удаляет объект линии тренда.|
||[format](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-format-member)|Представляет форматирование линии тренда диаграммы.|
||[перехват](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-intercept-member)|Представляет значение отсекаемого отрезка линии тренда.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-movingaverageperiod-member)|Представляет период трендовой линии диаграммы.|
||[name](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-name-member)|Представляет имя линии тренда.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-polynomialorder-member)|Представляет порядок трендовой линии диаграммы.|
||[type](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-type-member)|Представляет тип линии тренда диаграммы.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-add-member(1))|Добавляет новую линию тренда в коллекцию линий тренда.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getcount-member(1))|Возвращает количество линий тренда в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getitem-member(1))|Получает объект trendline по индексу, который является порядком вставки в массиве элементов.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#excel-excel-charttrendlineformat-line-member)|Представляет форматирование линий диаграммы.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-delete-member(1))|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-key-member)|Ключ настраиваемого свойства.|
||[type](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-type-member)|Тип значения, используемого для настраиваемого свойства.|
||[value](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-value-member)|Значение настраиваемого свойства.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-add-member(1))|Создает или задает настраиваемое свойство.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-deleteall-member(1))|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getcount-member(1))|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitem-member(1))|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitemornullobject-member(1))|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#excel-excel-dataconnectioncollection-refreshall-member(1))|Обновляет все подключения к данным в коллекции.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[автор](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-author-member)|Автор книги.|
||[категория](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-category-member)|Категория книги.|
||[comments](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-comments-member)|Комментарии книги.|
||[company](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-company-member)|Компания книги.|
||[creationDate](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-creationdate-member)|Получает дату создания книги.|
||[настраиваемый](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member)|Получает коллекцию настраиваемых свойств книги.|
||[ключевые слова](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-keywords-member)|Ключевые слова книги.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-lastauthor-member)|Получает последнего автора книги.|
||[manager](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-manager-member)|Менеджер книги.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-revisionnumber-member)|Получает номер редакции книги.|
||[subject](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-subject-member)|Тема книги.|
||[заголовок](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-title-member)|Название книги.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[arrayValues](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-arrayvalues-member)|Возвращает объект, содержащий значения и типы именованного элемента.|
||[formula](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-formula-member)|Формула названного элемента.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-types-member)|Представляет типы для каждого элемента в массиве именуемого элемента|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-values-member)|Представляет значения каждого элемента в массиве именованных элементов.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: number, numColumns: number)](/javascript/api/excel/excel.range#excel-excel-range-getabsoluteresizedrange-member(1))|Получает объект `Range` с той же верхней левой `Range` ячейкой, что и текущий объект, но с указанным числом строк и столбцов.|
||[getImage()](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1))|Отрисовка диапазона в качестве изображения png с кодом base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#excel-excel-range-getsurroundingregion-member(1))|Возвращает объект, `Range` который представляет окружающий регион для верхнего левого элемента в этом диапазоне.|
||[hyperlink](/javascript/api/excel/excel.range#excel-excel-range-hyperlink-member)|Представляет гиперссылку для текущего диапазона.|
||[isEntireColumn](/javascript/api/excel/excel.range#excel-excel-range-isentirecolumn-member)|Указывает, является ли текущий диапазон целым столбцом.|
||[isEntireRow](/javascript/api/excel/excel.range#excel-excel-range-isentirerow-member)|Указывает, является ли текущий диапазон целой строкой.|
||[numberFormatLocal](/javascript/api/excel/excel.range#excel-excel-range-numberformatlocal-member)|Представляет Excel формата номера для данного диапазона в зависимости от языковых параметров пользователя.|
||[showCard()](/javascript/api/excel/excel.range#excel-excel-range-showcard-member(1))|Отображает карточку для активной ячейки, если она имеет содержимое c форматированным значением.|
||[style](/javascript/api/excel/excel.range#excel-excel-range-style-member)|Представляет стиль текущего диапазона.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-textorientation-member)|Текстовая ориентация всех ячеек в диапазоне.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardheight-member)|Определяет, равна ли высота строки объекта `Range` стандартной высоте листа.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardwidth-member)|Указывает, равна ли ширина столбца объекту `Range` стандартную ширину листа.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-address-member)|Представляет url-адрес для гиперссылки.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-documentreference-member)|Представляет адресную цель документа для гиперссылки.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-screentip-member)|Представляет строку, отображаемую при наведении указателя на гиперссылку.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-texttodisplay-member)|Представляет строку, отображаемую в верхней левой ячейке диапазона.|
|[Style](/javascript/api/excel/excel.style)|[borders](/javascript/api/excel/excel.style#excel-excel-style-borders-member)|Коллекция из четырех пограничных объектов, которые представляют стиль четырех границ.|
||[builtIn](/javascript/api/excel/excel.style#excel-excel-style-builtin-member)|Указывает, является ли стиль встроенным.|
||[delete()](/javascript/api/excel/excel.style#excel-excel-style-delete-member(1))|Удаляет этот стиль.|
||[fill](/javascript/api/excel/excel.style#excel-excel-style-fill-member)|Заполнение стиля.|
||[font](/javascript/api/excel/excel.style#excel-excel-style-font-member)|Объект `Font` , который представляет шрифт стиля.|
||[formulaHidden](/javascript/api/excel/excel.style#excel-excel-style-formulahidden-member)|Указывает, будет ли формула скрыта при защите таблицы.|
||[horizontalAlignment](/javascript/api/excel/excel.style#excel-excel-style-horizontalalignment-member)|Представляет горизонтальное выравнивание для стиля.|
||[includeAlignment](/javascript/api/excel/excel.style#excel-excel-style-includealignment-member)|Указывает, включает ли стиль свойства автоотступа, горизонтальное выравнивание, вертикальное выравнивание, текст упаковки, уровень отступа и свойства ориентации текста.|
||[includeBorder](/javascript/api/excel/excel.style#excel-excel-style-includeborder-member)|Указывает, включает ли стиль свойства цвета, индекса цвета, стиля строки и весовых границ.|
||[includeFont](/javascript/api/excel/excel.style#excel-excel-style-includefont-member)|Указывает, включает ли стиль фон, жирный цвет, цвет, индекс цвета, стиль шрифта, italic, имя, размер, strikethrough, subscript, superscript и underline font properties.|
||[includeNumber](/javascript/api/excel/excel.style#excel-excel-style-includenumber-member)|Указывает, включает ли стиль свойство формата номеров.|
||[includePatterns](/javascript/api/excel/excel.style#excel-excel-style-includepatterns-member)|Указывает, включает ли стиль свойства цвета, индекса цвета, инверта, если отрицательный, шаблон, цвет шаблона и свойства индекса цвета шаблона.|
||[includeProtection](/javascript/api/excel/excel.style#excel-excel-style-includeprotection-member)|Указывает, включает ли стиль скрытые и заблокированные свойства защиты формулы.|
||[indentLevel](/javascript/api/excel/excel.style#excel-excel-style-indentlevel-member)|Целое число от 0 до 250, указывающее уровень отступа для стиля.|
||[locked](/javascript/api/excel/excel.style#excel-excel-style-locked-member)|Указывает, заблокирован ли объект при защите таблицы.|
||[name](/javascript/api/excel/excel.style#excel-excel-style-name-member)|Имя стиля.|
||[numberFormat](/javascript/api/excel/excel.style#excel-excel-style-numberformat-member)|Код числового формата для стиля.|
||[numberFormatLocal](/javascript/api/excel/excel.style#excel-excel-style-numberformatlocal-member)|Локализованный код числового формата для стиля.|
||[readingOrder](/javascript/api/excel/excel.style#excel-excel-style-readingorder-member)|Направление чтения для стиля.|
||[shrinkToFit](/javascript/api/excel/excel.style#excel-excel-style-shrinktofit-member)|Указывает, если текст автоматически сокращается, чтобы соответствовать ширине доступных столбцов.|
||[verticalAlignment](/javascript/api/excel/excel.style#excel-excel-style-verticalalignment-member)|Указывает вертикальное выравнивание для стиля.|
||[wrapText](/javascript/api/excel/excel.style#excel-excel-style-wraptext-member)|Указывает, Excel обертывание текста в объекте.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-add-member(1))|Добавляет новый стиль в коллекцию.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitem-member(1))|Получает имя `Style` .|
||[items](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member)|Происходит, когда данные в ячейках меняются на определенной таблице.|
||[onSelectionChanged](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)|Происходит, когда выбор изменяется на определенной таблице.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-address-member)|Получает адрес, представляющий измененную область таблицы на конкретном листе.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-changetype-member)|Получает тип изменений, который представляет, как запускается измененное событие.|
||[источник](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-source-member)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-tableid-member)|Получает ID таблицы, в которой изменились данные.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-worksheetid-member)|Получает ID таблицы, в которой изменились данные.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member)|Происходит, когда данные меняются на любой таблице в книге или в таблице.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-address-member)|Получает адрес диапазона, представляющий выбранную область таблицы на конкретном листе.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-isinsidetable-member)|Указывает, находится ли выбор внутри таблицы.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-tableid-member)|Получает ID таблицы, в которой изменился выбор.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-worksheetid-member)|Получает ID таблицы, в которой изменен выбор.|
|[Workbook](/javascript/api/excel/excel.workbook)|[dataConnections](/javascript/api/excel/excel.workbook#excel-excel-workbook-dataconnections-member)|Представляет все подключения к данным в книге.|
||[getActiveCell()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivecell-member(1))|Получает текущую активную ячейку из книги.|
||[name](/javascript/api/excel/excel.workbook#excel-excel-workbook-name-member)|Получает имя книги.|
||[properties](/javascript/api/excel/excel.workbook#excel-excel-workbook-properties-member)|Получает свойства книги.|
||[protection](/javascript/api/excel/excel.workbook#excel-excel-workbook-protection-member)|Возвращает объект защиты для книги.|
||[стили](/javascript/api/excel/excel.workbook#excel-excel-workbook-styles-member)|Представляет коллекцию стилей, связанных с книгой.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protect-member(1))|Защищает книгу.|
||[защищена](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protected-member)|Указывает, защищена ли книга.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-unprotect-member(1))|Снимает защиту с книги.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Таблица)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-copy-member(1))|Копирует таблицу и помещает ее в указанное положение.|
||[freezePanes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezepanes-member)|Получает объект, который можно использовать для управления замороженными стемнами на таблице.|
||[getRangeByIndexes (startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrangebyindexes-member(1))|Получает объект `Range` , начиная с определенного индекса строки и индекса столбцов, и охватывает определенное количество строк и столбцов.|
||[onActivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member)|Возникает при активации таблицы.|
||[onChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member)|Происходит при изменениях данных в определенном таблице.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)|Происходит при отключке таблицы.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)|Происходит, когда выбор изменяется на определенном таблице.|
||[standardHeight](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardheight-member)|Возвращает стандартную (по умолчанию) высоту всех строк на листе (в пунктах).|
||[standardWidth](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardwidth-member)|Указывает стандартную (по умолчанию) ширину всех столбцов в таблице.|
||[tabColor](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabcolor-member)|Цвет таблицы вкладок.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-worksheetid-member)|Получает ID активированного таблицы.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[источник](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-worksheetid-member)|Получает ID таблицы, добавляемой в книгу.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-address-member)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changetype-member)|Получает тип изменений, который представляет, как запускается измененное событие.|
||[источник](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-worksheetid-member)|Получает ID таблицы, в которой изменились данные.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member)|Возникает при активации любого таблицы в книге.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member)|Возникает при добавлении нового таблицы в книгу.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member)|Происходит при отключке любой таблицы в книге.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member)|Возникает при удалении таблицы из книги.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-worksheetid-member)|Получает ID деактивированной таблицы.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[источник](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-worksheetid-member)|Получает ID таблицы, удаляемой из книги.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezeat-member(1))|Задает закрепленные ячейки в представлении активного листа.|
||[freezeColumns (count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezecolumns-member(1))|Замораживание первого столбца или столбцов таблицы на месте.|
||[freezeRows (count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezerows-member(1))|Замораживание верхней строки или строки таблицы на месте.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocation-member(1))|Получает диапазон, описывающий закрепленные ячейки в представлении активного листа.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocationornullobject-member(1))|Получает диапазон, описывающий закрепленные ячейки в представлении активного листа.|
||[разморозка()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-unfreeze-member(1))|Удаляет все закрепленные области в листе.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-unprotect-member(1))|Снимает защиту с листа.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditobjects-member)|Представляет параметр защиты таблиц, позволяющий изменять объекты.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditscenarios-member)|Представляет параметр защиты таблицы, разрешающий редактирование сценариев.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-selectionmode-member)|Представляет параметр защиты рабочего листа для режима выделения.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-address-member)|Получает адрес диапазона, представляющий выделенную область конкретного листа.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-worksheetid-member)|Получает ID таблицы, в которой изменен выбор.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

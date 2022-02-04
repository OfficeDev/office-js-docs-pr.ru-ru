---
title: Excel API JavaScript установлено 1.8
description: Сведения о наборе требований ExcelApi 1.8.
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-18"></a>Новые возможности в Excel API JavaScript 1.8

Функции набора обязательных элементов API JavaScript для Excel 1.8 включают API для сводных таблиц, проверку данных, диаграммы, события для диаграмм, параметры производительности и создание рабочей книги.

## <a name="pivottable"></a>Сводная таблица

Этап 2 для API сводной таблицы позволяет надстройкам устанавливать иерархии сводной таблицы. Теперь вы можете управлять данными и способом их сведения. Наша [статья о сводной таблице](../../excel/excel-add-ins-pivottables.md) содержит дополнительные сведения о новых функциональных возможностях сводной таблицы.

## <a name="data-validation"></a>Проверка данных

Проверка данных позволяет управлять данными, которые вводит в лист пользователь. Вы можете ограничить ячейки предопределенными наборами ответов или задать всплывающие предупреждения о нежелательном вводе. Узнайте больше о [добавлении проверки данных в диапазоны](../../excel/excel-add-ins-data-validation.md) уже сегодня.

## <a name="charts"></a>Диаграммы

Еще один этап выпуска API диаграмм обеспечивает дополнительный программный контроль над элементами диаграммы. Теперь у вас есть расширенный доступ к условным обозначениям, осям, линии тренда и области построения.

## <a name="events"></a>События

Для диаграмм добавлены [дополнительные](../../excel/excel-add-ins-events.md) события. Пусть ваша надстройка реагирует на взаимодействие пользователей с диаграммой. Вы также можете [включать и отключать события](../../excel/performance.md#enable-and-disable-events), запускаемые во всей книге.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.8. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.8 или ранее, см. в Excel API в наборе требований [1.8 или ранее](/javascript/api/excel?view=excel-js-1.8&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula1-member)|Указывает операнд правой руки, когда свойство оператора задано двоичному оператору, такому как GreaterThan (левая операнд — это значение, в который пользователь пытается ввести в ячейку).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula2-member)|С помощью ternary operators Between and NotBetween указывается верхний операнд.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-operator-member)|Оператор, используемый для проверки данных.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#excel-excel-chart-categorylabellevel-member)|Указывает константу индексации уровня метки категорий диаграммы, ссылаясь на уровень меток исходных категорий.|
||[displayBlanksAs](/javascript/api/excel/excel.chart#excel-excel-chart-displayblanksas-member)|Указывает, как пустые ячейки заданы на диаграмме.|
||[onActivated](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member)|Возникает при активации диаграммы.|
||[onDeactivated](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member)|Происходит, когда диаграмма отключена.|
||[plotArea](/javascript/api/excel/excel.chart#excel-excel-chart-plotarea-member)|Представляет область сюжета для диаграммы.|
||[plotBy](/javascript/api/excel/excel.chart#excel-excel-chart-plotby-member)|Определяет способ использования столбцов или строк в качестве рядов данных на диаграмме.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#excel-excel-chart-plotvisibleonly-member)|True, если отображаются только видимые ячейки.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#excel-excel-chart-seriesnamelevel-member)|Указывает константу индексации имен на уровне серии диаграмм, ссылаясь на уровень имен исходных серий.|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#excel-excel-chart-showdatalabelsovermaximum-member)|Указывает, следует ли показывать метки данных, если значение превышает максимальное значение оси значения.|
||[style](/javascript/api/excel/excel.chart#excel-excel-chart-style-member)|Указывает стиль диаграммы для диаграммы.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-chartid-member)|Получает ID активированной диаграммы.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-worksheetid-member)|Получает ID таблицы, в которой активируется диаграмма.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-chartid-member)|Получает ID диаграммы, добавляемой в таблицу.|
||[источник](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-worksheetid-member)|Получает ID таблицы, в которую добавляется диаграмма.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[выравнивание](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-alignment-member)|Указывает выравнивание для указанной метки тик оси.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-isbetweencategories-member)|Указывает, пересекает ли ось значения ось категории между категориями.|
||[multiLevel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-multilevel-member)|Указывает, многоуровневая ли ось.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-numberformat-member)|Указывает код формата для метки тик оси.|
||[смещение](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-offset-member)|Указывает расстояние между уровнями меток и расстоянием между первым уровнем и линией оси.|
||[position](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-position-member)|Указывает указанное положение оси, где пересекается другая ось.|
||[positionAt](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-positionat-member)|Указывает положение оси, где пересекается другая ось.|
||[setPositionAt (значение: номер)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setpositionat-member(1))|Задает указанное положение оси, где пересекается другая ось.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-textorientation-member)|Указывает угол, на который ориентирован текст для метки тика оси диаграммы.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-fill-member)|Указывает форматирование заполнения диаграммы.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-setformula-member(1))|Строковое значение, представляющее формулу заголовка оси диаграммы с использованием нотации стиля A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[граница](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-border-member)|Указывает пограничный формат заголовка оси диаграммы, который включает цвет, листил и вес.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-fill-member)|Указывает форматирование заполнения заголовок оси диаграммы.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-clear-member(1))|Очищает формат границы элемента диаграммы.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member)|Возникает при активации диаграммы.|
||[onAdded](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member)|Возникает при добавлении новой диаграммы в таблицу.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member)|Происходит, когда диаграмма отключена.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member)|Возникает при удалении диаграммы.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-autotext-member)|Указывает, автоматически ли метка данных создает соответствующий текст на основе контекста.|
||[format](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-format-member)|Представляет формат метки данных диаграммы.|
||[formula](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-formula-member)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[height](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-height-member)|Возвращает высоту метки данных диаграммы (в пунктах).|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-horizontalalignment-member)|Представляет горизонтальное выравнивание для метки данных диаграммы.|
||[left](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-left-member)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах). |
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-numberformat-member)|Строковое значение, представляющее код формата для метки данных.|
||[text](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-text-member)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-textorientation-member)|Представляет угол, на который ориентирован текст для метки данных диаграммы.|
||[top](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-top-member)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-verticalalignment-member)|Представляет вертикальное выравнивание для метки данных диаграммы.|
||[width](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-width-member)|Возвращает ширину метки данных диаграммы (в пунктах).|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[граница](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-border-member)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-autotext-member)|Указывает, автоматически ли метки данных создают соответствующий текст на основе контекста.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-horizontalalignment-member)|Указывает горизонтальное выравнивание для метки данных диаграммы.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-numberformat-member)|Указывает код формата для меток данных.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-textorientation-member)|Представляет угол, на который ориентирован текст для меток данных.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-verticalalignment-member)|Представляет вертикальное выравнивание для метки данных диаграммы.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-chartid-member)|Получает ID отключаемой диаграммы.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-worksheetid-member)|Получает ID таблицы, в которой деактивируется диаграмма.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-chartid-member)|Получает ID диаграммы, удаляемой из таблицы.|
||[источник](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-worksheetid-member)|Получает ID таблицы, в которой удаляется диаграмма.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-height-member)|Указывает высоту записи легенды в легенде диаграммы.|
||[индекс](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-index-member)|Указывает индекс записи легенды в легенде диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-left-member)|Указывает левое значение записи легенды диаграммы.|
||[top](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-top-member)|Указывает верхнюю часть записи легенды диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-width-member)|Представляет ширину записи легенды на диаграмме Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[граница](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-border-member)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[format](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-format-member)|Указывает форматирование области сюжета диаграммы.|
||[height](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-height-member)|Указывает значение высоты области участка.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideheight-member)|Указывает внутреннее значение высоты области участка.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideleft-member)|Указывает внутреннее левое значение области сюжета.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidetop-member)|Указывает внутреннее верхнее значение области сюжета.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidewidth-member)|Указывает внутреннее значение ширины области участка.|
||[left](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-left-member)|Указывает левое значение области сюжета.|
||[position](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-position-member)|Указывает положение области сюжета.|
||[top](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-top-member)|Указывает верхнее значение области сюжета.|
||[width](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-width-member)|Указывает значение ширины области участка.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[граница](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-border-member)|Указывает атрибуты границы области диаграммы.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-fill-member)|Указывает формат заполнения объекта, который включает сведения о формате фона.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-axisgroup-member)|Указывает группу для указанной серии.|
||[dataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-datalabels-member)|Представляет коллекцию всех меток данных в серии.|
||[взрыв](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-explosion-member)|Указывает значение взрыва для среза круговой диаграммы или пончик-диаграммы.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-firstsliceangle-member)|Указывает угол первого среза круговой диаграммы или пончик-диаграммы в градусах (по часовой стрелке от вертикальной).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertifnegative-member)|Верно, Excel выверяет шаблон в элементе, если он соответствует отрицательному номеру.|
||[перекрытие](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-overlap-member)|Указывает на расположение строк и столбцов.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-secondplotsize-member)|Указывает размер вторичного раздела диаграммы пирога или диаграммы с круговым пирогом в процентах от размера первичного пирога.|
||[splitType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splittype-member)|Указывает способ разделения двух разделов диаграммы "пирог-пирог" или диаграммы "планка пирога".|
||[varyByCategories](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-varybycategories-member)|True, Excel назначит каждому маркеру данных другой цвет или шаблон.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-backwardperiod-member)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-forwardperiod-member)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[метка](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-label-member)|Представляет метку линии тренда диаграммы.|
||[showEquation](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showequation-member)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showrsquared-member)|Значение True, если значение r-squared для линии тренда отображается на диаграмме.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-autotext-member)|Указывает, автоматически ли метка trendline создает соответствующий текст на основе контекста.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-format-member)|Формат метки трендовой линии диаграммы.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-formula-member)|Строковая величина, которая представляет формулу метки трендовой линии диаграммы с помощью нотации в стиле A1.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-height-member)|Возвращает высоту подписи линии тренда диаграммы (в пунктах).|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-horizontalalignment-member)|Представляет горизонтальное выравнивание метки трендовой линии диаграммы.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-left-member)|Представляет расстояние в точках от левого края метки трендовой линии диаграммы до левого края области диаграммы.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-numberformat-member)|Строковое значение, которое представляет код формата для метки trendline.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-text-member)|Строка, представляющая текст подписи линии тренда на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-textorientation-member)|Представляет угол, на который ориентирован текст для метки трендовой линии диаграммы.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-top-member)|Представляет расстояние в точках от верхнего края метки трендовой линии диаграммы до верхней части области диаграммы.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-verticalalignment-member)|Представляет вертикальное выравнивание метки трендовой линии диаграммы.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-width-member)|Возвращает ширину подписи линии тренда диаграммы (в пунктах).|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[граница](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-border-member)|Указывает пограничный формат, который включает цвет, литейный стил и вес.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-fill-member)|Указывает формат заполнения текущей метки трендовой линии диаграммы.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-font-member)|Указывает атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для метки трендовой линии диаграммы.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#excel-excel-customdatavalidation-formula-member)|Формула проверки настраиваемых данных.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[поле](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-field-member)|Возвращает сводные поля, связанные с DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-id-member)|ID of the DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-name-member)|Имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-numberformat-member)|Числовой формат DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-position-member)|Положение DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-settodefault-member(1))|Сбрасывает DataPivotHierarchy до значений по умолчанию.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-showas-member)|Указывает, следует ли показывать данные в качестве определенного суммарного вычисления.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-summarizeby-member)|Указывает, показаны ли все элементы DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-add-member(1))|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getcount-member(1))|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitem-member(1))|Получает DataPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitemornullobject-member(1))|Получает DataPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-remove-member(1))|Удаляет PivotHierarchy из текущей оси.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-clear-member(1))|Очищает проверку данных из текущего диапазона.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-erroralert-member)|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-ignoreblanks-member)|Указывает, будет ли проверка данных выполняться на пустых ячейках.|
||[сообщение](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-prompt-member)|Подсказка, когда пользователи выбирают ячейку.|
||[правило](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-rule-member)|Правило проверки данных, которое содержит различные типы критериев проверки данных.|
||[type](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-type-member)|Тип проверки данных см. в `Excel.DataValidationType` подробностях.|
||[допустимо](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-valid-member)|Указывает, являются ли все значения ячеек допустимыми в соответствии с правилами проверки данных.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-message-member)|Представляет сообщение оповещений об ошибке.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-showalert-member)|Указывает, следует ли показывать диалоговое окно оповещения об ошибке при вводе пользователем недействительных данных.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-style-member)|Тип оповещений о проверке данных см. в `Excel.DataValidationAlertStyle` подробной информации.|
||[заголовок](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-title-member)|Представляет название диалоговое окно оповещений об ошибке.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-message-member)|Указывает сообщение запроса.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-showprompt-member)|Указывает, отображается ли подсказка, когда пользователь выбирает ячейку с проверкой данных.|
||[заголовок](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-title-member)|Указывает заголовок для запроса.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[настраиваемый](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-custom-member)|Условия проверки настраиваемых данных.|
||[дата](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-date-member)|Условия проверки данных даты.|
||[десятичной](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-decimal-member)|Условия проверки десятичных данных.|
||[list](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-list-member)|Условия проверки данных списка.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-textlength-member)|Критерии проверки данных длины текста.|
||[time](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-time-member)|Условия проверки данных времени.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-wholenumber-member)|Все критерии проверки данных номеров.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula1-member)|Указывает операнд правой руки, когда свойство оператора задано двоичному оператору, такому как GreaterThan (левая операнд — это значение, в который пользователь пытается ввести в ячейку).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula2-member)|С помощью ternary operators Between and NotBetween указывается верхний операнд.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-operator-member)|Оператор, используемый для проверки данных.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-enablemultiplefilteritems-member)|Определяет, следует ли разрешить несколько элементов фильтра.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-fields-member)|Возвращает сводные поля, связанные с FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-id-member)|ID of the FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-name-member)|Имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-position-member)|Положение FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-settodefault-member(1))|Сбрасывает FilterPivotHierarchy до значений по умолчанию.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-add-member(1))|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getcount-member(1))|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitem-member(1))|Получает filterPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitemornullobject-member(1))|Получает FilterPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-remove-member(1))|Удаляет PivotHierarchy из текущей оси.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-incelldropdown-member)|Указывает, следует ли отображать список в выпадаемой ячейке.|
||[source](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-source-member)|Источник списка для проверки данных|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[id](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-id-member)|ID of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-items-member)|Возвращает pivotItems, связанные с PivotField.|
||[name](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-name-member)|Имя сводного поля.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-showallitems-member)|Определяет, следует ли отображать все элементы сводного поля.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbylabels-member(1))|Сортирует сводное поле.|
||[subtotals](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-subtotals-member)|Промежуточные итоги сводного поля.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getcount-member(1))|Получает количество поворотных полей в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitem-member(1))|Получает PivotField по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitemornullobject-member(1))|Получает PivotField по имени.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[fields](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-fields-member)|Возвращает сводные поля, связанные с PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-id-member)|ID of the PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-name-member)|Имя PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getcount-member(1))|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitem-member(1))|Получает PivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitemornullobject-member(1))|Получает PivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[id](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-id-member)|ID of the PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-isexpanded-member)|Определяет, развернут ли элемент для отображения дочерних элементов или же свернут, а дочерние элементы являются скрытыми.|
||[name](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-name-member)|Имя элемента сводной таблицы.|
||[visible](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-visible-member)|Указывает, отображается ли pivotItem.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getcount-member(1))|Получает число pivotItems в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitem-member(1))|Получает PivotItem по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitemornullobject-member(1))|Получает PivotItem по имени.|
||[items](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcolumnlabelrange-member(1))|Возвращает диапазон, где находятся названия столбцов сводной таблицы.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatabodyrange-member(1))|Возвращает диапазон, где находятся значения данных сводной таблицы.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getfilteraxisrange-member(1))|Возвращает диапазон области фильтра сводной таблицы.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrange-member(1))|Возвращает диапазон, в котором существует сводная таблица, за исключением области фильтра.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrowlabelrange-member(1))|Возвращает диапазон, где находятся названия строк сводной таблицы.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-layouttype-member)|Это свойство указывает PivotLayoutType всех полей в сводной таблице.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showcolumngrandtotals-member)|Указывает, показывает ли отчет PivotTable общие итоги для столбцов.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showrowgrandtotals-member)|Указывает, показывает ли отчет PivotTable общие итоги для строк.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-subtotallocation-member)|Это свойство указывает все `SubtotalLocationType` поля на PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[columnHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-columnhierarchies-member)|Иерархии сводных столбцов сводной таблицы.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-datahierarchies-member)|Иерархии сводных данных сводной таблицы.|
||[delete()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-delete-member(1))|Удаляет сводную таблицу.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-filterhierarchies-member)|Иерархии сводных фильтров сводной таблицы.|
||[иерархии](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-hierarchies-member)|Иерархии сводного документа сводной таблицы.|
||[макет](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-layout-member)|PivotLayout, описывающий макет и визуальную структуру сводной таблицы.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-rowhierarchies-member)|Иерархии сводных строк сводной таблицы.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-add-member(1))|Добавьте pivotTable на основе указанных исходных данных и вставьте его в верхней левой ячейке диапазона назначения.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#excel-excel-range-datavalidation-member)|Возвращает объект проверки данных.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-fields-member)|Возвращает сводные поля, связанные с RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-id-member)|ID of the RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-name-member)|Имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-position-member)|Положение RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-settodefault-member(1))|Сбрасывает RowColumnPivotHierarchy до значений по умолчанию.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-add-member(1))|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getcount-member(1))|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitem-member(1))|Получает RowColumnPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitemornullobject-member(1))|Получает RowColumnPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-remove-member(1))|Удаляет PivotHierarchy из текущей оси.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#excel-excel-runtime-enableevents-member)|Добавление событий JavaScript в текущую области задач или надстройку контента.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-basefield-member)|PivotField на основе `ShowAs` расчета, если применимо в соответствии с типом `ShowAsCalculation` , еще `null`.|
||[baseItem](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-baseitem-member)|Элемент, на основе расчета `ShowAs` , если применимо в соответствии с типом `ShowAsCalculation` , еще `null`.|
||[вычисление](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-calculation-member)|Вычисление `ShowAs` , используемого для PivotField.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#excel-excel-style-autoindent-member)|Указывает, будет ли текст автоматически отступным, если выравнивание текста в ячейке задано на равное распределение.|
||[textOrientation](/javascript/api/excel/excel.style#excel-excel-style-textorientation-member)|Ориентация текста для стиля.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-automatic-member)|Если `Automatic` установлено значение `true`, все остальные значения будут игнорироваться при настройке `Subtotals`.|
||[среднее значение](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-average-member)||
||[count](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-count-member)||
||[countNumbers](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-countnumbers-member)||
||[max](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-max-member)||
||[min](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-min-member)||
||[продукт](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-product-member)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviation-member)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviationp-member)||
||[sum](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-sum-member)||
||[отклонение](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variance-member)||
||[varianceP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variancep-member)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#excel-excel-table-legacyid-member)|Возвращает числимый ID.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrange-member(1))|Получает диапазон, который представляет измененную область таблицы на определенном таблице.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrangeornullobject-member(1))|Получает диапазон, который представляет измененную область таблицы на определенном таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#excel-excel-workbook-readonly-member)|Возвращается `true` , если книга открыта в режиме только для чтения.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member)|Возникает при расчете таблицы.|
||[showGridlines](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showgridlines-member)|Указывает, видны ли линии сетки пользователю.|
||[showHeadings](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showheadings-member)|Указывает, видны ли заголовки пользователю.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-worksheetid-member)|Получает ID таблицы, в которой произошел расчет.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrange-member(1))|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrangeornullobject-member(1))|Получает диапазон, представляющий измененную область конкретного листа.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member)|Возникает при расчете любого таблицы в книге.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

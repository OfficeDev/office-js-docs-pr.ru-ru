---
title: Excel Набор API JavaScript 1.9
description: Сведения о наборе требований ExcelApi 1.9.
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: eb917ed75049f965178075f57e8d0e9e7630bc9081019763e7812b221a00f67c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098293"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Новые возможности в Excel API JavaScript 1.9

С набором обязательных элементов 1.9 добавлено более 500 новых API Excel. В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Фигуры](../../excel/excel-add-ins-shapes.md) | Вставка, размещение и форматирование изображений, геометрических фигур и текстовых полей. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Автофильтр](../../excel/excel-add-ins-worksheets.md#filter-data) | Добавление фильтров к диапазонам. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Области](../../excel/excel-add-ins-multiple-ranges.md) | Поддержка несплошных диапазонов. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Специальные ячейки](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Получение ячеек, содержащих даты, примечания или формулы в диапазоне. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Поиск](../../excel/excel-add-ins-ranges-string-match.md) | Поиск значений или формул в диапазоне или листе. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Копирование и вставка](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Копирование значений, форматов и формул из одного диапазона в другой. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Вычисление](../../excel/performance.md#suspend-calculation-temporarily) | Улучшенное управление модулем вычислений Excel. | [Application](/javascript/api/excel/excel.application) |
| Новые диаграммы | Познакомьтесь с новыми поддерживаемыми типами диаграмм: с картами, ящик с усами, каскадная, солнечные лучи, диаграмма Парето и воронка. | [Chart](/javascript/api/excel/excel.charttype) |
| Формат диапазона | Новые возможности для форматирования диапазонов. | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.9. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.9 или более ранних, см. в Excel API в наборе требований [1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)или более ранних .

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationEngineVersion)|Возвращает версию модуля вычислений Excel, использованного для последнего полного пересчета.|
||[calculationState](/javascript/api/excel/excel.application#calculationState)|Возвращает состояние вычисления приложения.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativeCalculation)|Возвращает параметры итеративных вычислений.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendScreenUpdatingUntilNextSync__)|Приостанавливать обновление экрана до `context.sync()` следующего.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply_range__columnIndex__criteria_)|Применяет автофильтр к диапазону.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearCriteria__)|Очищает условия фильтрации автофильтра.|
||[getRange()](/javascript/api/excel/excel.autofilter#getRange__)|Возвращает объект, который представляет диапазон, к которому `Range` применяется AutoFilter.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getRangeOrNullObject__)|Возвращает объект, который представляет диапазон, к которому `Range` применяется AutoFilter.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Массив, содержащий все условия фильтрации в диапазоне с примененным автофильтром.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Указывает, включен ли autoFilter.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isDataFiltered)|Указывает, есть ли у autoFilter критерии фильтрации.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply__)|Применяет указанный объект Autofilter, находящийся в настоящее время в диапазоне.|
||[remove()](/javascript/api/excel/excel.autofilter#remove__)|Удаляет автофильтр из диапазона.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Представляет свойство `color` одинарной границы.|
||[style](/javascript/api/excel/excel.cellborder#style)|Представляет свойство `style` одинарной границы.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintAndShade)|Представляет свойство `tintAndShade` одинарной границы.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Представляет свойство `weight` одинарной границы.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|Представляет свойство `format.borders.bottom`.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonalDown)|Представляет свойство `format.borders.diagonalDown`.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalUp)|Представляет свойство `format.borders.diagonalUp`.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Представляет свойство `format.borders.horizontal`.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Представляет свойство `format.borders.left`.|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|Представляет свойство `format.borders.right`.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Представляет свойство `format.borders.top`.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Представляет свойство `format.borders.vertical`.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addressLocal)|Представляет свойство `addressLocal`.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Представляет свойство `hidden`.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Представляет свойство `format.fill.color`.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Представляет свойство `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patternColor)|Представляет свойство `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patternTintAndShade)|Представляет свойство `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintAndShade)|Представляет свойство `format.fill.tintAndShade`.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Представляет свойство `format.font.bold`.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Представляет свойство `format.font.color`.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Представляет свойство `format.font.italic`.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Представляет свойство `format.font.name`.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Представляет свойство `format.font.size`.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Представляет свойство `format.font.strikethrough`.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Представляет свойство `format.font.subscript`.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Представляет свойство `format.font.superscript`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintAndShade)|Представляет свойство `format.font.tintAndShade`.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Представляет свойство `format.font.underline`.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoIndent)|Представляет свойство `autoIndent`.|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Представляет свойство `borders`.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Представляет свойство `fill`.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|Представляет свойство `font`.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalAlignment)|Представляет свойство `horizontalAlignment`.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentLevel)|Представляет свойство `indentLevel`.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Представляет свойство `protection`.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingOrder)|Представляет свойство `readingOrder`.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinkToFit)|Представляет свойство `shrinkToFit`.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textOrientation)|Представляет свойство `textOrientation`.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)|Представляет свойство `useStandardHeight`.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth)|Представляет свойство `useStandardWidth`.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalAlignment)|Представляет свойство `verticalAlignment`.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wrapText)|Представляет свойство `wrapText`.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulaHidden)|Представляет свойство `format.protection.formulaHidden`.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Представляет свойство `format.protection.locked`.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueAfter)|Представляет значение после изменения.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valueBefore)|Представляет значение перед изменением.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valueTypeAfter)|Представляет тип значения после изменения.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valueTypeBefore)|Представляет тип значения перед изменением.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate__)|Активирует диаграмму в пользовательском интерфейсе Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotOptions)|Объединяет параметры для сводной диаграммы.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorScheme)|Указывает цветовую схему диаграммы.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedCorners)|Указывает, имеет ли область диаграммы закругленные углы.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linkNumberFormat)|Указывает, связан ли формат номеров с ячейками.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowOverflow)|Указывает, включен ли переполнение бина в диаграмме гистограммы или диаграмме pareto.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowUnderflow)|Указывает, включен ли недополуч бин в диаграмме гистограммы или диаграмме pareto.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Указывает количество бинов диаграммы гистограммы или диаграммы pareto.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowValue)|Указывает значение переполнения ячейки диаграммы гистограммы или диаграммы pareto.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Указывает тип бина для диаграммы гистограммы или диаграммы pareto.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowValue)|Указывает значение недополука бина для диаграммы гистограммы или диаграммы pareto.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Указывает значение ширины ячейки диаграммы гистограммы или диаграммы pareto.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartileCalculation)|Указывает, указывается ли тип квартильного вычисления диаграммы полей и усов.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showInnerPoints)|Указывает, показаны ли внутренние точки в поле и диаграмме усов.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanLine)|Указывает, отображается ли в поле и диаграмме усов значимая строка.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanMarker)|Указывает, отображается ли маркер в поле и диаграмме усов.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showOutlierPoints)|Указывает, показаны ли точки выброса в поле и диаграмме усов.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linkNumberFormat)|Указывает, связан ли формат номеров с ячейками (чтобы формат номеров менял метки при изменениях в ячейках).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linkNumberFormat)|Указывает, связан ли формат номеров с ячейками.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endStyleCap)|Указывает, есть ли у баров ошибок крышка конца стиля.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Указывает, какие части планок погрешностей нужно включить.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Указывает тип форматирования планок погрешностей.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|Тип диапазона, помеченного планками погрешностей.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Указывает, отображаются ли бары ошибок.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Представляет форматирование линий диаграммы.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelStrategy)|Указывает стратегию меток на карте серии на диаграмме карты региона.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Указывает уровень сопоставления ряда диаграммы карты региона.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectionType)|Указывает тип проекции серии диаграммы карты региона.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showAxisFieldButtons)|Указывает, следует ли отображать кнопки поля оси на сводная диаграмма.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showLegendFieldButtons)|Указывает, следует ли отображать кнопки поля легенды на сводная диаграмма.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showReportFilterFieldButtons)|Указывает, следует ли отображать кнопки поля фильтрации отчетов на сводная диаграмма.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showValueFieldButtons)|Указывает, следует ли отображать кнопки поля отображения значения на сводная диаграмма.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubbleScale)|Может быть целым числом от 0 (нуля) до 300, представляющим процентное значение от размера по умолчанию.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientMaximumColor)|Указывает цвет для максимального значения серии диаграммы карты региона.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientMaximumType)|Указывает тип для максимального значения серии диаграммы карты региона.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientMaximumValue)|Указывает максимальное значение серии диаграммы карты региона.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientMidpointColor)|Указывает цвет для значения средней точки серии диаграммы карты региона.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientMidpointType)|Указывает тип для значения средней точки серии диаграммы карты региона.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientMidpointValue)|Указывает значение средней точки серии диаграммы карты региона.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientMinimumColor)|Указывает цвет для минимального значения серии диаграммы карты региона.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientMinimumType)|Указывает тип для минимального значения серии диаграммы карты региона.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientMinimumValue)|Указывает минимальное значение серии диаграммы карты региона.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientStyle)|Указывает стиль градиента серии диаграммы карты региона.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertColor)|Указывает цвет заполнения для отрицательных точек данных в серии.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentLabelStrategy)|Указывает область стратегии родительской метки серии для диаграммы treemap.|
||[binOptions](/javascript/api/excel/excel.chartseries#binOptions)|Объединяет параметры интервалов для гистограмм и диаграмм Парето.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskerOptions)|Объединяет параметры для диаграмм "ящик с усами"|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapOptions)|Объединяет параметры для диаграммы с картой региона.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xErrorBars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yErrorBars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showConnectorLines)|Указывает, показаны ли линии соединители в диаграммах водопада.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showLeaderLines)|Указывает, отображаются ли строки лидеров для каждой метки данных в серии.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitValue)|Указывает пороговое значение, которое разделяет два раздела диаграммы пирога или диаграммы "окантовка пирога".|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linkNumberFormat)|Указывает, связан ли формат номеров с ячейками (чтобы формат номеров менял метки при изменениях в ячейках).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addressLocal)|Представляет свойство `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnIndex)|Представляет свойство `columnIndex`.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getRanges__)|Возвращает один или несколько прямоугольных диапазонов, к которым применяется `RangeAreas` кондитональный формат.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getInvalidCells__)|Возвращает объект, состоящий из одного или нескольких прямоугольных `RangeAreas` диапазонов, с недействительными значениями ячейки.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getInvalidCellsOrNullObject__)|Возвращает объект, состоящий из одного или нескольких прямоугольных `RangeAreas` диапазонов, с недействительными значениями ячейки.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subField)|Свойство, используемее фильтром для фильтрации богатых значений.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Возвращает идентификатор фигуры.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Возвращает объект `Shape` для геометрической фигуры.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getCount__)|Возвращает количество фигур в группе фигур.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getItem_key_)|Получает фигуру с ее именем или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getItemAt_index_)|Получает фигуру на основе ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerFooter)|В центре таблицы.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerHeader)|Заглавный заглавный центр таблицы.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftFooter)|Левый футер таблицы.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftHeader)|Левый заготок таблицы.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightFooter)|Правый ступник таблицы.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightHeader)|Правый заготок таблицы.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultForAllPages)|Общий колонтитул, используемый для всех страниц, если не указан колонтитул четных и нечетных страниц или первой страницы.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenPages)|Колонтитул для четных страниц, для нечетных страниц нужно указывать отдельный колонтитул.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstPage)|Колонтитул первой страницы, для остальных страниц используется общий или четный и нечетный колонтитулы.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddPages)|Колонтитул для нечетных страниц, для четных страниц нужно указывать отдельный колонтитул.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Состояние, в котором задаются заглавные и пешеходные дорожки.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#useSheetMargins)|Получает или задает отметку, которая указывает, выровнены ли колонтитулы относительно полей страницы, установленных в параметрах макета страницы для листа.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#useSheetScale)|Получает или задает отметку, которая указывает, нужно ли масштабировать колонтитулы с помощью процентных значений, установленных в параметрах макета страницы для листа.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Возвращает формат изображения.|
||[id](/javascript/api/excel/excel.image#id)|Указывает идентификатор формы для объекта изображения.|
||[shape](/javascript/api/excel/excel.image#shape)|Возвращает `Shape` объект, связанный с изображением.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Значение true, если в Excel используется итерация для разрешения циклических ссылок.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxChange)|Указывает максимальное количество изменений между каждой итерацией, Excel устраняет круговые ссылки.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxIteration)|Указывает максимальное количество итераций, Excel можно использовать для решения круговой ссылки.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginArrowheadLength)|Представляет длину наконечника в начале указанной линии.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginArrowheadStyle)|Представляет стиль наконечника в начале указанной линии.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginArrowheadWidth)|Представляет ширину наконечника в начале указанной линии.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectBeginShape_shape__connectionSite_)|Привязывает начало указанного соединителя к указанной фигуре.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectEndShape_shape__connectionSite_)|Привязывает конец указанного соединителя к указанной фигуре.|
||[connectorType](/javascript/api/excel/excel.line#connectorType)|Представляет тип соединительной линии.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectBeginShape__)|Отвязывает начало указанного соединителя от фигуры.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectEndShape__)|Отвязывает конец указанного соединителя от фигуры.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endArrowheadLength)|Представляет длину наконечника в конце указанной линии.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endArrowheadStyle)|Представляет стиль наконечника в конце указанной линии.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endArrowheadWidth)|Представляет ширину наконечника в конце указанной линии.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginConnectedShape)|Представляет фигуру, к которой привязано начало указанной линии.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginConnectedSite)|Представляет точку соединения, к которой привязано начало соединительной линии.|
||[endConnectedShape](/javascript/api/excel/excel.line#endConnectedShape)|Представляет фигуру, к которой привязан конец указанной линии.|
||[endConnectedSite](/javascript/api/excel/excel.line#endConnectedSite)|Представляет точку соединения, к которой привязан конец соединительной линии.|
||[id](/javascript/api/excel/excel.line#id)|Указывает идентификатор формы.|
||[isBeginConnected](/javascript/api/excel/excel.line#isBeginConnected)|Указывает, подключено ли начало указанной строки к фигуре.|
||[isEndConnected](/javascript/api/excel/excel.line#isEndConnected)|Указывает, подключен ли конец указанной строки к фигуре.|
||[shape](/javascript/api/excel/excel.line#shape)|Возвращает `Shape` объект, связанный с строкой.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete__)|Удаляет объект разрыва страницы.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getCellAfterBreak__)|Получает первую ячейку после разрыва страницы.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnIndex)|Указывает индекс столбца для разрыва страницы.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowIndex)|Указывает индекс строки для разрыва страницы.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add_pageBreakRange_)|Добавляет разрыв страницы перед левой верхней ячейкой указанного диапазона.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getCount__)|Получает количество разрывов страниц в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getItem_index_)|Получает объект разрыва страницы по индексу.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removePageBreaks__)|Сбрасывает все добавленные вручную разрывы страниц в коллекции.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackAndWhite)|Параметр черной и белой печати таблицы.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottomMargin)|Поля нижней страницы таблицы, которые можно использовать для печати в точках.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerHorizontally)|Центр таблицы горизонтально флаг.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centerVertically)|Центр таблицы вертикально флаг.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftMode)|Вариант режима черновика таблицы.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstPageNumber)|Номер первой страницы таблицы для печати.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footerMargin)|Поле для подножки таблицы в точках для использования при печати.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getPrintArea__)|Получает объект, состоящий из одного или нескольких прямоугольных диапазонов, который представляет область печати `RangeAreas` для таблицы.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintAreaOrNullObject__)|Получает объект, состоящий из одного или нескольких прямоугольных диапазонов, который представляет область печати `RangeAreas` для таблицы.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumns__)|Получает объект range, представляющий столбцы заголовков.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumnsOrNullObject__)|Получает объект range, представляющий столбцы заголовков.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getPrintTitleRows__)|Получает объект range, представляющий строки заголовков.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleRowsOrNullObject__)|Получает объект range, представляющий строки заголовков.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headerMargin)|Поле заглавной таблицы в точках для использования при печати.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftMargin)|Левая маржа таблицы в точках для использования при печати.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|Ориентация таблицы страницы.|
||[paperSize](/javascript/api/excel/excel.pagelayout#paperSize)|Размер бумаги листа страницы.|
||[printComments](/javascript/api/excel/excel.pagelayout#printComments)|Указывает, должны ли при печати отображаться комментарии таблицы.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printErrors)|Параметр ошибки печати таблицы.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printGridlines)|Указывает, будут ли напечатаны сетки таблицы.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printHeadings)|Указывает, будут ли напечатаны заголовки таблицы.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printOrder)|Параметр распечатать страницы лист.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersFooters)|Настройка колонтитулов для листа.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightMargin)|Правое поле таблицы в точках для использования при печати.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setPrintArea_printArea_)|Задает область печати листа.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setPrintMargins_unit__marginOptions_)|Задает поля страницы с единицами измерения для листа.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleColumns_printTitleColumns_)|Задает столбцы, содержащие ячейки, которые должны повторяться слева на каждой странице при печати листа.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleRows_printTitleRows_)|Задает строки, содержащие ячейки, которые должны повторяться сверху каждой страницы при печати листа.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topMargin)|Верхняя маржа таблицы в точках для использования при печати.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Параметры масштабирования печати таблицы.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Указывает нижнюю маржу макета страницы в единице, указанной для печати.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Указывает поле для подножки макета страницы в единице, указанной для печати.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Указывает маржу загона макета страницы в единице, указанной для печати.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Указывает левое поле макета страницы в единице, указанной для печати.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Указывает правую маржу макета страницы в единице, указанной для печати.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Указывает верхнюю маржу макета страницы в единице, указанной для печати.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalFitToPages)|Количество страниц, размещаемых по горизонтали.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|Значение масштаба печатной страницы может быть равным от 10 до 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalFitToPages)|Количество страниц, размещаемых по вертикали.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortByValues_sortBy__valuesHierarchy__pivotItemScope_)|Сортирует сводную таблицу по указанным значениям в определенной области.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoFormat)|Указывает, будет ли форматирование автоматически отформатировано при обновлении или при перемещении полей.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getDataHierarchy_cell_)|Получает объект DataHierarchy, использующийся для вычисления значения в указанном диапазоне сводной таблицы.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getPivotItems_axis__cell_)|Получает объекты PivotItem с оси, образующие значение в указанном диапазоне сводной таблицы.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveFormatting)|Указывает, сохраняется ли форматирование при обновлении или пересчете отчета с помощью операций, таких как развязка, сортировка или изменение элементов поля страниц.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setAutoSortOnCell_cell__sortBy_)|Задает для сводной таблицы автоматическую сортировку, используя указанную ячейку, чтобы автоматически выбрать все необходимые условия и контекст.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enableDataValueEditing)|Указывает, разрешается ли пользователю изменять значения в теле данных.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#useCustomSortLists)|Указывает, использует ли pivotTable настраиваемые списки при сортировке.|
|[Range](/javascript/api/excel/excel.range)|[autoFill (destinationRange?: Range \| string, autoFillType?: Excel. AutoFillType)](/javascript/api/excel/excel.range#autoFill_destinationRange__autoFillType_)|Заполняет диапазон от текущего диапазона до диапазона назначения с помощью указанной логики AutoFill.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertDataTypeToText__)|Преобразует ячейки диапазона с типами данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#convertToLinkedDataType_serviceID__languageCulture_)|Преобразует ячейки диапазона в связанные типы данных в таблице.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|Копирует данные ячейки или форматирование из диапазона исходных данных или `RangeAreas` текущего диапазона.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find_text__criteria_)|Находит определенную строку на основе указанных условий.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findOrNullObject_text__criteria_)|Находит определенную строку на основе указанных условий.|
||[flashFill()](/javascript/api/excel/excel.range#flashFill__)|Делает флэш-заполнение для текущего диапазона.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getCellProperties_cellPropertiesLoadOptions_)|Возвращает двумерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой ячейки.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getColumnProperties_columnPropertiesLoadOptions_)|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждого столбца.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getRowProperties_rowPropertiesLoadOptions_)|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой строки.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCells_cellType__cellValueType_)|Получает объект, состоящий из одного или нескольких прямоугольных диапазонов, который представляет все ячейки, которые соответствуют указанному `RangeAreas` типу и значению.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCellsOrNullObject_cellType__cellValueType_)|Получает объект, состоящий из одного или нескольких диапазонов, который представляет все ячейки, которые соответствуют указанному `RangeAreas` типу и значению.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getTables_fullyContained_)|Получает коллекцию таблиц с заданной областью, перекрывающую диапазон.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkedDataTypeState)|Представляет состояние типа данных каждой ячейки.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_)|Удаляет повторяющиеся значения из диапазона, заданного столбцами.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceAll_text__replacement__criteria_)|Находит и заменяет определенную строку на основе условий, указанных в текущем диапазоне.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setCellProperties_cellPropertiesData_)|Обновляет диапазон на основе 2D-массива свойств ячейки, инкапсулируя такие вещи, как шрифт, заливка, границы и выравнивание.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setColumnProperties_columnPropertiesData_)|Обновляет диапазон на основе одномерного массива свойств столбцов, инкапсулируя такие вещи, как шрифт, заливка, границы и выравнивание.|
||[setDirty()](/javascript/api/excel/excel.range#setDirty__)|Устанавливает диапазон, предназначенный для пересчета при выполнении следующего пересчета.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setRowProperties_rowPropertiesData_)|Обновляет диапазон на основе одномерного массива свойств строки, инкапсулируя такие вещи, как шрифт, заливка, границы и выравнивание.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate__)|Вычисляет все ячейки `RangeAreas` в .|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear_applyTo_)|Очищает значения, формат, заполнение, границу и другие свойства в каждом из областей, в которых состоит `RangeAreas` этот объект.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertDataTypeToText__)|Преобразует все ячейки в `RangeAreas` типах данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#convertToLinkedDataType_serviceID__languageCulture_)|Преобразует все ячейки в связанные `RangeAreas` типы данных.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|Копирует данные ячейки или форматирование из диапазона исходных данных или `RangeAreas` текущего `RangeAreas` .|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getEntireColumn__)|Возвращает объект, который представляет целые столбцы (например, если ток представляет ячейки `RangeAreas` `RangeAreas` `RangeAreas` "B4:E11, H2", он возвращает столбцы `RangeAreas` "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getEntireRow__)|Возвращает объект, который представляет целые строки (например, если ток представляет ячейки `RangeAreas` `RangeAreas` `RangeAreas` "B4:E11", он возвращает строки `RangeAreas` "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersection_anotherRange_)|Возвращает `RangeAreas` объект, который представляет пересечение заданных диапазонов или `RangeAreas` .|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersectionOrNullObject_anotherRange_)|Возвращает `RangeAreas` объект, который представляет пересечение заданных диапазонов или `RangeAreas` .|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getOffsetRangeAreas_rowOffset__columnOffset_)|Возвращает `RangeAreas` объект, смещенный определенной строкой и смещением столбца.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCells_cellType__cellValueType_)|Возвращает объект, который представляет все ячейки, которые `RangeAreas` соответствуют указанному типу и значению.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCellsOrNullObject_cellType__cellValueType_)|Возвращает объект, который представляет все ячейки, которые `RangeAreas` соответствуют указанному типу и значению.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#getTables_fullyContained_)|Возвращает объемную коллекцию таблиц, которые перекрываются с любым диапазоном в этом `RangeAreas` объекте.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreas_valuesOnly_)|Возвращает используемое, которое включает все используемые области отдельных прямоугольных `RangeAreas` диапазонов `RangeAreas` объекта.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreasOrNullObject_valuesOnly_)|Возвращает используемое, которое включает все используемые области отдельных прямоугольных `RangeAreas` диапазонов `RangeAreas` объекта.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Возвращает `RangeAreas` ссылку в стиле A1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addressLocal)|Возвращает `RangeAreas` ссылку в локале пользователя.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areaCount)|Возвращает количество прямоугольных диапазонов, составляющих этот `RangeAreas` объект.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Возвращает коллекцию прямоугольных диапазонов, которые составляют этот `RangeAreas` объект.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellCount)|Возвращает количество ячеек в объекте, суммирует количество ячеек всех отдельных `RangeAreas` прямоугольных диапазонов.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalFormats)|Возвращает коллекцию условных форматов, которые пересекаются с любыми ячейками в этом `RangeAreas` объекте.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#dataValidation)|Возвращает объект проверки данных для всех диапазонов в `RangeAreas` .|
||[format](/javascript/api/excel/excel.rangeareas#format)|Возвращает объект, инкапсулируя шрифт, заполнять, границы, выравнивание и другие свойства для всех `RangeFormat` диапазонов `RangeAreas` объекта.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isEntireColumn)|Указывает, представляют ли все диапазоны на этом объекте целые столбцы `RangeAreas` (например, "A:C, Q:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isEntireRow)|Указывает, представляют ли все диапазоны на этом объекте целые строки `RangeAreas` (например, "1:3, 5:7").|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Возвращает таблицу для текущего `RangeAreas` .|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setDirty__)|Задает перерасчет при следующем `RangeAreas` пересчете.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Представляет стиль для всех диапазонов в этом `RangeAreas` объекте.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintAndShade)|Указывает двойной, который осветляет или темнеет цвет для границы диапазона, значение между -1 (самый темный) и 1 (самый яркий), с 0 для исходного цвета.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintAndShade)|Указывает двойник, который осветляет или темнеет цвет для границ диапазона.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getCount__)|Возвращает количество диапазонов в `RangeCollection` .|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getItemAt_index_)|Возвращает объект диапазона в зависимости от его положения в `RangeCollection` .|
||[items](/javascript/api/excel/excel.rangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Шаблон диапазона.|
||[patternColor](/javascript/api/excel/excel.rangefill#patternColor)|Цветовой код HTML, представляющий цвет шаблона диапазона, в форме #RRGGBB (например, "FFA500"), или в виде имени HTML-цвета (например, "оранжевый").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patternTintAndShade)|Указывает двойной номер, который осветляет или темнеет цвет шаблона для заполнения диапазона.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintAndShade)|Указывает двойной, который осветляет или затемнеет цвет для заполнения диапазона.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Указывает состояние забастовки шрифта.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Указывает состояние подписки шрифта.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Указывает состояние суперсписка шрифта.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintAndShade)|Указывает двойную, которая осветляет или темнеет цвет шрифта диапазона.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoIndent)|Указывает, будет ли текст автоматически отступным, если выравнивание текста задано для равного распространения.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentLevel)|Целое число от 0 до 250, указывающее уровень отступа.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingOrder)|Направление чтения для диапазона.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinkToFit)|Указывает, если текст автоматически сокращается, чтобы соответствовать ширине доступных столбцов.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Количество повторяющихся строк, удаленных операцией.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueRemaining)|Количество оставшихся уникальных строк, присутствующих в получившемся диапазоне.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completeMatch)|Указывает, должен ли совпадение быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchCase)|Указывает, является ли совпадение чувствительным к делу.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addressLocal)|Представляет свойство `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowIndex)|Представляет свойство `rowIndex`.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completeMatch)|Указывает, должен ли совпадение быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchCase)|Указывает, является ли совпадение чувствительным к делу.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchDirection)|Указывает направление поиска.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Представляет свойство `format`.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Представляет свойство `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Представляет свойство `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnHidden)|Представляет свойство `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnWidth)||
||[формат: Excel. CellPropertiesFormat & {
            columnWidth?] (/javascript/api/excel/excel.settablecolumnproperties#format)|Представляет свойство `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[формат: Excel. CellPropertiesFormat & {
            rowHeight?] (/javascript/api/excel/excel.settablerowproperties#format)|Представляет свойство `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowHeight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowHidden)|Представляет свойство `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#altTextDescription)|Указывает альтернативный текст описания `Shape` объекта.|
||[altTextTitle](/javascript/api/excel/excel.shape#altTextTitle)|Указывает альтернативный текст заголовка для `Shape` объекта.|
||[delete()](/javascript/api/excel/excel.shape#delete__)|Удаляет фигуру с листа.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricShapeType)|Указывает тип геометрической фигуры этой геометрической фигуры.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getAsImage_format_)|Преобразует фигуру в изображение и возвращает изображение в виде строки в кодировке base64.|
||[height](/javascript/api/excel/excel.shape#height)|Указывает высоту фигуры в точках.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementLeft_increment_)|Перемещает фигуру по горизонтали на указанное число пунктов.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementRotation_increment_)|Поворачивает фигуру по часовой стрелке относительно оси Z на указанное число градусов.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementTop_increment_)|Перемещает фигуру по вертикали на указанное число пунктов.|
||[left](/javascript/api/excel/excel.shape#left)|Расстояние в пунктах от левого края фигуры до левого края листа.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockAspectRatio)|Указывает, заблокировано ли соотношение аспектов этой фигуры.|
||[name](/javascript/api/excel/excel.shape#name)|Указывает имя фигуры.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionSiteCount)|Возвращает количество точек соединения на фигуре.|
||[fill](/javascript/api/excel/excel.shape#fill)|Возвращает формат заливки фигуры.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricShape)|Возвращает геометрическую фигуру, связанную с линией.|
||[group](/javascript/api/excel/excel.shape#group)|Возвращает группу фигур, связанную с фигурой.|
||[id](/javascript/api/excel/excel.shape#id)|Указывает идентификатор формы.|
||[image](/javascript/api/excel/excel.shape#image)|Возвращает изображение, связанное с фигурой.|
||[level](/javascript/api/excel/excel.shape#level)|Указывает уровень указанной формы.|
||[line](/javascript/api/excel/excel.shape#line)|Возвращает линию, связанную с фигурой.|
||[lineFormat](/javascript/api/excel/excel.shape#lineFormat)|Возвращает формат линии для фигуры.|
||[onActivated](/javascript/api/excel/excel.shape#onActivated)|Возникает, если фигура активирована.|
||[onDeactivated](/javascript/api/excel/excel.shape#onDeactivated)|Возникает, если фигура деактивирована.|
||[parentGroup](/javascript/api/excel/excel.shape#parentGroup)|Указывает родительную группу этой фигуры.|
||[textFrame](/javascript/api/excel/excel.shape#textFrame)|Возвращает объект рамки с текстом для фигуры.|
||[type](/javascript/api/excel/excel.shape#type)|Возвращает тип фигуры.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zOrderPosition)|Возвращает положение указанной фигуры по оси Z. Значение 0 представляет нижнее положение по оси.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Указывает вращение фигуры в градусах.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleHeight_scaleFactor__scaleType__scaleFrom_)|Масштабирует высоту фигуры с применением указанного коэффициента.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleWidth_scaleFactor__scaleType__scaleFrom_)|Масштабирует ширину фигуры с применением указанного коэффициента.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setZOrder_position_)|Перемещает указанную фигуру вверх или вниз по оси Z в коллекции, что переносит ее вперед или назад относительно других фигур.|
||[top](/javascript/api/excel/excel.shape#top)|Расстояние в пунктах от верхнего края фигуры до верхнего края листа.|
||[visible](/javascript/api/excel/excel.shape#visible)|Указывает, видна ли фигура.|
||[width](/javascript/api/excel/excel.shape#width)|Указывает ширину в точках формы.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeId)|Получает ID активированной фигуры.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetId)|Получает ID таблицы, в которой активируется фигура.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addGeometricShape_geometricShapeType_)|Добавляет геометрическую фигуру на лист.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addGroup_values_)|Группирует подмножество фигур на листе этой коллекции.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addImage_base64ImageString_)|Создает изображение из строки в кодировке base64 и добавляет его на лист.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addLine_startLeft__startTop__endLeft__endTop__connectorType_)|Добавляет линию на лист.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addTextBox_text_)|Добавляет текстовое поле на лист с указанным текстом в качестве содержимого.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getCount__)|Возвращает количество фигур на листе.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getItem_key_)|Получает фигуру с ее именем или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getItemAt_index_)|Получает фигуру с помощью ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeId)|Получает ID деактивированной фигуры.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetId)|Получает ID таблицы, в которой фигура деактивирована.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear__)|Очищает формат заливки фигуры.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundColor)|Представляет цвет переднего плана заполнения фигуры в формате HTML-цвета в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый")|
||[type](/javascript/api/excel/excel.shapefill#type)|Возвращает тип заливки фигуры.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setSolidColor_color_)|Задает заливку одним цветом для фигуры.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Указывает процент прозрачности заполнения как значение от 0.0 (непрозрачная) до 1.0 (clear).|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.shapefont#color)|Представление цветового кода HTML текстового цвета (например, "#FF0000" представляет красный цвет).|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.shapefont#name)|Представляет имя шрифта (например, "Калибри").|
||[size](/javascript/api/excel/excel.shapefont#size)|Представляет размер шрифта в точках (например, 11).|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Тип подчеркивания, применяемый для шрифта.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Указывает идентификатор формы.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Возвращает `Shape` объект, связанный с группой.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Возвращает коллекцию `Shape` объектов.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup__)|Отменяет группировку любых сгруппированных фигур в указанной группе фигур.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Представляет цвет строки в формате HTML-цвета в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashStyle)|Представляет тип линии фигуры.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Представляет тип линии фигуры.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Представляет степень прозрачности указанной линии как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Указывает, отображается ли форматирование строки элемента фигуры.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Представляет толщину линии (в пунктах).|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subField)|Указывает подполе, которое является целевым именем свойства для сортировки с богатым значением.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getCount__)|Получает количество стилей в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getItemAt_index_)|Получает стиль на основе его позиции в коллекции.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autoFilter)|Представляет объект `AutoFilter` таблицы.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableId)|Получает ID добавленной таблицы.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetId)|Получает ID таблицы, в которую добавляется таблица.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|Получает сведения о деталях изменений.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onAdded)|Возникает при добавлении новой таблицы в книгу.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#onDeleted)|Возникает, если указанная таблица удалена из книги.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableId)|Получает удаленный ID таблицы.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tableName)|Получает имя удаляемой таблицы.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetId)|Получает ID таблицы, в которой удаляется таблица.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getCount__)|Получает количество таблиц в коллекции.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getFirst__)|Получает первую таблицу в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItem_key_)|Получает таблицу по имени или ИД.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autoSizeSetting)|Автоматические параметры размеров для текстового кадра.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottomMargin)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[deleteText()](/javascript/api/excel/excel.textframe#deleteText__)|Удаляет весь текст в рамке с текстом.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalAlignment)|Представляет горизонтальное выравнивание рамки с текстом.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontalOverflow)|Представляет действие горизонтального переполнения рамки с текстом.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftMargin)|Представляет левое поле рамки с текстом (в пунктах).|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Представляет угол, на который ориентирован текст для текстового кадра.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingOrder)|Представляет направление чтения рамки с текстом (слева направо или справа налево).|
||[hasText](/javascript/api/excel/excel.textframe#hasText)|Указывает, содержит ли текстовый кадр текст.|
||[textRange](/javascript/api/excel/excel.textframe#textRange)|Представляет текст, присоединенный к фигуре в текстовой рамке, а также свойства и методы для операций с текстом.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightMargin)|Представляет правое поле рамки с текстом (в пунктах).|
||[topMargin](/javascript/api/excel/excel.textframe#topMargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalAlignment)|Представляет вертикальное выравнивание для рамки с текстом.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticalOverflow)|Представляет действие вертикального переполнения рамки с текстом.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getSubstring_start__length_)|Возвращает объект TextRange для подстроки в указанном диапазоне.|
||[font](/javascript/api/excel/excel.textrange#font)|Возвращает `ShapeFont` объект, который представляет атрибуты шрифта для диапазона текста.|
||[text](/javascript/api/excel/excel.textrange#text)|Представляет содержимое с обычным текстом в диапазоне текста.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartDataPointTrack)|Значение true, если все диаграммы в книге отслеживают точки фактических данных, с которыми они связаны.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getActiveChart__)|Получает текущую активную диаграмму в книге.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getActiveChartOrNullObject__)|Получает текущую активную диаграмму в книге.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getIsActiveCollabSession__)|`true`Возвращается, если книга редактирована несколькими пользователями (с помощью соавторов).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getSelectedRanges__)|Получает текущий выделенный диапазон (один или несколько) в книге.|
||[isDirty](/javascript/api/excel/excel.workbook#isDirty)|Указывает, были ли внесены изменения с момента последнего сберегаемого книги.|
||[autoSave](/javascript/api/excel/excel.workbook#autoSave)|Указывает, находится ли книга в режиме AutoSave.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationEngineVersion)|Возвращает номер версии модуля вычислений Excel.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged)|Возникает при смене параметра AutoSave в книге.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslySaved)|Указывает, была ли книга сохранена локально или в Интернете.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#usePrecisionAsDisplayed)|Значение true, если вычисления в книге выполняются только с той точностью чисел, с которой они отображаются.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Получает тип события.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enableCalculation)|Определяет, следует ли Excel при необходимости пересчитать таблицу.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAll_text__criteria_)|Находит все вхождения данной строки на основе указанных критериев и возвращает их как объект, состоящий из одного или `RangeAreas` нескольких прямоугольных диапазонов.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAllOrNullObject_text__criteria_)|Находит все вхождения данной строки на основе указанных критериев и возвращает их как объект, состоящий из одного или `RangeAreas` нескольких прямоугольных диапазонов.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getRanges_address_)|Получает объект, представляющий один или несколько блоков прямоугольных диапазонов, указанных `RangeAreas` по адресу или имени.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autoFilter)|Представляет объект `AutoFilter` таблицы.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalPageBreaks)|Получает коллекцию горизонтальных разрывов страницы для листа.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onFormatChanged)|Возникает, если изменен формат указанного листа.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pageLayout)|Получает `PageLayout` объект таблицы.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Возвращает коллекцию всех объектов Shape на листе.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalPageBreaks)|Получает коллекцию вертикальных разрывов страницы для листа.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceAll_text__replacement__criteria_)|Находит и заменяет определенную строку на основе условий, указанных в текущем листе.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Представляет сведения об изменениях.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onChanged)|Возникает при изменении любого листа в книге.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onFormatChanged)|Возникает при смене формата любого таблицы в книге.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged)|Возникает при изменениях выделения на любом листе.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRange_ctx_)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRangeOrNullObject_ctx_)|Получает диапазон, представляющий измененную область конкретного листа.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetId)|Получает ID таблицы, в которой изменились данные.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completeMatch)|Указывает, должен ли совпадение быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchCase)|Указывает, является ли совпадение чувствительным к делу.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

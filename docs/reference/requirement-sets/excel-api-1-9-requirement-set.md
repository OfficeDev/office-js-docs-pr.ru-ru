---
title: Набор обязательных элементов API JavaScript для Excel 1,9
description: Сведения о наборе требований ExcelApi 1,9.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e3878954bca943e1895a44ea9482f1c67cba9211
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996510"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Новые возможности API JavaScript для Excel 1,9

С набором обязательных элементов 1.9 добавлено более 500 новых API Excel. В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Фигуры](../../excel/excel-add-ins-shapes.md) | Вставка, размещение и форматирование изображений, геометрических фигур и текстовых полей. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Автофильтр](../../excel/excel-add-ins-worksheets.md#filter-data) | Добавление фильтров к диапазонам. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Области](../../excel/excel-add-ins-multiple-ranges.md) | Поддержка несплошных диапазонов. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Специальные ячейки](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Получение ячеек, содержащих даты, примечания или формулы в диапазоне. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Поиск](../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | Поиск значений или формул в диапазоне или листе. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Копирование и вставка](../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | Копирование значений, форматов и формул из одного диапазона в другой. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Вычисление](../../excel/performance.md#suspend-calculation-temporarily) | Улучшенное управление модулем вычислений Excel. | [Application](/javascript/api/excel/excel.application) |
| Новые диаграммы | Познакомьтесь с новыми поддерживаемыми типами диаграмм: с картами, ящик с усами, каскадная, солнечные лучи, диаграмма Парето и воронка. | [Chart](/javascript/api/excel/excel.charttype) |
| Формат диапазона | Новые возможности для форматирования диапазонов. | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,9. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,9 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Возвращает версию модуля вычислений Excel, использованного для последнего полного пересчета.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Возвращает состояние вычисления приложения.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Возвращает параметры итеративных вычислений.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Приостанавливает обновление экрана до тех пор, пока `context.sync()` не будет вызван следующий.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Применяет автофильтр к диапазону.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Очищает условия фильтрации автофильтра.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Возвращает объект Range, представляющий диапазон, к которому применяется автофильтр.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Возвращает объект Range, представляющий диапазон, к которому применяется автофильтр.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Массив, содержащий все условия фильтрации в диапазоне с примененным автофильтром.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Указывает, включен ли автофильтр.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Указывает, есть ли условия фильтра для автофильтра.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Применяет указанный объект Autofilter, находящийся в настоящее время в диапазоне.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Удаляет автофильтр из диапазона.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Представляет свойство `color` одинарной границы.|
||[style](/javascript/api/excel/excel.cellborder#style)|Представляет свойство `style` одинарной границы.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|Представляет свойство `tintAndShade` одинарной границы.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Представляет свойство `weight` одинарной границы.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|Представляет свойство `format.borders.bottom`.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|Представляет свойство `format.borders.diagonalDown`.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|Представляет свойство `format.borders.diagonalUp`.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Представляет свойство `format.borders.horizontal`.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Представляет свойство `format.borders.left`.|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|Представляет свойство `format.borders.right`.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Представляет свойство `format.borders.top`.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Представляет свойство `format.borders.vertical`.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|Представляет свойство `addressLocal`.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Представляет свойство `hidden`.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Представляет свойство `format.fill.color`.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Представляет свойство `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Представляет свойство `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Представляет свойство `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Представляет свойство `format.fill.tintAndShade`.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Представляет свойство `format.font.bold`.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Представляет свойство `format.font.color`.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Представляет свойство `format.font.italic`.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Представляет свойство `format.font.name`.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Представляет свойство `format.font.size`.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Представляет свойство `format.font.strikethrough`.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Представляет свойство `format.font.subscript`.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Представляет свойство `format.font.superscript`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|Представляет свойство `format.font.tintAndShade`.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Представляет свойство `format.font.underline`.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|Представляет свойство `autoIndent`.|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Представляет свойство `borders`.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Представляет свойство `fill`.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|Представляет свойство `font`.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|Представляет свойство `horizontalAlignment`.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|Представляет свойство `indentLevel`.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Представляет свойство `protection`.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|Представляет свойство `readingOrder`.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|Представляет свойство `shrinkToFit`.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|Представляет свойство `textOrientation`.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|Представляет свойство `useStandardHeight`.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|Представляет свойство `useStandardWidth`.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|Представляет свойство `verticalAlignment`.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Представляет свойство `wrapText`.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|Представляет свойство `format.protection.formulaHidden`.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Представляет свойство `format.protection.locked`.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Представляет значение после изменения.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Представляет значение до изменения.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Представляет тип значения после изменения|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Представляет тип значения до изменения|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Активирует диаграмму в пользовательском интерфейсе Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Объединяет параметры для сводной диаграммы.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Задает цветовую схему диаграммы.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|Указывает, содержит ли область диаграммы скругленные углы.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Указывает, связан ли числовой формат с ячейками.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Указывает, включено ли переполнение ячеек в гистограмме или диаграмме Парето.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Указывает, включена ли потеря ячеек в гистограмме или диаграмме Парето.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Указывает количество ячеек гистограммы или диаграммы Парето.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Указывает значение переполнения ячейки гистограммы или диаграммы Парето.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Указывает тип ячейки для гистограммы или диаграммы Парето.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Задает значение недотекания в ячейке гистограммы или диаграммы Парето.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Задает значение ширины ячейки гистограммы или диаграммы Парето.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Указывает, является ли тип расчета квартиль для поля и диаграммы вхискер.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Указывает, отображаются ли внутренние точки в поле и диаграмме вхискер.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Указывает, отображается ли средняя строка в поле и диаграмме вхискер.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Указывает, отображается ли индикатор среднего значения в поле и диаграмме вхискер.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Указывает, отображаются ли точки выбросов в поле и диаграмме вхискер.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Указывает, связан ли числовой формат с ячейками (поэтому формат чисел изменяется в метках при изменении в ячейках).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Указывает, связан ли числовой формат с ячейками.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Указывает, есть ли в планах погрешностей стиль End Cap.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Указывает, какие части планок погрешностей нужно включить.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Указывает тип форматирования планок погрешностей.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|Тип диапазона, помеченного планками погрешностей.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Указывает, отображаются ли планки погрешностей.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Представляет форматирование линий диаграммы.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Задает стратегию меток карты рядов для диаграммы с картой областей.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Задает уровень сопоставления ряда в диаграмме с картой областей.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Задает тип проекции ряда для диаграммы с областью отображения.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Указывает, следует ли отображать кнопки поля оси в сводной диаграмме.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Указывает, следует ли отображать кнопки поля условных обозначений в сводной диаграмме.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Указывает, следует ли отображать кнопки поля фильтра отчета в сводной диаграмме.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Указывает, следует ли отображать кнопки поля Показать значение в сводной диаграмме.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Может быть целым числом от 0 (нуля) до 300, представляющим процентное значение от размера по умолчанию.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Задает цвет для максимального значения ряда диаграммы карты областей.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Указывает тип максимального значения ряда диаграммы карты областей.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Указывает максимальное значение ряда диаграмм с картой областей.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Указывает цвет для среднего значения ряда диаграммы карты областей.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Указывает тип среднего значения для ряда диаграммы карты областей.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Указывает среднее значение ряда диаграммы карты областей.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Задает цвет для минимального значения ряда диаграммы карты областей.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Указывает тип для минимального значения ряда диаграммы карты областей.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Указывает минимальное значение ряда диаграммы карты областей.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Задает стиль градиента ряда в диаграмме с картой областей.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Задает цвет заливки для отрицательных точек данных в ряду.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Задает область схемы родительской метки ряда для диаграммы эта.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Объединяет параметры интервалов для гистограмм и диаграмм Парето.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Объединяет параметры для диаграмм "ящик с усами"|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Объединяет параметры для диаграммы с картой региона.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Указывает, отображаются ли соединительные линии в каскадных диаграммах.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Указывает, отображаются ли линии выноски для каждой подписи данных в ряду.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Задает пороговое значение, которое разделяет два раздела круговой диаграммы или гистограммы.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Указывает, связан ли числовой формат с ячейками (поэтому формат чисел изменяется в метках при изменении в ячейках).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Представляет свойство `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Представляет свойство `columnIndex`.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, к которым применено условное форматирование.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, с недопустимыми значениями ячеек.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, с недопустимыми значениями ячеек.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|Свойство, используемое фильтром для расширенной фильтрации по объектам richvalue.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Возвращает идентификатор фигуры.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Возвращает объект Shape для геометрической фигуры.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Возвращает количество фигур в группе фигур.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Получает фигуру по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Получает фигуру на основе ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Центральный нижний колонтитул листа.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Центральный верхний колонтитул рабочего листа.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Левый нижний колонтитул листа.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Левый верхний колонтитул листа.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Правый нижний колонтитул листа.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Правый верхний колонтитул листа.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|Общий колонтитул, используемый для всех страниц, если не указан колонтитул четных и нечетных страниц или первой страницы.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|Колонтитул для четных страниц, для нечетных страниц нужно указывать отдельный колонтитул.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|Колонтитул первой страницы, для остальных страниц используется общий или четный и нечетный колонтитулы.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|Колонтитул для нечетных страниц, для четных страниц нужно указывать отдельный колонтитул.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Состояние, в котором задаются верхние и нижние колонтитулы.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Получает или задает отметку, которая указывает, выровнены ли колонтитулы относительно полей страницы, установленных в параметрах макета страницы для листа.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Получает или задает отметку, которая указывает, нужно ли масштабировать колонтитулы с помощью процентных значений, установленных в параметрах макета страницы для листа.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Возвращает формат изображения.|
||[id](/javascript/api/excel/excel.image#id)|Задает идентификатор фигуры для объекта Image.|
||[shape](/javascript/api/excel/excel.image#shape)|Возвращает объект Shape, связанный с изображением.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Значение true, если в Excel используется итерация для разрешения циклических ссылок.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Задает максимальный объем изменений между итерациями по мере того, как Excel разрешает циклические ссылки.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Указывает максимальное количество итераций, которые Excel может использовать для устранения циклической ссылки.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|Представляет длину наконечника в начале указанной линии.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|Представляет стиль наконечника в начале указанной линии.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|Представляет ширину наконечника в начале указанной линии.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|Привязывает начало указанного соединителя к указанной фигуре.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|Привязывает конец указанного соединителя к указанной фигуре.|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|Представляет тип соединительной линии.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|Отвязывает начало указанного соединителя от фигуры.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|Отвязывает конец указанного соединителя от фигуры.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|Представляет длину наконечника в конце указанной линии.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|Представляет стиль наконечника в конце указанной линии.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|Представляет ширину наконечника в конце указанной линии.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Представляет фигуру, к которой привязано начало указанной линии.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Представляет точку соединения, к которой привязано начало соединительной линии.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Представляет фигуру, к которой привязан конец указанной линии.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Представляет точку соединения, к которой привязан конец соединительной линии.|
||[id](/javascript/api/excel/excel.line#id)|Задает идентификатор фигуры.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Указывает, подключается ли начало указанной линии к фигуре.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Указывает, подключена ли конец указанной линии к фигуре.|
||[shape](/javascript/api/excel/excel.line#shape)|Возвращает объект Shape, связанный с линией.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Удаляет объект разрыва страницы.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Получает первую ячейку после разрыва страницы.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Указывает индекс столбца для разрыва страницы|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Указывает индекс строки для разрыва страницы|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Добавляет разрыв страницы перед левой верхней ячейкой указанного диапазона.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Получает количество разрывов страниц в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Получает объект разрыва страницы по индексу.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Сбрасывает все добавленные вручную разрывы страниц в коллекции.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|Параметр печати "черно-белый" на листе.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|Нижнее поле страницы листа, которое будет использоваться для печати в точках.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|Флаг горизонтального центрирования листа.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|Вертикальный флаг в центре листа.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|Режим черновика листа.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|Номер первой страницы листа, который требуется напечатать.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|Поле нижнего колонтитула листа в пунктах, которое будет использоваться при печати.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Получает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, представляющих область печати для листа.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Получает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, представляющих область печати для листа.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Получает объект range, представляющий столбцы заголовков.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Получает объект range, представляющий столбцы заголовков.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Получает объект range, представляющий строки заголовков.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Получает объект range, представляющий строки заголовков.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|Поле заголовка листа (в пунктах), используемое при печати.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|Левое поле листа (в пунктах), которое будет использоваться при печати.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|Ориентация страницы на листе.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|Размер бумаги страницы на листе.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Указывает, следует ли отображать комментарии листа при печати.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|Параметр "ошибки печати листа".|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Указывает, будут ли печататься линии сетки листа.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Указывает, будут ли печататься заголовки листа.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|Параметр порядка печати на странице листа.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Настройка колонтитулов для листа.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|Правое поле листа (в пунктах), используемое при печати.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Задает область печати листа.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Задает поля страницы с единицами измерения для листа.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Задает столбцы, содержащие ячейки, которые должны повторяться слева на каждой странице при печати листа.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Задает строки, содержащие ячейки, которые должны повторяться сверху каждой страницы при печати листа.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Верхнее поле листа (в пунктах), которое будет использоваться при печати.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Параметры масштаба листа для печати.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Задает нижнее поле макета страницы в единицах, используемых для печати.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Задает поле нижнего колонтитула макета страницы в единицах, используемых для печати.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Задает поле заголовка макета страницы в единицах, используемых для печати.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Задает левое поле макета страницы в единице измерения, используемой для печати.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Задает правое поле макета страницы в единице измерения, используемой для печати.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Задает верхнее поле макета страницы в единицах, используемых для печати.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Количество страниц, размещаемых по горизонтали.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|Значение масштаба печатной страницы может быть равным от 10 до 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Количество страниц, размещаемых по вертикали.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Сортирует сводную таблицу по указанным значениям в определенной области.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Указывает, будет ли форматирование автоматически форматироваться при его обновлении или при перемещении полей.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Получает объект DataHierarchy, использующийся для вычисления значения в указанном диапазоне сводной таблицы.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Получает объекты PivotItem с оси, образующие значение в указанном диапазоне сводной таблицы.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Указывает, сохраняется ли форматирование при обновлении или пересчете отчета с помощью таких операций, как сведение, сортировка или изменение элементов поля страницы.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Задает для сводной таблицы автоматическую сортировку, используя указанную ячейку, чтобы автоматически выбрать все необходимые условия и контекст.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Указывает, позволяет ли Сводная таблица изменять значения в основном тексте данных пользователем.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Указывает, используются ли в сводной таблице настраиваемые списки при сортировке.|
|[Range](/javascript/api/excel/excel.range)|[Автозаполнение (Дестинатионранже?: \| строка диапазона, аутофиллтипе?: Excel. аутофиллтипе)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Заполняет диапазон от текущего диапазона до конечного диапазона, используя указанную логику автозаполнения.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Преобразует диапазон ячеек с типами данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Преобразует ячейки диапазона в связанный тип данных на листе.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Копирует данные ячейки или форматирование из исходного диапазона или объекта RangeAreas в текущий диапазон.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Находит определенную строку на основе указанных условий.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Находит определенную строку на основе указанных условий.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Выполняет мгновенное заполнение текущего диапазона. Функция мгновенного заполнения автоматически подставляет данные, когда обнаруживает закономерность, поэтому диапазон должен состоять из одного столбца со смежными данными, чтобы выявить закономерность.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Возвращает двумерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой ячейки.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждого столбца.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой строки.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Получает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, представляющих все ячейки, которые соответствуют указанному типу и значению.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Получает объект RangeAreas, состоящий из одного или нескольких диапазонов, представляющих все ячейки, которые соответствуют указанному типу и значению.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Получает коллекцию таблиц с заданной областью, перекрывающую диапазон.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Представляет состояние типа данных каждой ячейки.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Удаляет повторяющиеся значения из диапазона, заданного столбцами.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Находит и заменяет определенную строку на основе условий, указанных в текущем диапазоне.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Обновляет диапазон на основе двумерного массива свойств ячейки, в который включены такие элементы, как шрифт, заливка, границы, выравнивание и т. д.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Обновляет диапазон на основе одномерного массива свойств столбца, в который включены такие элементы, как шрифт, заливка, границы, выравнивание и т. д.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Устанавливает диапазон, предназначенный для пересчета при выполнении следующего пересчета.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Обновляет диапазон на основе одномерного массива свойств строки, в который включены такие элементы, как шрифт, заливка, границы, выравнивание и т. д.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Вычисляет все ячейки в объекте RangeAreas.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Удаляет значения, формат, заливку, границу и т. д. для каждой области, входящей в этот объект RangeAreas.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Преобразует все ячейки в объекте RangeAreas с типами данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Преобразует все ячейки в объекте RangeAreas в связанный тип данных.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Копирует данные ячейки или форматирование из исходного диапазона или объекта RangeAreas в текущий объект RangeAreas.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Возвращает объект RangeAreas, представляющий все столбцы объекта RangeAreas (например, если текущий объект RangeAreas представляет ячейки "B4:E11, H2", возвращается объект RangeAreas, представляющий столбцы "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Возвращает объект RangeAreas, представляющий все строки объекта RangeAreas (например, если текущий объект RangeAreas представляет ячейки "B4:E11", возвращается объект RangeAreas, представляющий строки "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Возвращает объект RangeAreas, представляющий пересечение заданных диапазонов или RangeAreas.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Возвращает объект RangeAreas, представляющий пересечение заданных диапазонов или RangeAreas.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Возвращает объект RangeAreas, смещенный на определенное количество строк и столбцов.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Возвращает объект RangeAreas, представляющий все ячейки, которые соответствуют указанному типу и значению.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Возвращает объект RangeAreas, представляющий все ячейки, которые соответствуют указанному типу и значению.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Возвращает коллекцию таблиц с заданной областью, перекрывающую любой диапазон в объекте RangeAreas.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Возвращает использованный объект RangeAreas, включающий все использованные области отдельных прямоугольных диапазонов в объекте RangeAreas.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Возвращает использованный объект RangeAreas, включающий все использованные области отдельных прямоугольных диапазонов в объекте RangeAreas.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Возвращает ссылку RangeAreas в стиле a1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Возвращает ссылку на RangeAreas в языковом стандарте пользователя.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Возвращает количество прямоугольных диапазонов, составляющих этот объект RangeAreas.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Возвращает коллекцию прямоугольных диапазонов, составляющих этот объект RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Возвращает число ячеек в объекте RangeAreas с суммированием количества ячеек всех отдельных прямоугольных диапазонов.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Возвращает коллекцию объектов ConditionalFormat, пересекающихся с любыми ячейками в этом объекте RangeAreas.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Возвращает объект dataValidation для всех диапазонов в объекте RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Возвращает объект RangeFormat, который инкапсулирует шрифт, заливку, границы, выравнивание и другие свойства для всех диапазонов в объекте RangeAreas.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Указывает, являются ли все диапазоны этого объекта RangeAreas целыми столбцами (например, "А:К, К:З").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Указывает, являются ли все диапазоны этого объекта RangeAreas целыми строками (например, "1:3, 5:7").|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Возвращает лист для текущего объекта RangeAreas.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Устанавливает объект RangeAreas, предназначенный для пересчета при выполнении следующего пересчета.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Представляет стиль всех диапазонов в этом объекте RangeAreas.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Задает двойное значение, которое осветляет или затемняет цвет границы диапазона, значение находится в диапазоне от-1 (самая темная) и 1 (самое яркое) с 0 для исходного цвета.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Задает двойное значение, которое осветляет или затемняет цвет границ диапазона, значение находится в диапазоне от-1 (самая темная) и 1 (самое яркое) с 0 для исходного цвета.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Возвращает количество диапазонов в объекте RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Возвращает объект диапазона в зависимости от его позиции в объекте RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Шаблон диапазона.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|HTML-код цвета, представляющий цвет шаблона диапазона, для #RRGGBB формы (например, "FFA500") или в виде именованного цвета HTML (например, "Апельсин").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Задает двойное значение, которое осветляет или затемняет цвет узора для заливки диапазона, значение находится в пределах от-1 (самая темная) и 1 (самое яркое) с 0 для исходного цвета.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Задает двойное значение, которое осветляет или затемняет цвет заливки диапазона, значение находится в диапазоне от-1 (самая темная) и 1 (самое яркое) с 0 для исходного цвета.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Указывает состояние зачеркивания шрифта.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Задает состояние подстрочного шрифта.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Задает состояние шрифта в виде надстрочного знака.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Задает двойное значение, которое осветляет или затемняет цвет шрифта диапазона, значение находится в диапазоне от-1 (самая темная) и 1 (самое яркое) с 0 для исходного цвета.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Указывает, отображается ли отступ текста автоматически, если для выравнивания текста задано равномерное распределение.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|Направление чтения для диапазона.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Указывает, сжимается ли текст автоматически в соответствии с шириной доступной ширины столбца.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Количество повторяющихся строк, удаленных операцией.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Количество оставшихся уникальных строк, присутствующих в получившемся диапазоне.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Указывает, должно ли соответствие быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Указывает, учитывается ли регистр в выражении.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Представляет свойство `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Представляет свойство `rowIndex`.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Указывает, должно ли соответствие быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Указывает, учитывается ли регистр в выражении.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Указывает направление поиска.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Представляет свойство `format`.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Представляет свойство `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Представляет свойство `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Представляет свойство `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[Формат: Excel. Целлпропертиесформат & {columnWidth?](/javascript/api/excel/excel.settablecolumnproperties#format)|Представляет свойство `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[Формат: Excel. Целлпропертиесформат & {rowHeight?](/javascript/api/excel/excel.settablerowproperties#format)|Представляет свойство `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Представляет свойство `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Задает текст альтернативного описания для объекта Shape.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Указывает альтернативный текст заголовка для объекта Shape.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Удаляет фигуру с листа.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Указывает тип геометрической фигуры для этой геометрической фигуры.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Преобразует фигуру в изображение и возвращает изображение в виде строки в кодировке base64.|
||[height](/javascript/api/excel/excel.shape#height)|Задает высоту фигуры в пунктах.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Перемещает фигуру по горизонтали на указанное число пунктов.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|Поворачивает фигуру по часовой стрелке относительно оси Z на указанное число градусов.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Перемещает фигуру по вертикали на указанное число пунктов.|
||[left](/javascript/api/excel/excel.shape#left)|Расстояние в пунктах от левого края фигуры до левого края листа.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Указывает, заблокировано ли пропорции данной фигуры.|
||[name](/javascript/api/excel/excel.shape#name)|Задает имя фигуры.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Возвращает количество точек соединения на фигуре.|
||[fill](/javascript/api/excel/excel.shape#fill)|Возвращает формат заливки фигуры.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Возвращает геометрическую фигуру, связанную с линией.|
||[group](/javascript/api/excel/excel.shape#group)|Возвращает группу фигур, связанную с фигурой.|
||[id](/javascript/api/excel/excel.shape#id)|Задает идентификатор фигуры.|
||[image](/javascript/api/excel/excel.shape#image)|Возвращает изображение, связанное с фигурой.|
||[level](/javascript/api/excel/excel.shape#level)|Задает уровень указанной фигуры.|
||[line](/javascript/api/excel/excel.shape#line)|Возвращает линию, связанную с фигурой.|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Возвращает формат линии для фигуры.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Возникает, если фигура активирована.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Возникает, если фигура деактивирована.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Задает родительскую группу этой фигуры.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Возвращает объект рамки с текстом для фигуры.|
||[type](/javascript/api/excel/excel.shape#type)|Возвращает тип фигуры.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|Возвращает положение указанной фигуры по оси Z. Значение 0 представляет нижнее положение по оси.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Задает поворот фигуры в градусах.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Масштабирует высоту фигуры с применением указанного коэффициента.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Масштабирует ширину фигуры с применением указанного коэффициента.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Перемещает указанную фигуру вверх или вниз по оси Z в коллекции, что переносит ее вперед или назад относительно других фигур.|
||[top](/javascript/api/excel/excel.shape#top)|Расстояние в пунктах от верхнего края фигуры до верхнего края листа.|
||[visible](/javascript/api/excel/excel.shape#visible)|Указывает, отображается ли фигура.|
||[width](/javascript/api/excel/excel.shape#width)|Задает ширину фигуры в пунктах.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Получает идентификатор активированной фигуры.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором активирована фигура.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Добавляет геометрическую фигуру на лист.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Группирует подмножество фигур на листе этой коллекции.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Создает изображение из строки в кодировке base64 и добавляет его на лист.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Добавляет линию на лист.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Добавляет текстовое поле на лист с указанным текстом в качестве содержимого.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Возвращает количество фигур на листе.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Получает фигуру по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Получает фигуру с помощью ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Получает идентификатор деактивированной фигуры.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором деактивирована фигура.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Очищает формат заливки фигуры.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Представляет цвет переднего плана заливки фигуры в формате HTML #RRGGBB (например, "FFA500") или в виде именованного цвета HTML (например, "Апельсин")|
||[type](/javascript/api/excel/excel.shapefill#type)|Возвращает тип заливки фигуры.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Задает заливку одним цветом для фигуры.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Задает процент прозрачности заливки в виде значения от 0,0 (непрозрачный) до 1,0 (Clear).|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.shapefont#color)|HTML-код цвета цвета текста (например, "#FF0000" представляет собой красный цвет).|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.shapefont#name)|Представляет имя шрифта (например, "Calibri").|
||[size](/javascript/api/excel/excel.shapefont#size)|Представляет размер шрифта в пунктах (например, 11).|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Тип подчеркивания, применяемый для шрифта.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Задает идентификатор фигуры.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Возвращает объект Shape, связанный с группой.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Возвращает коллекцию объектов Shape.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Отменяет группировку любых сгруппированных фигур в указанной группе фигур.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Представляет цвет линии в формате HTML формы #RRGGBB (например, "FFA500") или в виде именованного цвета HTML (например, "Апельсин").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Представляет тип линии фигуры.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Представляет тип линии фигуры.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Представляет степень прозрачности указанной линии как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Указывает, отображается ли форматирование линии элемента фигуры.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Представляет толщину линии (в пунктах).|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Задает подполе, которое является именем целевого свойства форматированного значения для сортировки.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Получает количество стилей в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Получает стиль на основе его позиции в коллекции.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|Представляет объект AutoFilter таблицы.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Получает идентификатор добавленной таблицы.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Получает идентификатор листа, в который добавлена таблица.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|Получает сведения о сведениях об изменении.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Возникает, если в книгу добавлена новая таблица.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Возникает, если указанная таблица удалена из книги.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Получает идентификатор удаляемой таблицы.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Получает имя таблицы, которая удалена.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Получает идентификатор листа, в котором удаляется таблица.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Получает количество таблиц в коллекции.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Получает первую таблицу в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Получает таблицу по имени или идентификатору.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|Параметры автоматической настройки размера для текстового фрейма.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Удаляет весь текст в рамке с текстом.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Представляет горизонтальное выравнивание рамки с текстом.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Представляет действие горизонтального переполнения рамки с текстом.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Представляет левое поле рамки с текстом (в пунктах).|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Представляет угол, на который текст ориентирован для текстового фрейма.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Представляет направление чтения рамки с текстом (слева направо или справа налево).|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Указывает, содержит ли текстовая рамка текст.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Представляет текст, присоединенный к фигуре в текстовой рамке, а также свойства и методы для операций с текстом.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Представляет правое поле рамки с текстом (в пунктах).|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Представляет вертикальное выравнивание для рамки с текстом.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Представляет действие вертикального переполнения рамки с текстом.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Возвращает объект TextRange для подстроки в указанном диапазоне.|
||[font](/javascript/api/excel/excel.textrange#font)|Возвращает объект ShapeFont, представляющий атрибуты шрифта для диапазона текста.|
||[text](/javascript/api/excel/excel.textrange#text)|Представляет содержимое с обычным текстом в диапазоне текста.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|Значение true, если все диаграммы в книге отслеживают точки фактических данных, с которыми они связаны.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Получает текущую активную диаграмму в книге.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Получает текущую активную диаграмму в книге.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|Значение true, если книга редактируется несколькими пользователями (совместное редактирование).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Получает текущий выделенный диапазон (один или несколько) в книге.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Указывает, были ли внесены изменения с момента последнего сохранения книги.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Указывает, находится ли книга в режиме автосохранения.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Возвращает номер версии модуля вычислений Excel.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Возникает при изменении параметра автосохранения для книги.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Указывает, была ли книга ранее сохранена локально или в сети.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|Значение true, если вычисления в книге выполняются только с той точностью чисел, с которой они отображаются.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Получает тип события.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Определяет, должен ли Excel пересчитывать лист при необходимости.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Находит все вхождения определенной строки на основе указанных условий и возвращает их в виде объекта RangeAreas, состоящего из одного или нескольких прямоугольных диапазонов.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Находит все вхождения определенной строки на основе указанных условий и возвращает их в виде объекта RangeAreas, состоящего из одного или нескольких прямоугольных диапазонов.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Получает объект RangeAreas, представляющий один или несколько блоков прямоугольных диапазонов, указанных по адресу или имени.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Представляет объект AutoFilter листа.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Получает коллекцию горизонтальных разрывов страницы для листа.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Возникает, если изменен формат указанного листа.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Получает объект PageLayout листа.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Возвращает коллекцию всех объектов Shape на листе.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Получает коллекцию вертикальных разрывов страницы для листа.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Находит и заменяет определенную строку на основе условий, указанных в текущем листе.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Представляет сведения об изменении сведений.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Возникает при изменении любого листа в книге.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Возникает при изменении формата любого листа в книге.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Возникает при изменениях выделения на любом листе.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Указывает, должно ли соответствие быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Указывает, учитывается ли регистр в выражении.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

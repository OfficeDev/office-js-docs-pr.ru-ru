---
title: Набор обязательных элементов API JavaScript для Excel 1,9
description: Сведения о наборе требований ExcelApi 1,9.
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a32705cc7557ae2f6f7214dd05f7a757188aba4c
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819660"
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
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Возвращает версию модуля вычислений Excel, использованного для последнего полного пересчета. Только для чтения.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Возвращает состояние вычисления приложения. Дополнительные сведения см. в статье Excel.CalculationState. Только для чтения.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Возвращает параметры итеративных вычислений.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Приостанавливает обновление экрана до вызова следующего метода context.sync().|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Применяет автофильтр к диапазону. При этом фильтруется столбец, если указаны индекс столбца и условия фильтрации.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Очищает условия фильтрации автофильтра.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Возвращает объект Range, представляющий диапазон, к которому применяется автофильтр.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Возвращает объект Range, представляющий диапазон, к которому применяется автофильтр.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Массив, содержащий все условия фильтрации в диапазоне с примененным автофильтром. Только для чтения.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Указывает, включен ли автофильтр. Только для чтения.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Указывает, есть ли в автофильтре условия фильтрации. Только для чтения.|
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
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Представляет значение после изменения. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Представляет значение до изменения. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Представляет тип значения после изменения|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Представляет тип значения до изменения|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Активирует диаграмму в пользовательском интерфейсе Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Объединяет параметры для сводной диаграммы. Только для чтения.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Возвращает или задает цветовую схему диаграммы. Для чтения и записи.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|Указывает, содержит ли область диаграммы скругленные углы. Для чтения и записи.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Указывает, разрешен ли выход за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Указывает, разрешен ли выход за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Возвращает или задает количество интервалов в гистограмме или диаграмме Парето. Для чтения и записи.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Возвращает или задает значение выхода за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Возвращает или задает тип интервалов для гистограммы или диаграммы Парето. Для чтения и записи.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Возвращает или задает значение выхода за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Возвращает или задает значение ширины интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Возвращает или задает тип вычисления квартилей для диаграммы "ящик с усами". Для чтения и записи.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Указывает, отображаются ли внутренние точки на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Указывает, отображается ли линия медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Указывает, отображается ли маркер медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Указывает, отображаются ли точки выбросов на диаграмме "ящик с усами". Для чтения и записи.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Указывает, содержат ли планки погрешностей точки с конечным стилем.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Указывает, какие части планок погрешностей нужно включить.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Указывает тип форматирования планок погрешностей.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|Тип диапазона, помеченного планками погрешностей.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Указывает, отображаются ли планки погрешностей.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Представляет форматирование линий диаграммы.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Возвращает или задает стратегию подписей карт ряда для диаграммы с картой региона. Для чтения и записи.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Возвращает или задает уровень карты ряда для диаграммы с картой региона. Для чтения и записи.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Возвращает или задает тип проекции ряда для диаграммы с картой региона. Для чтения и записи.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Указывает, следует ли отображать кнопки поля оси в сводной диаграмме. Свойство ShowAxisFieldButtons соответствует команде "Показать кнопки поля оси" в раскрывающемся списке "Кнопки полей" вкладки "Анализировать", доступной при выделении сводной диаграммы.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Указывает, следует ли отображать кнопки поля легенды в сводной диаграмме.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Указывает, следует ли отображать кнопки поля фильтра отчета в сводной диаграмме.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Указывает, следует ли отображать кнопки поля значения в сводной диаграмме.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Может быть целым числом от 0 (нуля) до 300, представляющим процентное значение от размера по умолчанию. Это свойство применяется только к пузырьковым диаграммам. Для чтения и записи.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Возвращает или задает цвет максимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Возвращает или задает тип максимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Возвращает или задает максимальное значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Возвращает или задает цвет среднего значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Возвращает или задает тип среднего значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Возвращает или задает среднее значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Возвращает или задает цвет минимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Возвращает или задает тип минимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Возвращает или задает минимальное значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Возвращает или задает стиль градиента ряда для диаграммы с картой региона. Для чтения и записи.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Возвращает или задает цвет заливки для точек отрицательных данных в ряду. Для чтения и записи.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Возвращает или задает область стратегии родительских подписей ряда для диаграммы "дерево". Для чтения и записи.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Объединяет параметры интервалов для гистограмм и диаграмм Парето. Только для чтения.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Объединяет параметры для диаграмм "ящик с усами" Только для чтения.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Объединяет параметры для диаграммы с картой региона. Только для чтения.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Указывает, отображаются ли соединительные линии в каскадных диаграммах. Для чтения и записи.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Указывает, отображаются ли линии выноски для каждой подписи данных в ряду. Для чтения и записи.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Возвращает или задает пороговое значение, разделяющее два раздела вторичной круговой диаграммы или вторичной гистограммы. Для чтения и записи.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Представляет свойство `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Представляет свойство `columnIndex`.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, к которым применено условное форматирование. Только для чтения.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, с недопустимыми значениями ячеек. Если все значения ячеек являются допустимыми, эта функция выдаст ошибку ItemNotFound.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, с недопустимыми значениями ячеек. Если все значения ячеек являются допустимыми, эта функция вернет значение null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|Свойство, используемое фильтром для расширенной фильтрации по объектам richvalue.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Возвращает идентификатор фигуры. Только для чтения.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Возвращает объект Shape для геометрической фигуры. Только для чтения.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Возвращает количество фигур в группе фигур. Только для чтения.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Получает фигуру по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Получает фигуру на основе ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Получает или задает центральный нижний колонтитул листа.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Получает или задает центральный верхний колонтитул листа.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Получает или задает левый нижний колонтитул листа.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Получает или задает левый верхний колонтитул листа.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Получает или задает правый нижний колонтитул листа.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Получает или задает правый верхний колонтитул листа.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|Общий колонтитул, используемый для всех страниц, если не указан колонтитул четных и нечетных страниц или первой страницы.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|Колонтитул для четных страниц, для нечетных страниц нужно указывать отдельный колонтитул.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|Колонтитул первой страницы, для остальных страниц используется общий или четный и нечетный колонтитулы.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|Колонтитул для нечетных страниц, для четных страниц нужно указывать отдельный колонтитул.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Получает или задает состояние, в котором находятся колонтитулы. Дополнительные сведения см. в статье Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Получает или задает отметку, которая указывает, выровнены ли колонтитулы относительно полей страницы, установленных в параметрах макета страницы для листа.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Получает или задает отметку, которая указывает, нужно ли масштабировать колонтитулы с помощью процентных значений, установленных в параметрах макета страницы для листа.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Возвращает формат изображения. Только для чтения.|
||[id](/javascript/api/excel/excel.image#id)|Представляет идентификатор фигуры для объекта image. Только для чтения.|
||[shape](/javascript/api/excel/excel.image#shape)|Возвращает объект Shape, связанный с изображением. Только для чтения.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Значение true, если в Excel используется итерация для разрешения циклических ссылок.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Возвращает или задает максимальное изменение между итерациями при разрешении в Excel циклических ссылок.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Возвращает или задает максимальное количество итераций, которое можно использовать в Excel для разрешения циклической ссылки.|
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
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Представляет фигуру, к которой привязано начало указанной линии. Только для чтения.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Представляет точку соединения, к которой привязано начало соединительной линии. Только для чтения. Возвращает значение null, если начало линии не привязано к фигуре.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Представляет фигуру, к которой привязан конец указанной линии. Только для чтения.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Представляет точку соединения, к которой привязан конец соединительной линии. Только для чтения. Возвращает значение null, если конец линии не привязан к фигуре.|
||[id](/javascript/api/excel/excel.line#id)|Представляет идентификатор фигуры. Только для чтения.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Указывает, привязано ли начало указанной линии к фигуре. Только для чтения.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Указывает, привязан ли конец указанной линии к фигуре. Только для чтения.|
||[shape](/javascript/api/excel/excel.line#shape)|Возвращает объект Shape, связанный с линией. Только для чтения.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Удаляет объект разрыва страницы.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Получает первую ячейку после разрыва страницы.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Представляет индекс столбца для разрыва страницы|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Представляет индекс строки для разрыва страницы|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Добавляет разрыв страницы перед левой верхней ячейкой указанного диапазона.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Получает количество разрывов страниц в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Получает объект разрыва страницы по индексу.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Сбрасывает все добавленные вручную разрывы страниц в коллекции.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|Получает или задает параметр черно-белой печати листа.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|Получает или задает нижнее поле страницы листа, чтобы использовать для печати в пунктах.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|Получает или задает отметку выравнивания листа по горизонтали относительно центра. Эта отметка определяет, выравнивается ли лист по горизонтали относительно центра при печати.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|Получает или задает отметку выравнивания листа по вертикали относительно центра. Эта отметка определяет, выравнивается ли лист по вертикали относительно центра при печати.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|Получает или задает параметр режима черновика листа. Если присвоено значение true, лист будет печататься без рисунков.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|Получает или задает номер первой страницы листа для печати. Значение null представляет автоматическую нумерацию страниц.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|Получает или задает поле нижнего колонтитула листа (в пунктах) для использования при печати.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Получает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, представляющих область печати для листа. Если область печати отсутствует, возникает ошибка ItemNotFound.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Получает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, представляющих область печати для листа. Если область печати отсутствует, возвращается пустой объект.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Получает объект range, представляющий столбцы заголовков.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Получает объект range, представляющий столбцы заголовков. Если значение не установлено, возвращается пустой объект.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Получает объект range, представляющий строки заголовков.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Получает объект range, представляющий строки заголовков. Если значение не установлено, возвращается пустой объект.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|Получает или задает поле верхнего колонтитула листа (в пунктах) для использования при печати.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|Получает или задает левое поле листа (в пунктах) для использования при печати.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|Получает или задает ориентацию страницы для листа.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|Получает или задает размер бумаги для листа.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Получает или задает, должны ли отображаться примечания листа при печати.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|Получает или задает параметр ошибок печати листа.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Получает или задает отметку печати линий сетки листа. Эта отметка определяет, печатаются ли линии сетки.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Получает или задает отметку печати заголовков листа. Эта отметка определяет, печатаются ли заголовки.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|Получает или задает параметр порядка печати листа. Определяет порядок, использующийся при обработке распечатываемых номеров страниц.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Настройка колонтитулов для листа.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|Получает или задает правое поле листа (в пунктах) для использования при печати.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Задает область печати листа.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Задает поля страницы с единицами измерения для листа.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Задает столбцы, содержащие ячейки, которые должны повторяться слева на каждой странице при печати листа.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Задает строки, содержащие ячейки, которые должны повторяться сверху каждой страницы при печати листа.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Получает или задает верхнее поле листа (в пунктах) для использования при печати.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Получает или задает параметры масштабирования при печати листа.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Представляет нижнее поле макета страницы в указанных единицах измерения для использования при печати.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Представляет поле нижнего колонтитула макета страницы в указанных единицах измерения для использования при печати.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Представляет поле верхнего колонтитула макета страницы в указанных единицах измерения для использования при печати.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Представляет левое поле макета страницы в указанных единицах измерения для использования при печати.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Представляет правое поле макета страницы в указанных единицах измерения для использования при печати.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Представляет верхнее поле макета страницы в указанных единицах измерения для использования при печати.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Количество страниц, размещаемых по горизонтали. Это значение может быть равно null, если используется процентный масштаб.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|Значение масштаба печатной страницы может быть равным от 10 до 400. Это значение может быть равно null, если указано размещение по высоте или ширине страницы.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Количество страниц, размещаемых по вертикали. Это значение может быть равно null, если используется процентный масштаб.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Сортирует сводную таблицу по указанным значениям в определенной области. Область определяет, какие конкретные значения будут использоваться для сортировки|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Указывает, применяется ли форматирование автоматически при его обновлении или перемещении полей|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Получает объект DataHierarchy, использующийся для вычисления значения в указанном диапазоне сводной таблицы.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Получает объекты PivotItem с оси, образующие значение в указанном диапазоне сводной таблицы.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Указывает, сохраняется ли форматирование при обновлении или пересчете отчета с помощью таких операций, как сведение, сортировка или изменение элементов полей страницы.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Задает для сводной таблицы автоматическую сортировку, используя указанную ячейку, чтобы автоматически выбрать все необходимые условия и контекст. Это работает аналогично применению автоматической сортировки из пользовательского интерфейса.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Указывает, разрешается ли пользователю изменять значения данных сводной таблицы.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Указывает, используются ли при сортировке в сводной таблице настраиваемые списки.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Заполняет конечный диапазон из текущего диапазона.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Преобразует диапазон ячеек с типами данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Преобразует ячейки диапазона в связанный тип данных на листе.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Копирует данные ячейки или форматирование из исходного диапазона или объекта RangeAreas в текущий диапазон.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Находит определенную строку на основе указанных условий.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Находит определенную строку на основе указанных условий.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Выполняет мгновенное заполнение текущего диапазона. Функция мгновенного заполнения автоматически подставляет данные, когда обнаруживает закономерность, поэтому диапазон должен состоять из одного столбца со смежными данными, чтобы выявить закономерность.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Возвращает двумерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой ячейки.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждого столбца.  Для свойств, не являющихся одинаковыми в каждой ячейке определенного столбца, возвращается значение null.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой строки.  Для свойств, не являющихся одинаковыми в каждой ячейке определенной строки, возвращается значение null.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Получает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, представляющих все ячейки, которые соответствуют указанному типу и значению.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Получает объект RangeAreas, состоящий из одного или нескольких диапазонов, представляющих все ячейки, которые соответствуют указанному типу и значению.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Получает коллекцию таблиц с заданной областью, перекрывающую диапазон.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Представляет состояние типа данных каждой ячейки. Только для чтения.|
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
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Возвращает объект RangeAreas, представляющий пересечение заданных диапазонов или RangeAreas. Если пересечение не найдено, возвращается сообщение об ошибке ItemNotFound.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Возвращает объект RangeAreas, представляющий пересечение заданных диапазонов или RangeAreas. Если пересечение не найдено, возвращается пустой объект.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Возвращает объект RangeAreas, смещенный на определенное количество строк и столбцов. Измерение возвращаемого объекта RangeAreas будет соответствовать исходному объекту. Если результирующий объект RangeAreas выходит за пределы таблицы листа, возникнет ошибка.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Возвращает объект RangeAreas, представляющий все ячейки, которые соответствуют указанному типу и значению. Выдает ошибку, если не найдено специальных ячеек, соответствующих условиям. |
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Возвращает объект RangeAreas, представляющий все ячейки, которые соответствуют указанному типу и значению. Возвращает пустой объект, если не найдено специальных ячеек, соответствующих условиям. |
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Возвращает коллекцию таблиц с заданной областью, перекрывающую любой диапазон в объекте RangeAreas.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Возвращает использованный объект RangeAreas, включающий все использованные области отдельных прямоугольных диапазонов в объекте RangeAreas.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Возвращает использованный объект RangeAreas, включающий все использованные области отдельных прямоугольных диапазонов в объекте RangeAreas.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Возвращает ссылку на RageAreas в стиле A1. Значение адреса содержит имя листа для каждого прямоугольного блока или ячейки (например, "Лист1!A1:B4, Лист1!D1:D4"). Только для чтения.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Возвращает ссылку на RageAreas в языковом стандарте пользователя. Только для чтения.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Возвращает количество прямоугольных диапазонов, составляющих этот объект RangeAreas.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Возвращает коллекцию прямоугольных диапазонов, составляющих этот объект RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Возвращает число ячеек в объекте RangeAreas с суммированием количества ячеек всех отдельных прямоугольных диапазонов. Возвращает значение -1, если количество ячеек превышает 2^31-1 (2 147 483 647). Только для чтения.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Возвращает коллекцию объектов ConditionalFormat, пересекающихся с любыми ячейками в этом объекте RangeAreas. Только для чтения.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Возвращает объект dataValidation для всех диапазонов в объекте RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Возвращает объект rangeFormat, в который включены шрифт, заливка, границы, выравнивание и другие свойства всех диапазонов в объекте RangeAreas. Только для чтения.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Указывает, представляют ли все диапазоны в объекте RangeAreas целые столбцы (например, "A:C, Q:Z"). Только для чтения.|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Указывает, представляют ли все диапазоны в объекте RangeAreas целые строки (например, "1:3, 5:7"). Только для чтения.|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Возвращает лист для текущего объекта RangeAreas. Только для чтения.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Устанавливает объект RangeAreas, предназначенный для пересчета при выполнении следующего пересчета.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Представляет стиль всех диапазонов в этом объекте RangeAreas.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для границы диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для границ диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Возвращает количество диапазонов в объекте RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Возвращает объект диапазона в зависимости от его позиции в объекте RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Получает или задает шаблон объекта Range. Дополнительные сведения см. в статье Excel.FillPattern. LinearGradient и RectangularGradient не поддерживаются.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Задает HTML-код, представляющий шаблон объекта Range в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет шаблона для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Указывает, зачеркнут ли шрифт. Значение null указывает, что для всего диапазона не применяется единый параметр зачеркивания.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Указывает, является ли шрифт подстрочным.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Указывает, является ли шрифт надстрочным.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для шрифта диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста установлено на равномерное распределение.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|Направление чтения для диапазона.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Количество повторяющихся строк, удаленных операцией.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Количество оставшихся уникальных строк, присутствующих в получившемся диапазоне.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Указывает, должно ли совпадение быть полным или частичным. Значение по умолчанию: false (частичное).|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Указывает, учитывается ли регистр при сопоставлении. Значение по умолчанию: false (без учета регистра).|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Представляет свойство `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Представляет свойство `rowIndex`.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Указывает, должно ли совпадение быть полным или частичным. Полное совпадение соответствует всему содержимому ячейки. Значение по умолчанию: false (частичное).|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Указывает, учитывается ли регистр при сопоставлении. Значение по умолчанию: false (без учета регистра).|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Указывает направление поиска. Значение по умолчанию: вперед. См. статью Excel.SearchDirection.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Представляет свойство `format`.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Представляет свойство `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Представляет свойство `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Представляет свойство `columnHidden`.|
||[format: Excel.CellPropertiesFormat](/javascript/api/excel/excel.settablecolumnproperties#format)|Представляет свойство `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel.CellPropertiesFormat](/javascript/api/excel/excel.settablerowproperties#format)|Представляет свойство `format`.|
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Представляет свойство `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Возвращает или задает замещающий текст описания для объекта Shape.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Возвращает или задает замещающий текст заголовка для объекта Shape.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Удаляет фигуру с листа.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Представляет геометрический тип фигуры. Дополнительные сведения см. в статье Excel.GeometricShapeType. Возвращает значение null, если тип фигуры отличается от GeometricShape.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Преобразует фигуру в изображение и возвращает изображение в виде строки в кодировке base64. Число точек на дюйм: 96. Единственные поддерживаемые форматы: `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG` и `Excel.PictureFormat.GIF`.|
||[height](/javascript/api/excel/excel.shape#height)|Представляет высоту фигуры (в пунктах).|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Перемещает фигуру по горизонтали на указанное число пунктов.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|Поворачивает фигуру по часовой стрелке относительно оси Z на указанное число градусов.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Перемещает фигуру по вертикали на указанное число пунктов.|
||[left](/javascript/api/excel/excel.shape#left)|Расстояние в пунктах от левого края фигуры до левого края листа.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Указывает, заблокированы ли пропорции фигуры.|
||[name](/javascript/api/excel/excel.shape#name)|Представляет название фигуры.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Возвращает количество точек соединения на фигуре. Только для чтения.|
||[fill](/javascript/api/excel/excel.shape#fill)|Возвращает формат заливки фигуры. Только для чтения.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Возвращает геометрическую фигуру, связанную с линией. Если тип фигуры отличается от GeometricShape, возникает ошибка.|
||[group](/javascript/api/excel/excel.shape#group)|Возвращает группу фигур, связанную с фигурой. Если тип фигуры отличается от GroupShape, возникает ошибка.|
||[id](/javascript/api/excel/excel.shape#id)|Представляет идентификатор фигуры. Только для чтения.|
||[image](/javascript/api/excel/excel.shape#image)|Возвращает изображение, связанное с фигурой. Если тип фигуры отличается от Image, возникает ошибка.|
||[level](/javascript/api/excel/excel.shape#level)|Представляет уровень указанной фигуры. Например, уровень 0 означает, что фигура не является частью групп; уровень 1 означает, что фигура является частью группы верхнего уровня; уровень 2 означает, что фигура является частью подгруппы верхнего уровня.|
||[line](/javascript/api/excel/excel.shape#line)|Возвращает линию, связанную с фигурой. Если тип фигуры отличается от Line, возникает ошибка.|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Возвращает формат линии для фигуры. Только для чтения.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Возникает, если фигура активирована.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Возникает, если фигура деактивирована.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Представляет родительскую группу фигуры.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Возвращает объект рамки с текстом для фигуры. Только для чтения.|
||[type](/javascript/api/excel/excel.shape#type)|Возвращает тип фигуры. Дополнительные сведения см. в статье Excel.ShapeType. Только для чтения.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|Возвращает положение указанной фигуры по оси Z. Значение 0 представляет нижнее положение по оси. Только для чтения.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Представляет поворот фигуры в градусах.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Масштабирует высоту фигуры с применением указанного коэффициента. Для изображений можно указать изменение масштаба фигуры относительно исходного или текущего размера. Фигуры, не являющиеся изображениями, всегда масштабируются относительно их текущей высоты.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Масштабирует ширину фигуры с применением указанного коэффициента. Для изображений можно указать изменение масштаба фигуры относительно исходного или текущего размера. Фигуры, не являющиеся изображениями, всегда масштабируются относительно их текущей ширины.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Перемещает указанную фигуру вверх или вниз по оси Z в коллекции, что переносит ее вперед или назад относительно других фигур.|
||[top](/javascript/api/excel/excel.shape#top)|Расстояние в пунктах от верхнего края фигуры до верхнего края листа.|
||[visible](/javascript/api/excel/excel.shape#visible)|Представляет видимость фигуры.|
||[width](/javascript/api/excel/excel.shape#width)|Представляет ширину фигуры (в пунктах).|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Получает идентификатор активированной фигуры.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором активирована фигура.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Добавляет геометрическую фигуру на лист. Возвращает объект Shape, представляющий новую фигуру.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Группирует подмножество фигур на листе этой коллекции. Возвращает объект Shape, представляющий новую группу фигур.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Создает изображение из строки в кодировке base64 и добавляет его на лист. Возвращает объект Shape, представляющий новое изображение.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Добавляет линию на лист. Возвращает объект Shape, представляющий новую линию.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Добавляет текстовое поле на лист с указанным текстом в качестве содержимого. Возвращает объект Shape, представляющий новое текстовое поле.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Возвращает количество фигур на листе. Только для чтения.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Получает фигуру по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Получает фигуру с помощью ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Получает идентификатор деактивированной фигуры.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором деактивирована фигура.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Очищает формат заливки фигуры.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Представляет цвет переднего плана заливки фигуры в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[type](/javascript/api/excel/excel.shapefill#type)|Возвращает тип заливки фигуры. Только для чтения. Дополнительные сведения см. в статье Excel.ShapeFillType.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Задает заливку одним цветом для фигуры. При этом тип заливки изменяется на сплошную.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Возвращает или задает процентное значение прозрачности заливки как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если тип фигуры не поддерживает прозрачность или заливка фигуры имеет несогласованную прозрачность, например при использовании градиентной заливки.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Указывает, является ли шрифт полужирным. Возвращает значение null, если объект TextRange включает фрагменты как с полужирным, так и без полужирного текста.|
||[color](/javascript/api/excel/excel.shapefont#color)|HTML-код цвета текста (например, значение #FF0000 обозначает красный цвет). Возвращает значение null, если объект TextRange включает фрагменты текста с разными цветами.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Указывает, применяется ли курсив. Возвращает значение null, если объект TextRange включает фрагменты текста как выделенные, так и не выделенные курсивом.|
||[name](/javascript/api/excel/excel.shapefont#name)|Представляет имя шрифта (например, Calibri). Если текст является набором сложных знаков или написан на восточноазиатских языках, этот параметр является соответствующим именем шрифта. В противном случае это имя шрифта на латинице.|
||[size](/javascript/api/excel/excel.shapefont#size)|Представляет размер шрифта в пунктах (например, 11). Возвращает значение null, если объект TextRange включает фрагменты текста с разными размерами шрифта.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Тип подчеркивания, применяемый для шрифта. Возвращает значение null, если объект TextRange включает фрагменты текста с разными стилями подчеркивания. Дополнительные сведения см. в статье Excel.ShapeFontUnderlineStyle.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Представляет идентификатор фигуры. Только для чтения.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Возвращает объект Shape, связанный с группой. Только для чтения.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Возвращает коллекцию объектов Shape. Только для чтения.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Отменяет группировку любых сгруппированных фигур в указанной группе фигур.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Представляет цвет линии в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные типы штриха. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные стили. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Представляет степень прозрачности указанной линии как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если в фигуре используются несогласованные параметры прозрачности.|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Указывает, отображается ли форматирование линии элемента фигуры. Возвращает значение null, если в фигуре используются несогласованные параметры видимости.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Представляет толщину линии (в пунктах). Возвращает значение null, если линия является невидимой или используются линии с несогласованной толщиной.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Представляет подполе, являющееся именем целевого свойства форматированного значения, по которому выполняется сортировка.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Получает количество стилей в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Получает стиль на основе его позиции в коллекции.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|Представляет объект AutoFilter таблицы. Только для чтения.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Получает идентификатор добавленной таблицы.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Получает идентификатор листа, в который добавлена таблица.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|Представляет сведения об изменении|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Возникает, если в книгу добавлена новая таблица.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Возникает, если указанная таблица удалена из книги.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Указывает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Указывает идентификатор удаленной таблицы.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Указывает имя удаленной таблицы.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Указывает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Указывает идентификатор листа, в котором удаляется таблица.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Получает количество таблиц в коллекции.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Получает первую таблицу в коллекции. Таблицы в коллекции сортируются сверху вниз и слева направо, поэтому верхняя левая таблица является первой в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Получает таблицу по имени или идентификатору.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|Возвращает или задает параметры автоматического подбора размера для рамки с текстом. Для рамки с текстом можно настроить автоматический подбор размера текста в соответствии с размером рамки, автоматический подбор размера рамки в соответствии с содержимым или не выполнять автоматический подбор размера.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Удаляет весь текст в рамке с текстом.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Представляет горизонтальное выравнивание рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextHorizontalAlignment.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Представляет действие горизонтального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextHorizontalOverflow.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Представляет левое поле рамки с текстом (в пунктах).|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Представляет ориентацию текста для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextOrientation.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Представляет направление чтения рамки с текстом (слева направо или справа налево). Дополнительные сведения см. в статье Excel.ShapeTextReadingOrder.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Указывает, содержится ли в текстовой рамке текст.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Представляет текст, присоединенный к фигуре в текстовой рамке, а также свойства и методы для операций с текстом. Дополнительные сведения см. в статье Excel.TextRange.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Представляет правое поле рамки с текстом (в пунктах).|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Представляет вертикальное выравнивание для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalAlignment.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Представляет действие вертикального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalOverflow.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Возвращает объект TextRange для подстроки в указанном диапазоне.|
||[font](/javascript/api/excel/excel.textrange#font)|Возвращает объект ShapeFont, представляющий атрибуты шрифта для диапазона текста. Только для чтения.|
||[text](/javascript/api/excel/excel.textrange#text)|Представляет содержимое с обычным текстом в диапазоне текста.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|Значение true, если все диаграммы в книге отслеживают точки фактических данных, с которыми они связаны.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Получает текущую активную диаграмму в книге. Если активная диаграмма отсутствует, при вызове этого оператора возникает исключение|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Получает текущую активную диаграмму в книге. Если активная диаграмма отсутствует, возвращает пустой объект|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|Значение true, если книга редактируется несколькими пользователями (совместное редактирование).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Получает текущий выделенный диапазон (один или несколько) в книге. В отличие от getSelectedRange() этот метод возвращает объект RangeAreas, представляющий все выделенные диапазоны.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Указывает, внесены ли изменения с момента последнего сохранении книги.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Указывает, применяется ли в книге режим автосохранения. Только для чтения.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Возвращает номер версии модуля вычислений Excel. Только для чтения.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Возникает при изменении параметра автосохранения для книги.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Указывает, сохранялась ли книга ранее (локально или в Интернете). Только для чтения.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|Значение true, если вычисления в книге выполняются только с той точностью чисел, с которой они отображаются.|
|[воркбукаутосавесеттингчанжедевентаргс](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Представляет тип события. Дополнительные сведения см. в статье Excel.EventType.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Получает или задает свойство enableCalculation для листа.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Находит все вхождения определенной строки на основе указанных условий и возвращает их в виде объекта RangeAreas, состоящего из одного или нескольких прямоугольных диапазонов.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Находит все вхождения определенной строки на основе указанных условий и возвращает их в виде объекта RangeAreas, состоящего из одного или нескольких прямоугольных диапазонов.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Получает объект RangeAreas, представляющий один или несколько блоков прямоугольных диапазонов, указанных по адресу или имени.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Представляет объект AutoFilter листа. Только для чтения.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Получает коллекцию горизонтальных разрывов страницы для листа. Эта коллекция содержит только добавленные вручную разрывы страниц.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Возникает, если изменен формат указанного листа.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Получает объект PageLayout листа.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Возвращает коллекцию всех объектов Shape на листе. Только для чтения.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Получает коллекцию вертикальных разрывов страницы для листа. Эта коллекция содержит только добавленные вручную разрывы страниц.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Находит и заменяет определенную строку на основе условий, указанных в текущем листе.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Представляет сведения об изменении|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Возникает при изменении любого листа в книге.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Возникает при изменении формата любого листа в книге.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Возникает при изменениях выделения на любом листе.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область конкретного листа. Может возвращать пустой объект.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Указывает, должно ли совпадение быть полным или частичным. Полное совпадение соответствует всему содержимому ячейки. Значение по умолчанию: false (частичное).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Указывает, учитывается ли регистр при сопоставлении. Значение по умолчанию: false (без учета регистра).|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

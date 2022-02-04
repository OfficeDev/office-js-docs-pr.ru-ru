---
title: Excel API JavaScript установлено 1.9
description: Сведения о наборе требований ExcelApi 1.9.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
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

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.9. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.9 или более ранних, см. Excel API в наборе требований [1.9 или ранее](/javascript/api/excel?view=excel-js-1.9&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#excel-excel-application-calculationengineversion-member)|Возвращает версию модуля вычислений Excel, использованного для последнего полного пересчета.|
||[calculationState](/javascript/api/excel/excel.application#excel-excel-application-calculationstate-member)|Возвращает состояние вычисления приложения.|
||[iterativeCalculation](/javascript/api/excel/excel.application#excel-excel-application-iterativecalculation-member)|Возвращает параметры итеративных вычислений.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendscreenupdatinguntilnextsync-member(1))|Приостанавливать обновление экрана до следующего `context.sync()` .|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-apply-member(1))|Применяет автофильтр к диапазону.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcriteria-member(1))|Очищает критерии фильтрации и сортировать состояние автофильтера.|
||[criteria](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-criteria-member)|Массив, содержащий все условия фильтрации в диапазоне с примененным автофильтром.|
||[enabled](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-enabled-member)|Указывает, включен ли autoFilter.|
||[getRange()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrange-member(1))|Возвращает объект `Range` , который представляет диапазон, к которому применяется AutoFilter.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrangeornullobject-member(1))|Возвращает объект `Range` , который представляет диапазон, к которому применяется AutoFilter.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-isdatafiltered-member)|Указывает, есть ли у autoFilter критерии фильтрации.|
||[reapply()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-reapply-member(1))|Применяет указанный объект Autofilter, находящийся в настоящее время в диапазоне.|
||[remove()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-remove-member(1))|Удаляет автофильтр из диапазона.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-color-member)|Представляет свойство `color` одинарной границы.|
||[style](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-style-member)|Представляет свойство `style` одинарной границы.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-tintandshade-member)|Представляет свойство `tintAndShade` одинарной границы.|
||[weight](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-weight-member)|Представляет свойство `weight` одинарной границы.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-bottom-member)|Представляет свойство `format.borders.bottom`.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonaldown-member)|Представляет свойство `format.borders.diagonalDown`.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonalup-member)|Представляет свойство `format.borders.diagonalUp`.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-horizontal-member)|Представляет свойство `format.borders.horizontal`.|
||[left](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-left-member)|Представляет свойство `format.borders.left`.|
||[right](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-right-member)|Представляет свойство `format.borders.right`.|
||[top](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-top-member)|Представляет свойство `format.borders.top`.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-vertical-member)|Представляет свойство `format.borders.vertical`.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-address-member)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-addresslocal-member)|Представляет свойство `addressLocal`.|
||[hidden](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-hidden-member)|Представляет свойство `hidden`.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-color-member)|Представляет свойство `format.fill.color`.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-pattern-member)|Представляет свойство `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterncolor-member)|Представляет свойство `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterntintandshade-member)|Представляет свойство `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-tintandshade-member)|Представляет свойство `format.fill.tintAndShade`.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-bold-member)|Представляет свойство `format.font.bold`.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-color-member)|Представляет свойство `format.font.color`.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-italic-member)|Представляет свойство `format.font.italic`.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-name-member)|Представляет свойство `format.font.name`.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-size-member)|Представляет свойство `format.font.size`.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-strikethrough-member)|Представляет свойство `format.font.strikethrough`.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-subscript-member)|Представляет свойство `format.font.subscript`.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-superscript-member)|Представляет свойство `format.font.superscript`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-tintandshade-member)|Представляет свойство `format.font.tintAndShade`.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-underline-member)|Представляет свойство `format.font.underline`.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-autoindent-member)|Представляет свойство `autoIndent`.|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-borders-member)|Представляет свойство `borders`.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-fill-member)|Представляет свойство `fill`.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-font-member)|Представляет свойство `font`.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-horizontalalignment-member)|Представляет свойство `horizontalAlignment`.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-indentlevel-member)|Представляет свойство `indentLevel`.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-protection-member)|Представляет свойство `protection`.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-readingorder-member)|Представляет свойство `readingOrder`.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-shrinktofit-member)|Представляет свойство `shrinkToFit`.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-textorientation-member)|Представляет свойство `textOrientation`.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member)|Представляет свойство `useStandardHeight`.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member)|Представляет свойство `useStandardWidth`.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-verticalalignment-member)|Представляет свойство `verticalAlignment`.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-wraptext-member)|Представляет свойство `wrapText`.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-formulahidden-member)|Представляет свойство `format.protection.formulaHidden`.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-locked-member)|Представляет свойство `format.protection.locked`.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valueafter-member)|Представляет значение после изменения.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuebefore-member)|Представляет значение перед изменением.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypeafter-member)|Представляет тип значения после изменения.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypebefore-member)|Представляет тип значения перед изменением.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#excel-excel-chart-activate-member(1))|Активирует диаграмму в пользовательском интерфейсе Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#excel-excel-chart-pivotoptions-member)|Объединяет параметры для сводной диаграммы.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-colorscheme-member)|Указывает цветовую схему диаграммы.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-roundedcorners-member)|Указывает, имеет ли область диаграммы закругленные углы.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-linknumberformat-member)|Указывает, связан ли формат номеров с ячейками.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowoverflow-member)|Указывает, включен ли переполнение бина в диаграмме гистограммы или диаграмме pareto.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowunderflow-member)|Указывает, включен ли недополуч бин в диаграмме гистограммы или диаграмме pareto.|
||[count](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-count-member)|Указывает количество бинов диаграммы гистограммы или диаграммы pareto.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-overflowvalue-member)|Указывает значение переполнения ячейки диаграммы гистограммы или диаграммы pareto.|
||[type](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-type-member)|Указывает тип бина для диаграммы гистограммы или диаграммы pareto.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-underflowvalue-member)|Указывает значение недополука бина для диаграммы гистограммы или диаграммы pareto.|
||[width](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-width-member)|Указывает значение ширины ячейки диаграммы гистограммы или диаграммы pareto.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-quartilecalculation-member)|Указывает, указывается ли тип квартильного вычисления диаграммы полей и усов.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showinnerpoints-member)|Указывает, показаны ли внутренние точки в поле и диаграмме усов.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanline-member)|Указывает, отображается ли в поле и диаграмме усов значимая строка.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanmarker-member)|Указывает, отображается ли маркер в поле и диаграмме усов.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showoutlierpoints-member)|Указывает, показаны ли точки выброса в поле и диаграмме усов.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-linknumberformat-member)|Указывает, связан ли формат номеров с ячейками (чтобы формат номеров менял метки при изменениях в ячейках).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-linknumberformat-member)|Указывает, связан ли формат номеров с ячейками.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-endstylecap-member)|Указывает, есть ли у баров ошибок крышка конца стиля.|
||[format](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-format-member)|Указывает тип форматирования планок погрешностей.|
||[include](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-include-member)|Указывает, какие части планок погрешностей нужно включить.|
||[type](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-type-member)|Тип диапазона, помеченного планками погрешностей.|
||[visible](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-visible-member)|Указывает, отображаются ли бары ошибок.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#excel-excel-charterrorbarsformat-line-member)|Представляет форматирование линий диаграммы.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-labelstrategy-member)|Указывает стратегию меток на карте серии на диаграмме карты региона.|
||[level](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-level-member)|Указывает уровень сопоставления ряда диаграммы карты региона.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-projectiontype-member)|Указывает тип проекции серии диаграммы карты региона.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showaxisfieldbuttons-member)|Указывает, следует ли отображать кнопки поля оси на сводная диаграмма.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showlegendfieldbuttons-member)|Указывает, следует ли отображать кнопки поля легенды на сводная диаграмма.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showreportfilterfieldbuttons-member)|Указывает, следует ли отображать кнопки поля фильтрации отчетов на сводная диаграмма.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showvaluefieldbuttons-member)|Указывает, следует ли отображать кнопки поля отображения значения на сводная диаграмма.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[binOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-binoptions-member)|Объединяет параметры интервалов для гистограмм и диаграмм Парето.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-boxwhiskeroptions-member)|Объединяет параметры для диаграмм "ящик с усами"|
||[bubbleScale](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-bubblescale-member)|Может быть целым числом от 0 (нуля) до 300, представляющим процентное значение от размера по умолчанию.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumcolor-member)|Указывает цвет для максимального значения серии диаграммы карты региона.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumtype-member)|Указывает тип для максимального значения серии диаграммы карты региона.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumvalue-member)|Указывает максимальное значение серии диаграммы карты региона.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointcolor-member)|Указывает цвет для значения средней точки серии диаграммы карты региона.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointtype-member)|Указывает тип для значения средней точки серии диаграммы карты региона.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointvalue-member)|Указывает значение средней точки серии диаграммы карты региона.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumcolor-member)|Указывает цвет для минимального значения серии диаграммы карты региона.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumtype-member)|Указывает тип для минимального значения серии диаграммы карты региона.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumvalue-member)|Указывает минимальное значение серии диаграммы карты региона.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientstyle-member)|Указывает стиль градиента серии диаграммы карты региона.|
||[invertColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertcolor-member)|Указывает цвет заполнения для отрицательных точек данных в серии.|
||[mapOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-mapoptions-member)|Объединяет параметры для диаграммы с картой региона.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-parentlabelstrategy-member)|Указывает область стратегии родительской метки серии для диаграммы treemap.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showconnectorlines-member)|Указывает, показаны ли линии соединители в диаграммах водопада.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showleaderlines-member)|Указывает, отображаются ли строки лидеров для каждой метки данных в серии.|
||[splitValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splitvalue-member)|Указывает пороговое значение, которое разделяет два раздела диаграммы пирога или диаграммы "окантовка пирога".|
||[xErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-xerrorbars-member)|Представляет объект планки погрешностей для ряда диаграммы.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-yerrorbars-member)|Представляет объект планки погрешностей для ряда диаграммы.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-linknumberformat-member)|Указывает, связан ли формат номеров с ячейками (чтобы формат номеров менял метки при изменениях в ячейках).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-address-member)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-addresslocal-member)|Представляет свойство `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-columnindex-member)|Представляет свойство `columnIndex`.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getranges-member(1))|Возвращает один `RangeAreas`или несколько прямоугольных диапазонов, к которым применяется кондитональный формат.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcells-member(1))|Возвращает объект `RangeAreas` , состоящий из одного или нескольких прямоугольных диапазонов, с недействительными значениями ячейки.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcellsornullobject-member(1))|Возвращает объект `RangeAreas` , состоящий из одного или нескольких прямоугольных диапазонов, с недействительными значениями ячейки.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#excel-excel-filtercriteria-subfield-member)|Свойство, используемее фильтром для фильтрации богатых значений.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-id-member)|Возвращает идентификатор фигуры.|
||[shape](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-shape-member)|Возвращает объект для `Shape` геометрической фигуры.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getcount-member(1))|Возвращает количество фигур в группе фигур.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitem-member(1))|Получает фигуру с ее именем или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemat-member(1))|Получает фигуру на основе ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerfooter-member)|В центре таблицы.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerheader-member)|Заглавный заглавный центр таблицы.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftfooter-member)|Левый футер таблицы.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftheader-member)|Левый заготок таблицы.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightfooter-member)|Правый ступник таблицы.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightheader-member)|Правый заготок таблицы.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-defaultforallpages-member)|Общий колонтитул, используемый для всех страниц, если не указан колонтитул четных и нечетных страниц или первой страницы.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-evenpages-member)|Колонтитул для четных страниц, для нечетных страниц нужно указывать отдельный колонтитул.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-firstpage-member)|Колонтитул первой страницы, для остальных страниц используется общий или четный и нечетный колонтитулы.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-oddpages-member)|Колонтитул для нечетных страниц, для четных страниц нужно указывать отдельный колонтитул.|
||[state](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-state-member)|Состояние, в котором задаются заглавные и пешеходные дорожки.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetmargins-member)|Получает или задает отметку, которая указывает, выровнены ли колонтитулы относительно полей страницы, установленных в параметрах макета страницы для листа.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetscale-member)|Получает или задает отметку, которая указывает, нужно ли масштабировать колонтитулы с помощью процентных значений, установленных в параметрах макета страницы для листа.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#excel-excel-image-format-member)|Возвращает формат изображения.|
||[id](/javascript/api/excel/excel.image#excel-excel-image-id-member)|Указывает идентификатор формы для объекта изображения.|
||[shape](/javascript/api/excel/excel.image#excel-excel-image-shape-member)|Возвращает объект `Shape` , связанный с изображением.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-enabled-member)|Значение true, если в Excel используется итерация для разрешения циклических ссылок.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxchange-member)|Указывает максимальное количество изменений между каждой итерацией, Excel устраняет круговые ссылки.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxiteration-member)|Указывает максимальное количество итераций, Excel можно использовать для решения круговой ссылки.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadlength-member)|Представляет длину наконечника в начале указанной линии.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadstyle-member)|Представляет стиль наконечника в начале указанной линии.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadwidth-member)|Представляет ширину наконечника в начале указанной линии.|
||[beginConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedshape-member)|Представляет фигуру, к которой привязано начало указанной линии.|
||[beginConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedsite-member)|Представляет точку соединения, к которой привязано начало соединительной линии.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectbeginshape-member(1))|Привязывает начало указанного соединителя к указанной фигуре.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectendshape-member(1))|Привязывает конец указанного соединителя к указанной фигуре.|
||[connectorType](/javascript/api/excel/excel.line#excel-excel-line-connectortype-member)|Представляет тип соединительной линии.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectbeginshape-member(1))|Отвязывает начало указанного соединителя от фигуры.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectendshape-member(1))|Отвязывает конец указанного соединителя от фигуры.|
||[endArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadlength-member)|Представляет длину наконечника в конце указанной линии.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadstyle-member)|Представляет стиль наконечника в конце указанной линии.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadwidth-member)|Представляет ширину наконечника в конце указанной линии.|
||[endConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-endconnectedshape-member)|Представляет фигуру, к которой привязан конец указанной линии.|
||[endConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-endconnectedsite-member)|Представляет точку соединения, к которой привязан конец соединительной линии.|
||[id](/javascript/api/excel/excel.line#excel-excel-line-id-member)|Указывает идентификатор формы.|
||[isBeginConnected](/javascript/api/excel/excel.line#excel-excel-line-isbeginconnected-member)|Указывает, подключено ли начало указанной строки к фигуре.|
||[isEndConnected](/javascript/api/excel/excel.line#excel-excel-line-isendconnected-member)|Указывает, подключен ли конец указанной строки к фигуре.|
||[shape](/javascript/api/excel/excel.line#excel-excel-line-shape-member)|Возвращает объект `Shape` , связанный с строкой.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[columnIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-columnindex-member)|Указывает индекс столбца для разрыва страницы.|
||[delete()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-delete-member(1))|Удаляет объект разрыва страницы.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-getcellafterbreak-member(1))|Получает первую ячейку после разрыва страницы.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-rowindex-member)|Указывает индекс строки для разрыва страницы.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-add-member(1))|Добавляет разрыв страницы перед левой верхней ячейкой указанного диапазона.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getcount-member(1))|Получает количество разрывов страниц в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getitem-member(1))|Получает объект разрыва страницы по индексу.|
||[items](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-removepagebreaks-member(1))|Сбрасывает все добавленные вручную разрывы страниц в коллекции.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-blackandwhite-member)|Параметр черной и белой печати таблицы.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-bottommargin-member)|Поля нижней страницы таблицы, которые можно использовать для печати в точках.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centerhorizontally-member)|Центр таблицы горизонтально флаг.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centervertically-member)|Центр таблицы вертикально флаг.|
||[draftMode](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-draftmode-member)|Вариант режима черновика таблицы.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-firstpagenumber-member)|Номер первой страницы таблицы для печати.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-footermargin-member)|Поле для подножки таблицы в точках для использования при печати.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintarea-member(1))|Получает объект `RangeAreas` , состоящий из одного или нескольких прямоугольных диапазонов, который представляет область печати для таблицы.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintareaornullobject-member(1))|Получает объект `RangeAreas` , состоящий из одного или нескольких прямоугольных диапазонов, который представляет область печати для таблицы.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumns-member(1))|Получает объект range, представляющий столбцы заголовков.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumnsornullobject-member(1))|Получает объект range, представляющий столбцы заголовков.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerows-member(1))|Получает объект range, представляющий строки заголовков.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerowsornullobject-member(1))|Получает объект range, представляющий строки заголовков.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headermargin-member)|Поле заглавной таблицы в точках для использования при печати.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headersfooters-member)|Настройка колонтитулов для листа.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-leftmargin-member)|Левая маржа таблицы в точках для использования при печати.|
||[orientation](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-orientation-member)|Ориентация таблицы страницы.|
||[paperSize](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-papersize-member)|Размер бумаги листа страницы.|
||[printComments](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printcomments-member)|Указывает, должны ли при печати отображаться комментарии таблицы.|
||[printErrors](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printerrors-member)|Параметр ошибки печати таблицы.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printgridlines-member)|Указывает, будут ли напечатаны сетки таблицы.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printheadings-member)|Указывает, будут ли напечатаны заголовки таблицы.|
||[printOrder](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printorder-member)|Параметр распечатать страницы лист.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-rightmargin-member)|Правое поле таблицы в точках для использования при печати.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintarea-member(1))|Задает область печати листа.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintmargins-member(1))|Задает поля страницы с единицами измерения для листа.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlecolumns-member(1))|Задает столбцы, содержащие ячейки, которые должны повторяться слева на каждой странице при печати листа.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlerows-member(1))|Задает строки, содержащие ячейки, которые должны повторяться сверху каждой страницы при печати листа.|
||[topMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-topmargin-member)|Верхняя маржа таблицы в точках для использования при печати.|
||[zoom](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-zoom-member)|Параметры масштабирования печати таблицы.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-bottom-member)|Указывает нижнюю маржу макета страницы в единице, указанной для печати.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-footer-member)|Указывает поле для подножки макета страницы в единице, указанной для печати.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-header-member)|Указывает маржу загона макета страницы в единице, указанной для печати.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-left-member)|Указывает левое поле макета страницы в единице, указанной для печати.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-right-member)|Указывает правую маржу макета страницы в единице, указанной для печати.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-top-member)|Указывает верхнюю маржу макета страницы в единице, указанной для печати.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-horizontalfittopages-member)|Количество страниц, размещаемых по горизонтали.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-scale-member)|Значение масштаба печатной страницы может быть равным от 10 до 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-verticalfittopages-member)|Количество страниц, размещаемых по вертикали.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbyvalues-member(1))|Сортирует сводную таблицу по указанным значениям в определенной области.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-autoformat-member)|Указывает, будет ли форматирование автоматически отформатировано при обновлении или при перемещении полей.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatahierarchy-member(1))|Получает объект DataHierarchy, использующийся для вычисления значения в указанном диапазоне сводной таблицы.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getpivotitems-member(1))|Получает объекты PivotItem с оси, образующие значение в указанном диапазоне сводной таблицы.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-preserveformatting-member)|Указывает, сохраняется ли форматирование при обновлении или пересчете отчета с помощью операций, таких как развязка, сортировка или изменение элементов поля страниц.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setautosortoncell-member(1))|Задает для сводной таблицы автоматическую сортировку, используя указанную ячейку, чтобы автоматически выбрать все необходимые условия и контекст.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-enabledatavalueediting-member)|Указывает, разрешается ли пользователю изменять значения в теле данных.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-usecustomsortlists-member)|Указывает, использует ли pivotTable настраиваемые списки при сортировке.|
|[Range](/javascript/api/excel/excel.range)|[autoFill (destinationRange?: Range \| string, autoFillType?: Excel. AutoFillType)](/javascript/api/excel/excel.range#excel-excel-range-autofill-member(1))|Заполняет диапазон от текущего диапазона до диапазона назначения с помощью указанной логики AutoFill.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#excel-excel-range-convertdatatypetotext-member(1))|Преобразует ячейки диапазона с типами данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#excel-excel-range-converttolinkeddatatype-member(1))|Преобразует ячейки диапазона в связанные типы данных в таблице.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1))|Копирует данные ячейки или форматирование из диапазона исходных данных или `RangeAreas` текущего диапазона.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-find-member(1))|Находит определенную строку на основе указанных условий.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-findornullobject-member(1))|Находит определенную строку на основе указанных условий.|
||[flashFill()](/javascript/api/excel/excel.range#excel-excel-range-flashfill-member(1))|Делает флэш-заполнение для текущего диапазона.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcellproperties-member(1))|Возвращает двумерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой ячейки.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcolumnproperties-member(1))|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждого столбца.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getrowproperties-member(1))|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой строки.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1))|Получает объект `RangeAreas` , состоящий из одного или нескольких прямоугольных диапазонов, который представляет все ячейки, которые соответствуют указанному типу и значению.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1))|Получает объект `RangeAreas` , состоящий из одного или нескольких диапазонов, который представляет все ячейки, которые соответствуют указанному типу и значению.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-gettables-member(1))|Получает коллекцию таблиц с заданной областью, перекрывающую диапазон.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#excel-excel-range-linkeddatatypestate-member)|Представляет состояние типа данных каждой ячейки.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))|Удаляет повторяющиеся значения из диапазона, заданного столбцами.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#excel-excel-range-replaceall-member(1))|Находит и заменяет определенную строку на основе условий, указанных в текущем диапазоне.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#excel-excel-range-setcellproperties-member(1))|Обновляет диапазон на основе 2D-массива свойств ячейки, инкапсулируя такие вещи, как шрифт, заливка, границы и выравнивание.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setcolumnproperties-member(1))|Обновляет диапазон на основе одномерного массива свойств столбцов, инкапсулируя такие вещи, как шрифт, заливка, границы и выравнивание.|
||[setDirty()](/javascript/api/excel/excel.range#excel-excel-range-setdirty-member(1))|Устанавливает диапазон, предназначенный для пересчета при выполнении следующего пересчета.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setrowproperties-member(1))|Обновляет диапазон на основе одномерного массива свойств строки, инкапсулируя такие вещи, как шрифт, заливка, границы и выравнивание.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[address](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-address-member)|Возвращает ссылку `RangeAreas` в стиле A1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-addresslocal-member)|Возвращает ссылку `RangeAreas` в локале пользователя.|
||[areaCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areacount-member)|Возвращает количество прямоугольных диапазонов, составляющих этот `RangeAreas` объект.|
||[areas](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areas-member)|Возвращает коллекцию прямоугольных диапазонов, которые составляют этот `RangeAreas` объект.|
||[calculate()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-calculate-member(1))|Вычисляет все ячейки в `RangeAreas`.|
||[cellCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-cellcount-member)|Возвращает количество ячеек `RangeAreas` в объекте, суммирует количество ячеек всех отдельных прямоугольных диапазонов.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-clear-member(1))|Очищает значения, формат, заполнение, границу и другие свойства в каждом из областей, в которых состоит этот `RangeAreas` объект.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-conditionalformats-member)|Возвращает коллекцию условных форматов, которые пересекаются с любыми ячейками в этом объекте `RangeAreas` .|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-convertdatatypetotext-member(1))|Преобразует все ячейки в типах `RangeAreas` данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-converttolinkeddatatype-member(1))|Преобразует все ячейки в связанные `RangeAreas` типы данных.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-copyfrom-member(1))|Копирует данные ячейки или форматирование из диапазона исходных данных или `RangeAreas` текущего `RangeAreas`.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-datavalidation-member)|Возвращает объект проверки данных для всех диапазонов в `RangeAreas`.|
||[format](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-format-member)|Возвращает объект `RangeFormat` , инкапсулируя шрифт, заполнять, границы, выравнивание и другие свойства для всех диапазонов объекта `RangeAreas` .|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirecolumn-member(1))|`RangeAreas` `RangeAreas` Возвращает объект, который представляет целые столбцы (например, `RangeAreas` если ток представляет ячейки "B4:E11, H2", `RangeAreas` он возвращает столбцы "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirerow-member(1))|`RangeAreas` `RangeAreas` Возвращает объект, который представляет целые строки (например, `RangeAreas` если ток представляет ячейки "B4:E11", `RangeAreas` он возвращает строки "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersection-member(1))|Возвращает объект `RangeAreas` , который представляет пересечение заданных диапазонов или `RangeAreas`.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersectionornullobject-member(1))|Возвращает объект `RangeAreas` , который представляет пересечение заданных диапазонов или `RangeAreas`.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getoffsetrangeareas-member(1))|Возвращает объект, `RangeAreas` смещенный определенной строкой и смещением столбца.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcells-member(1))|Возвращает объект, `RangeAreas` который представляет все ячейки, которые соответствуют указанному типу и значению.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcellsornullobject-member(1))|Возвращает объект, `RangeAreas` который представляет все ячейки, которые соответствуют указанному типу и значению.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-gettables-member(1))|Возвращает объемную коллекцию таблиц, которые перекрываются с любым диапазоном в этом объекте `RangeAreas` .|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareas-member(1))|Возвращает используемое, `RangeAreas` которое включает все используемые области отдельных прямоугольных диапазонов объекта `RangeAreas` .|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareasornullobject-member(1))|Возвращает используемое, `RangeAreas` которое включает все используемые области отдельных прямоугольных диапазонов объекта `RangeAreas` .|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirecolumn-member)|Указывает, представляют ли `RangeAreas` все диапазоны на этом объекте целые столбцы (например, "A:C, Q:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirerow-member)|Указывает, представляют `RangeAreas` ли все диапазоны на этом объекте целые строки (например, "1:3, 5:7").|
||[setDirty()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-setdirty-member(1))|Задает перерасчет `RangeAreas` при следующем пересчете.|
||[style](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-style-member)|Представляет стиль для всех диапазонов в этом объекте `RangeAreas` .|
||[worksheet](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-worksheet-member)|Возвращает таблицу для текущего `RangeAreas`.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-tintandshade-member)|Указывает двойной, который осветляет или темнеет цвет для границы диапазона, значение между -1 (самый темный) и 1 (самый яркий), с 0 для исходного цвета.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-tintandshade-member)|Указывает двойник, который осветляет или темнеет цвет для границ диапазона.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getcount-member(1))|Возвращает количество диапазонов в `RangeCollection`.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getitemat-member(1))|Возвращает объект диапазона в зависимости от его положения в `RangeCollection`.|
||[items](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-pattern-member)|Шаблон диапазона.|
||[patternColor](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterncolor-member)|Цветовой код HTML, представляющий цвет шаблона диапазона, в форме #RRGGBB (например, "FFA500"), или в виде имени HTML-цвета (например, "оранжевый").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterntintandshade-member)|Указывает двойной номер, который осветляет или темнеет цвет шаблона для заполнения диапазона.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-tintandshade-member)|Указывает двойной, который осветляет или затемнеет цвет для заполнения диапазона.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-strikethrough-member)|Указывает состояние забастовки шрифта.|
||[subscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-subscript-member)|Указывает состояние подписки шрифта.|
||[superscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-superscript-member)|Указывает состояние суперсписка шрифта.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-tintandshade-member)|Указывает двойную, которая осветляет или темнеет цвет шрифта диапазона.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-autoindent-member)|Указывает, будет ли текст автоматически отступным, если выравнивание текста задано для равного распространения.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-indentlevel-member)|Целое число от 0 до 250, указывающее уровень отступа.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-readingorder-member)|Направление чтения для диапазона.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-shrinktofit-member)|Указывает, если текст автоматически сокращается, чтобы соответствовать ширине доступных столбцов.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-removed-member)|Количество повторяющихся строк, удаленных операцией.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-uniqueremaining-member)|Количество оставшихся уникальных строк, присутствующих в получившемся диапазоне.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-completematch-member)|Указывает, должен ли совпадение быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-matchcase-member)|Указывает, является ли совпадение чувствительным к делу.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-address-member)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-addresslocal-member)|Представляет свойство `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-rowindex-member)|Представляет свойство `rowIndex`.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-completematch-member)|Указывает, должен ли совпадение быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-matchcase-member)|Указывает, является ли совпадение чувствительным к делу.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-searchdirection-member)|Указывает направление поиска.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-format-member)|Представляет свойство `format`.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-hyperlink-member)|Представляет свойство `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-style-member)|Представляет свойство `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnhidden-member)|Представляет свойство `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnwidth-member)||
||[формат: Excel. CellPropertiesFormat & { columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-format-member)|Представляет свойство `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[формат: Excel. CellPropertiesFormat & { rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-format-member)|Представляет свойство `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowheight-member)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowhidden-member)|Представляет свойство `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#excel-excel-shape-alttextdescription-member)|Указывает альтернативный текст описания объекта `Shape` .|
||[altTextTitle](/javascript/api/excel/excel.shape#excel-excel-shape-alttexttitle-member)|Указывает альтернативный текст заголовка для `Shape` объекта.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#excel-excel-shape-connectionsitecount-member)|Возвращает количество точек соединения на фигуре.|
||[delete()](/javascript/api/excel/excel.shape#excel-excel-shape-delete-member(1))|Удаляет фигуру с листа.|
||[fill](/javascript/api/excel/excel.shape#excel-excel-shape-fill-member)|Возвращает формат заливки фигуры.|
||[geometricShape](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshape-member)|Возвращает геометрическую фигуру, связанную с линией.|
||[geometricShapeType](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshapetype-member)|Указывает тип геометрической фигуры этой геометрической фигуры.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1))|Преобразует фигуру в изображение и возвращает изображение в виде строки в кодировке base64.|
||[group](/javascript/api/excel/excel.shape#excel-excel-shape-group-member)|Возвращает группу фигур, связанную с фигурой.|
||[height](/javascript/api/excel/excel.shape#excel-excel-shape-height-member)|Указывает высоту фигуры в точках.|
||[id](/javascript/api/excel/excel.shape#excel-excel-shape-id-member)|Указывает идентификатор формы.|
||[image](/javascript/api/excel/excel.shape#excel-excel-shape-image-member)|Возвращает изображение, связанное с фигурой.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementleft-member(1))|Перемещает фигуру по горизонтали на указанное число пунктов.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementrotation-member(1))|Поворачивает фигуру по часовой стрелке относительно оси Z на указанное число градусов.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementtop-member(1))|Перемещает фигуру по вертикали на указанное число пунктов.|
||[left](/javascript/api/excel/excel.shape#excel-excel-shape-left-member)|Расстояние в пунктах от левого края фигуры до левого края листа.|
||[level](/javascript/api/excel/excel.shape#excel-excel-shape-level-member)|Указывает уровень указанной формы.|
||[line](/javascript/api/excel/excel.shape#excel-excel-shape-line-member)|Возвращает линию, связанную с фигурой.|
||[lineFormat](/javascript/api/excel/excel.shape#excel-excel-shape-lineformat-member)|Возвращает формат линии для фигуры.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#excel-excel-shape-lockaspectratio-member)|Указывает, заблокировано ли соотношение аспектов этой фигуры.|
||[name](/javascript/api/excel/excel.shape#excel-excel-shape-name-member)|Указывает имя фигуры.|
||[onActivated](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)|Возникает, если фигура активирована.|
||[onDeactivated](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)|Возникает, если фигура деактивирована.|
||[parentGroup](/javascript/api/excel/excel.shape#excel-excel-shape-parentgroup-member)|Указывает родительную группу этой фигуры.|
||[rotation](/javascript/api/excel/excel.shape#excel-excel-shape-rotation-member)|Указывает вращение фигуры в градусах.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scaleheight-member(1))|Масштабирует высоту фигуры с применением указанного коэффициента.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scalewidth-member(1))|Масштабирует ширину фигуры с применением указанного коэффициента.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#excel-excel-shape-setzorder-member(1))|Перемещает указанную фигуру вверх или вниз по оси Z в коллекции, что переносит ее вперед или назад относительно других фигур.|
||[textFrame](/javascript/api/excel/excel.shape#excel-excel-shape-textframe-member)|Возвращает объект рамки с текстом для фигуры.|
||[top](/javascript/api/excel/excel.shape#excel-excel-shape-top-member)|Расстояние в пунктах от верхнего края фигуры до верхнего края листа.|
||[type](/javascript/api/excel/excel.shape#excel-excel-shape-type-member)|Возвращает тип фигуры.|
||[visible](/javascript/api/excel/excel.shape#excel-excel-shape-visible-member)|Указывает, видна ли фигура.|
||[width](/javascript/api/excel/excel.shape#excel-excel-shape-width-member)|Указывает ширину в точках формы.|
||[zOrderPosition](/javascript/api/excel/excel.shape#excel-excel-shape-zorderposition-member)|Возвращает положение указанной фигуры по оси Z. Значение 0 представляет нижнее положение по оси.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-shapeid-member)|Получает ID активированной фигуры.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-worksheetid-member)|Получает ID таблицы, в которой активируется фигура.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1))|Добавляет геометрическую фигуру на лист.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgroup-member(1))|Группирует подмножество фигур на листе этой коллекции.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1))|Создает изображение из строки в кодировке base64 и добавляет его на лист.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1))|Добавляет линию на лист.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1))|Добавляет текстовое поле на лист с указанным текстом в качестве содержимого.|
||[getCount()](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getcount-member(1))|Возвращает количество фигур на листе.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitem-member(1))|Получает фигуру с ее именем или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemat-member(1))|Получает фигуру с помощью ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-shapeid-member)|Получает ID деактивированной фигуры.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-worksheetid-member)|Получает ID таблицы, в которой фигура деактивирована.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-clear-member(1))|Очищает формат заливки фигуры.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-foregroundcolor-member)|Представляет цвет переднего плана заполнения фигуры в формате HTML-цвета в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый")|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-setsolidcolor-member(1))|Задает заливку одним цветом для фигуры.|
||[transparency](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-transparency-member)|Указывает процент прозрачности заполнения как значение от 0.0 (непрозрачная) до 1.0 (clear).|
||[type](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-type-member)|Возвращает тип заливки фигуры.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-bold-member)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-color-member)|Представление цветового кода HTML текстового цвета (например, "#FF0000" представляет красный цвет).|
||[italic](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-italic-member)|Указывает, применяется ли курсив.|
||[name](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-name-member)|Представляет имя шрифта (например, "Калибри").|
||[size](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-size-member)|Представляет размер шрифта в точках (например, 11).|
||[underline](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-underline-member)|Тип подчеркивания, применяемый для шрифта.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-id-member)|Указывает идентификатор формы.|
||[shape](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shape-member)|Возвращает объект `Shape` , связанный с группой.|
||[shapes](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shapes-member)|Возвращает коллекцию объектов `Shape` .|
||[ungroup()](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-ungroup-member(1))|Отменяет группировку любых сгруппированных фигур в указанной группе фигур.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-color-member)|Представляет цвет строки в формате HTML-цвета в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-dashstyle-member)|Представляет тип линии фигуры.|
||[style](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-style-member)|Представляет тип линии фигуры.|
||[transparency](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-transparency-member)|Представляет степень прозрачности указанной линии как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная).|
||[visible](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-visible-member)|Указывает, отображается ли форматирование строки элемента фигуры.|
||[weight](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-weight-member)|Представляет толщину линии (в пунктах).|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#excel-excel-sortfield-subfield-member)|Указывает подполе, которое является целевым именем свойства для сортировки с богатым значением.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getcount-member(1))|Получает количество стилей в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemat-member(1))|Получает стиль на основе его позиции в коллекции.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#excel-excel-table-autofilter-member)|Представляет объект `AutoFilter` таблицы.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-source-member)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-tableid-member)|Получает ID добавленной таблицы.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-worksheetid-member)|Получает ID таблицы, в которую добавляется таблица.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-details-member)|Получает сведения о деталях изменений.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member)|Возникает при добавлении новой таблицы в книгу.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member)|Возникает, если указанная таблица удалена из книги.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-source-member)|Получает источник события.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tableid-member)|Получает удаленный ID таблицы.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tablename-member)|Получает имя удаляемой таблицы.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-worksheetid-member)|Получает ID таблицы, в которой удаляется таблица.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getcount-member(1))|Получает количество таблиц в коллекции.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getfirst-member(1))|Получает первую таблицу в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitem-member(1))|Получает таблицу по имени или ИД.|
||[items](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#excel-excel-textframe-autosizesetting-member)|Автоматические параметры размеров для текстового кадра.|
||[bottomMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-bottommargin-member)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[deleteText()](/javascript/api/excel/excel.textframe#excel-excel-textframe-deletetext-member(1))|Удаляет весь текст в рамке с текстом.|
||[hasText](/javascript/api/excel/excel.textframe#excel-excel-textframe-hastext-member)|Указывает, содержит ли текстовый кадр текст.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontalalignment-member)|Представляет горизонтальное выравнивание рамки с текстом.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontaloverflow-member)|Представляет действие горизонтального переполнения рамки с текстом.|
||[leftMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-leftmargin-member)|Представляет левое поле рамки с текстом (в пунктах).|
||[orientation](/javascript/api/excel/excel.textframe#excel-excel-textframe-orientation-member)|Представляет угол, на который ориентирован текст для текстового кадра.|
||[readingOrder](/javascript/api/excel/excel.textframe#excel-excel-textframe-readingorder-member)|Представляет направление чтения рамки с текстом (слева направо или справа налево).|
||[rightMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-rightmargin-member)|Представляет правое поле рамки с текстом (в пунктах).|
||[textRange](/javascript/api/excel/excel.textframe#excel-excel-textframe-textrange-member)|Представляет текст, присоединенный к фигуре в текстовой рамке, а также свойства и методы для операций с текстом.|
||[topMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-topmargin-member)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticalalignment-member)|Представляет вертикальное выравнивание для рамки с текстом.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticaloverflow-member)|Представляет действие вертикального переполнения рамки с текстом.|
|[TextRange](/javascript/api/excel/excel.textrange)|[font](/javascript/api/excel/excel.textrange#excel-excel-textrange-font-member)|Возвращает объект `ShapeFont` , который представляет атрибуты шрифта для диапазона текста.|
||[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#excel-excel-textrange-getsubstring-member(1))|Возвращает объект TextRange для подстроки в указанном диапазоне.|
||[text](/javascript/api/excel/excel.textrange#excel-excel-textrange-text-member)|Представляет содержимое с обычным текстом в диапазоне текста.|
|[Workbook](/javascript/api/excel/excel.workbook)|[autoSave](/javascript/api/excel/excel.workbook#excel-excel-workbook-autosave-member)|Указывает, находится ли книга в режиме AutoSave.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#excel-excel-workbook-calculationengineversion-member)|Возвращает номер версии модуля вычислений Excel.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbook#excel-excel-workbook-chartdatapointtrack-member)|Значение true, если все диаграммы в книге отслеживают точки фактических данных, с которыми они связаны.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechart-member(1))|Получает текущую активную диаграмму в книге.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechartornullobject-member(1))|Получает текущую активную диаграмму в книге.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getisactivecollabsession-member(1))|Возвращается `true` , если книга редактирована несколькими пользователями (с помощью соавторов).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedranges-member(1))|Получает текущий выделенный диапазон (один или несколько) в книге.|
||[isDirty](/javascript/api/excel/excel.workbook#excel-excel-workbook-isdirty-member)|Указывает, были ли внесены изменения с момента последнего сберегаемого книги.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member)|Возникает при смене параметра AutoSave в книге.|
||[previouslySaved](/javascript/api/excel/excel.workbook#excel-excel-workbook-previouslysaved-member)|Указывает, была ли книга сохранена локально или в Интернете.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#excel-excel-workbook-useprecisionasdisplayed-member)|Значение true, если вычисления в книге выполняются только с той точностью чисел, с которой они отображаются.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#excel-excel-workbookautosavesettingchangedeventargs-type-member)|Получает тип события.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[autoFilter](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-autofilter-member)|Представляет объект `AutoFilter` таблицы.|
||[enableCalculation](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-enablecalculation-member)|Определяет, следует ли Excel при необходимости пересчитать таблицу.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findall-member(1))|Находит все вхождения `RangeAreas` данной строки на основе указанных критериев и возвращает их как объект, состоящий из одного или нескольких прямоугольных диапазонов.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findallornullobject-member(1))|Находит все вхождения `RangeAreas` данной строки на основе указанных критериев и возвращает их как объект, состоящий из одного или нескольких прямоугольных диапазонов.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1))|Получает объект `RangeAreas` , представляющий один или несколько блоков прямоугольных диапазонов, указанных по адресу или имени.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-horizontalpagebreaks-member)|Получает коллекцию горизонтальных разрывов страницы для листа.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)|Возникает, если изменен формат указанного листа.|
||[pageLayout](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pagelayout-member)|Получает объект `PageLayout` таблицы.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-replaceall-member(1))|Находит и заменяет определенную строку на основе условий, указанных в текущем листе.|
||[shapes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-shapes-member)|Возвращает коллекцию всех объектов Shape на листе.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-verticalpagebreaks-member)|Получает коллекцию вертикальных разрывов страницы для листа.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-details-member)|Представляет сведения об изменениях.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member)|Возникает при изменении любого листа в книге.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member)|Возникает при смене формата любого таблицы в книге.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member)|Возникает при изменениях выделения на любом листе.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-address-member)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrange-member(1))|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrangeornullobject-member(1))|Получает диапазон, представляющий измененную область конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-worksheetid-member)|Получает ID таблицы, в которой изменились данные.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-completematch-member)|Указывает, должен ли совпадение быть полным или частичным.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-matchcase-member)|Указывает, является ли совпадение чувствительным к делу.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

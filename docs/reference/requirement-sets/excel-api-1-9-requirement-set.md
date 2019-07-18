---
title: Набор обязательных элементов API JavaScript для Excel 1,9
description: Сведения о наборе требований ExcelApi 1,9
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1c7361debe7ba09c3477d39d9337c35bf5df3066
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772004"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Новые возможности API JavaScript для Excel 1.9

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

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Возвращает версию модуля вычислений Excel, использованного для последнего полного пересчета. Только для чтения.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Возвращает состояние вычисления приложения. Дополнительные сведения см. в статье Excel.CalculationState. Только для чтения.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Возвращает параметры итеративных вычислений.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Приостанавливает обновление экрана до вызова следующего метода context.sync().|
|[Аппликатиондата](/javascript/api/excel/excel.applicationdata)|[calculationEngineVersion](/javascript/api/excel/excel.applicationdata#calculationengineversion)|Возвращает версию модуля вычислений Excel, использованного для последнего полного пересчета. Только для чтения.|
||[calculationState](/javascript/api/excel/excel.applicationdata#calculationstate)|Возвращает состояние вычисления приложения. Дополнительные сведения см. в статье Excel.CalculationState. Только для чтения.|
||[iterativeCalculation](/javascript/api/excel/excel.applicationdata#iterativecalculation)|Возвращает параметры итеративных вычислений.|
|[Аппликатионлоадоптионс](/javascript/api/excel/excel.applicationloadoptions)|[calculationEngineVersion](/javascript/api/excel/excel.applicationloadoptions#calculationengineversion)|Возвращает версию модуля вычислений Excel, использованного для последнего полного пересчета. Только для чтения.|
||[calculationState](/javascript/api/excel/excel.applicationloadoptions#calculationstate)|Возвращает состояние вычисления приложения. Дополнительные сведения см. в статье Excel.CalculationState. Только для чтения.|
||[iterativeCalculation](/javascript/api/excel/excel.applicationloadoptions#iterativecalculation)|Возвращает параметры итеративных вычислений.|
|[Аппликатионупдатедата](/javascript/api/excel/excel.applicationupdatedata)|[iterativeCalculation](/javascript/api/excel/excel.applicationupdatedata#iterativecalculation)|Возвращает параметры итеративных вычислений.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Применяет автофильтр к диапазону. При этом фильтруется столбец, если указаны индекс столбца и условия фильтрации.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Очищает условия фильтрации автофильтра.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Возвращает объект Range, представляющий диапазон, к которому применяется автофильтр.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Возвращает объект Range, представляющий диапазон, к которому применяется автофильтр.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Массив, содержащий все условия фильтрации в диапазоне с примененным автофильтром. Только для чтения.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Указывает, включен ли автофильтр. Только для чтения.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Указывает, есть ли в автофильтре условия фильтрации. Только для чтения.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Применяет указанный объект Autofilter, находящийся в настоящее время в диапазоне.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Удаляет автофильтр из диапазона.|
|[Аутофилтердата](/javascript/api/excel/excel.autofilterdata)|[criteria](/javascript/api/excel/excel.autofilterdata#criteria)|Массив, содержащий все условия фильтрации в диапазоне с примененным автофильтром. Только для чтения.|
||[enabled](/javascript/api/excel/excel.autofilterdata#enabled)|Указывает, включен ли автофильтр. Только для чтения.|
||[isDataFiltered](/javascript/api/excel/excel.autofilterdata#isdatafiltered)|Указывает, есть ли в автофильтре условия фильтрации. Только для чтения.|
|[Аутофилтерлоадоптионс](/javascript/api/excel/excel.autofilterloadoptions)|[$all](/javascript/api/excel/excel.autofilterloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.autofilterloadoptions#criteria)|Массив, содержащий все условия фильтрации в диапазоне с примененным автофильтром. Только для чтения.|
||[enabled](/javascript/api/excel/excel.autofilterloadoptions#enabled)|Указывает, включен ли автофильтр. Только для чтения.|
||[isDataFiltered](/javascript/api/excel/excel.autofilterloadoptions#isdatafiltered)|Указывает, есть ли в автофильтре условия фильтрации. Только для чтения.|
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
|[Целлпропертиесбордерлоадоптионс](/javascript/api/excel/excel.cellpropertiesborderloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesborderloadoptions#color)|Указывает, следует ли загружать `color` свойство.|
||[style](/javascript/api/excel/excel.cellpropertiesborderloadoptions#style)|Указывает, следует ли загружать `style` свойство.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesborderloadoptions#tintandshade)|Указывает, следует ли загружать `tintAndShade` свойство.|
||[weight](/javascript/api/excel/excel.cellpropertiesborderloadoptions#weight)|Указывает, следует ли загружать `weight` свойство.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Представляет свойство `format.fill.color`.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Представляет свойство `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Представляет свойство `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Представляет свойство `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Представляет свойство `format.fill.tintAndShade`.|
|[Целлпропертиесфилллоадоптионс](/javascript/api/excel/excel.cellpropertiesfillloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesfillloadoptions#color)|Указывает, следует ли загружать `color` свойство.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfillloadoptions#pattern)|Указывает, следует ли загружать `pattern` свойство.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterncolor)|Указывает, следует ли загружать `patternColor` свойство.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterntintandshade)|Указывает, следует ли загружать `patternTintAndShade` свойство.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#tintandshade)|Указывает, следует ли загружать `tintAndShade` свойство.|
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
|[Целлпропертиесфонтлоадоптионс](/javascript/api/excel/excel.cellpropertiesfontloadoptions)|[bold](/javascript/api/excel/excel.cellpropertiesfontloadoptions#bold)|Указывает, следует ли загружать `bold` свойство.|
||[color](/javascript/api/excel/excel.cellpropertiesfontloadoptions#color)|Указывает, следует ли загружать `color` свойство.|
||[italic](/javascript/api/excel/excel.cellpropertiesfontloadoptions#italic)|Указывает, следует ли загружать `italic` свойство.|
||[name](/javascript/api/excel/excel.cellpropertiesfontloadoptions#name)|Указывает, следует ли загружать `name` свойство.|
||[size](/javascript/api/excel/excel.cellpropertiesfontloadoptions#size)|Указывает, следует ли загружать `size` свойство.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfontloadoptions#strikethrough)|Указывает, следует ли загружать `strikethrough` свойство.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#subscript)|Указывает, следует ли загружать `subscript` свойство.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#superscript)|Указывает, следует ли загружать `superscript` свойство.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfontloadoptions#tintandshade)|Указывает, следует ли загружать `tintAndShade` свойство.|
||[underline](/javascript/api/excel/excel.cellpropertiesfontloadoptions#underline)|Указывает, следует ли загружать `underline` свойство.|
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
|[Целлпропертиесформатлоадоптионс](/javascript/api/excel/excel.cellpropertiesformatloadoptions)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformatloadoptions#autoindent)|Указывает, следует ли загружать `autoIndent` свойство.|
||[borders](/javascript/api/excel/excel.cellpropertiesformatloadoptions#borders)|Указывает, следует ли загружать `borders` свойство.|
||[fill](/javascript/api/excel/excel.cellpropertiesformatloadoptions#fill)|Указывает, следует ли загружать `fill` свойство.|
||[font](/javascript/api/excel/excel.cellpropertiesformatloadoptions#font)|Указывает, следует ли загружать `font` свойство.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#horizontalalignment)|Указывает, следует ли загружать `horizontalAlignment` свойство.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformatloadoptions#indentlevel)|Указывает, следует ли загружать `indentLevel` свойство.|
||[protection](/javascript/api/excel/excel.cellpropertiesformatloadoptions#protection)|Указывает, следует ли загружать `protection` свойство.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformatloadoptions#readingorder)|Указывает, следует ли загружать `readingOrder` свойство.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformatloadoptions#shrinktofit)|Указывает, следует ли загружать `shrinkToFit` свойство.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformatloadoptions#textorientation)|Указывает, следует ли загружать `textOrientation` свойство.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardheight)|Указывает, следует ли загружать `useStandardHeight` свойство.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardwidth)|Указывает, следует ли загружать `useStandardWidth` свойство.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#verticalalignment)|Указывает, следует ли загружать `verticalAlignment` свойство.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformatloadoptions#wraptext)|Указывает, следует ли загружать `wrapText` свойство.|
|[Целлпропертиеслоадоптионс](/javascript/api/excel/excel.cellpropertiesloadoptions)|[address](/javascript/api/excel/excel.cellpropertiesloadoptions#address)|Указывает, следует ли загружать `address` свойство.|
||[addressLocal](/javascript/api/excel/excel.cellpropertiesloadoptions#addresslocal)|Указывает, следует ли загружать `addressLocal` свойство.|
||[format](/javascript/api/excel/excel.cellpropertiesloadoptions#format)|Указывает, следует ли загружать `format` свойство.|
||[hidden](/javascript/api/excel/excel.cellpropertiesloadoptions#hidden)|Указывает, следует ли загружать `hidden` свойство.|
||[hyperlink](/javascript/api/excel/excel.cellpropertiesloadoptions#hyperlink)|Указывает, следует ли загружать `hyperlink` свойство.|
||[style](/javascript/api/excel/excel.cellpropertiesloadoptions#style)|Указывает, следует ли загружать `style` свойство.|
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
|[Чартареаформатдата](/javascript/api/excel/excel.chartareaformatdata)|[colorScheme](/javascript/api/excel/excel.chartareaformatdata#colorscheme)|Возвращает или задает цветовую схему диаграммы. Для чтения и записи.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatdata#roundedcorners)|Указывает, содержит ли область диаграммы скругленные углы. Для чтения и записи.|
|[Чартареаформатлоадоптионс](/javascript/api/excel/excel.chartareaformatloadoptions)|[colorScheme](/javascript/api/excel/excel.chartareaformatloadoptions#colorscheme)|Возвращает или задает цветовую схему диаграммы. Для чтения и записи.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatloadoptions#roundedcorners)|Указывает, содержит ли область диаграммы скругленные углы. Для чтения и записи.|
|[Чартареаформатупдатедата](/javascript/api/excel/excel.chartareaformatupdatedata)|[colorScheme](/javascript/api/excel/excel.chartareaformatupdatedata#colorscheme)|Возвращает или задает цветовую схему диаграммы. Для чтения и записи.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatupdatedata#roundedcorners)|Указывает, содержит ли область диаграммы скругленные углы. Для чтения и записи.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках.|
|[Чартаксисдата](/javascript/api/excel/excel.chartaxisdata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisdata#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках.|
|[Чартаксислоадоптионс](/javascript/api/excel/excel.chartaxisloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisloadoptions#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках.|
|[Чартаксисупдатедата](/javascript/api/excel/excel.chartaxisupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisupdatedata#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Указывает, разрешен ли выход за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Указывает, разрешен ли выход за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Возвращает или задает количество интервалов в гистограмме или диаграмме Парето. Для чтения и записи.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Возвращает или задает значение выхода за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[Set (Properties: Excel. Чартбиноптионс)](/javascript/api/excel/excel.chartbinoptions#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартбиноптионсупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartbinoptions#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Возвращает или задает тип интервалов для гистограммы или диаграммы Парето. Для чтения и записи.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Возвращает или задает значение выхода за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Возвращает или задает значение ширины интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
|[Чартбиноптионсдата](/javascript/api/excel/excel.chartbinoptionsdata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsdata#allowoverflow)|Указывает, разрешен ли выход за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsdata#allowunderflow)|Указывает, разрешен ли выход за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[count](/javascript/api/excel/excel.chartbinoptionsdata#count)|Возвращает или задает количество интервалов в гистограмме или диаграмме Парето. Для чтения и записи.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsdata#overflowvalue)|Возвращает или задает значение выхода за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[type](/javascript/api/excel/excel.chartbinoptionsdata#type)|Возвращает или задает тип интервалов для гистограммы или диаграммы Парето. Для чтения и записи.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsdata#underflowvalue)|Возвращает или задает значение выхода за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[width](/javascript/api/excel/excel.chartbinoptionsdata#width)|Возвращает или задает значение ширины интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
|[Чартбиноптионслоадоптионс](/javascript/api/excel/excel.chartbinoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartbinoptionsloadoptions#$all)||
||[allowOverflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowoverflow)|Указывает, разрешен ли выход за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowunderflow)|Указывает, разрешен ли выход за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[count](/javascript/api/excel/excel.chartbinoptionsloadoptions#count)|Возвращает или задает количество интервалов в гистограмме или диаграмме Парето. Для чтения и записи.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#overflowvalue)|Возвращает или задает значение выхода за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[type](/javascript/api/excel/excel.chartbinoptionsloadoptions#type)|Возвращает или задает тип интервалов для гистограммы или диаграммы Парето. Для чтения и записи.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#underflowvalue)|Возвращает или задает значение выхода за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[width](/javascript/api/excel/excel.chartbinoptionsloadoptions#width)|Возвращает или задает значение ширины интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
|[Чартбиноптионсупдатедата](/javascript/api/excel/excel.chartbinoptionsupdatedata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowoverflow)|Указывает, разрешен ли выход за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowunderflow)|Указывает, разрешен ли выход за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[count](/javascript/api/excel/excel.chartbinoptionsupdatedata#count)|Возвращает или задает количество интервалов в гистограмме или диаграмме Парето. Для чтения и записи.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#overflowvalue)|Возвращает или задает значение выхода за верхнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[type](/javascript/api/excel/excel.chartbinoptionsupdatedata#type)|Возвращает или задает тип интервалов для гистограммы или диаграммы Парето. Для чтения и записи.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#underflowvalue)|Возвращает или задает значение выхода за нижнюю границу интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
||[width](/javascript/api/excel/excel.chartbinoptionsupdatedata#width)|Возвращает или задает значение ширины интервала в гистограмме или диаграмме Парето. Для чтения и записи.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Возвращает или задает тип вычисления квартилей для диаграммы "ящик с усами". Для чтения и записи.|
||[Set (Properties: Excel. Чартбоксвхискероптионс)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартбоксвхискероптионсупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Указывает, отображаются ли внутренние точки на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Указывает, отображается ли линия медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Указывает, отображается ли маркер медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Указывает, отображаются ли точки выбросов на диаграмме "ящик с усами". Для чтения и записи.|
|[Чартбоксвхискероптионсдата](/javascript/api/excel/excel.chartboxwhiskeroptionsdata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#quartilecalculation)|Возвращает или задает тип вычисления квартилей для диаграммы "ящик с усами". Для чтения и записи.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showinnerpoints)|Указывает, отображаются ли внутренние точки на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanline)|Указывает, отображается ли линия медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanmarker)|Указывает, отображается ли маркер медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showoutlierpoints)|Указывает, отображаются ли точки выбросов на диаграмме "ящик с усами". Для чтения и записи.|
|[Чартбоксвхискероптионслоадоптионс](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions)|[$all](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#$all)||
||[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#quartilecalculation)|Возвращает или задает тип вычисления квартилей для диаграммы "ящик с усами". Для чтения и записи.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showinnerpoints)|Указывает, отображаются ли внутренние точки на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanline)|Указывает, отображается ли линия медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanmarker)|Указывает, отображается ли маркер медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showoutlierpoints)|Указывает, отображаются ли точки выбросов на диаграмме "ящик с усами". Для чтения и записи.|
|[Чартбоксвхискероптионсупдатедата](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#quartilecalculation)|Возвращает или задает тип вычисления квартилей для диаграммы "ящик с усами". Для чтения и записи.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showinnerpoints)|Указывает, отображаются ли внутренние точки на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanline)|Указывает, отображается ли линия медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanmarker)|Указывает, отображается ли маркер медианы на диаграмме "ящик с усами". Для чтения и записи.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showoutlierpoints)|Указывает, отображаются ли точки выбросов на диаграмме "ящик с усами". Для чтения и записи.|
|[Чартколлектионлоадоптионс](/javascript/api/excel/excel.chartcollectionloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartcollectionloadoptions#pivotoptions)|Для каждого элемента в коллекции: инкапсулирует параметры для сводной диаграммы.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[pivotOptions](/javascript/api/excel/excel.chartdata#pivotoptions)|Объединяет параметры для сводной диаграммы. Только для чтения.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[Чартдаталабелдата](/javascript/api/excel/excel.chartdatalabeldata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabeldata#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[Чартдаталабеллоадоптионс](/javascript/api/excel/excel.chartdatalabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[Чартдаталабелупдатедата](/javascript/api/excel/excel.chartdatalabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках|
|[Чартдаталабелсдата](/javascript/api/excel/excel.chartdatalabelsdata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsdata#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках|
|[Чартдаталабелслоадоптионс](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках|
|[Чартдаталабелсупдатедата](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#linknumberformat)|Указывает, связан ли числовой формат с ячейками. Если указано значение True, числовой формат изменяется в подписях при его изменении в ячейках|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Указывает, содержат ли планки погрешностей точки с конечным стилем.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Указывает, какие части планок погрешностей нужно включить.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Указывает тип форматирования планок погрешностей.|
||[Set (Properties: Excel. Чартеррорбарс)](/javascript/api/excel/excel.charterrorbars#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартеррорбарсупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.charterrorbars#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|Тип диапазона, помеченного планками погрешностей.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Указывает, отображаются ли планки погрешностей.|
|[Чартеррорбарсдата](/javascript/api/excel/excel.charterrorbarsdata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsdata#endstylecap)|Указывает, содержат ли планки погрешностей точки с конечным стилем.|
||[format](/javascript/api/excel/excel.charterrorbarsdata#format)|Указывает тип форматирования планок погрешностей.|
||[include](/javascript/api/excel/excel.charterrorbarsdata#include)|Указывает, какие части планок погрешностей нужно включить.|
||[type](/javascript/api/excel/excel.charterrorbarsdata#type)|Тип диапазона, помеченного планками погрешностей.|
||[visible](/javascript/api/excel/excel.charterrorbarsdata#visible)|Указывает, отображаются ли планки погрешностей.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Представляет форматирование линий диаграммы.|
||[Set (Properties: Excel. Чартеррорбарсформат)](/javascript/api/excel/excel.charterrorbarsformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартеррорбарсформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.charterrorbarsformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартеррорбарсформатдата](/javascript/api/excel/excel.charterrorbarsformatdata)|[line](/javascript/api/excel/excel.charterrorbarsformatdata#line)|Представляет форматирование линий диаграммы.|
|[Чартеррорбарсформатлоадоптионс](/javascript/api/excel/excel.charterrorbarsformatloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charterrorbarsformatloadoptions#line)|Представляет форматирование линий диаграммы.|
|[Чартеррорбарсформатупдатедата](/javascript/api/excel/excel.charterrorbarsformatupdatedata)|[line](/javascript/api/excel/excel.charterrorbarsformatupdatedata#line)|Представляет форматирование линий диаграммы.|
|[Чартеррорбарслоадоптионс](/javascript/api/excel/excel.charterrorbarsloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsloadoptions#$all)||
||[endStyleCap](/javascript/api/excel/excel.charterrorbarsloadoptions#endstylecap)|Указывает, содержат ли планки погрешностей точки с конечным стилем.|
||[format](/javascript/api/excel/excel.charterrorbarsloadoptions#format)|Указывает тип форматирования планок погрешностей.|
||[include](/javascript/api/excel/excel.charterrorbarsloadoptions#include)|Указывает, какие части планок погрешностей нужно включить.|
||[type](/javascript/api/excel/excel.charterrorbarsloadoptions#type)|Тип диапазона, помеченного планками погрешностей.|
||[visible](/javascript/api/excel/excel.charterrorbarsloadoptions#visible)|Указывает, отображаются ли планки погрешностей.|
|[Чартеррорбарсупдатедата](/javascript/api/excel/excel.charterrorbarsupdatedata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsupdatedata#endstylecap)|Указывает, содержат ли планки погрешностей точки с конечным стилем.|
||[format](/javascript/api/excel/excel.charterrorbarsupdatedata#format)|Указывает тип форматирования планок погрешностей.|
||[include](/javascript/api/excel/excel.charterrorbarsupdatedata#include)|Указывает, какие части планок погрешностей нужно включить.|
||[type](/javascript/api/excel/excel.charterrorbarsupdatedata#type)|Тип диапазона, помеченного планками погрешностей.|
||[visible](/javascript/api/excel/excel.charterrorbarsupdatedata#visible)|Указывает, отображаются ли планки погрешностей.|
|[Чартлоадоптионс](/javascript/api/excel/excel.chartloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartloadoptions#pivotoptions)|Объединяет параметры для сводной диаграммы.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Возвращает или задает стратегию подписей карт ряда для диаграммы с картой региона. Для чтения и записи.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Возвращает или задает уровень карты ряда для диаграммы с картой региона. Для чтения и записи.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Возвращает или задает тип проекции ряда для диаграммы с картой региона. Для чтения и записи.|
||[Set (Properties: Excel. Чартмапоптионс)](/javascript/api/excel/excel.chartmapoptions#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартмапоптионсупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartmapoptions#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Чартмапоптионсдата](/javascript/api/excel/excel.chartmapoptionsdata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsdata#labelstrategy)|Возвращает или задает стратегию подписей карт ряда для диаграммы с картой региона. Для чтения и записи.|
||[level](/javascript/api/excel/excel.chartmapoptionsdata#level)|Возвращает или задает уровень карты ряда для диаграммы с картой региона. Для чтения и записи.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsdata#projectiontype)|Возвращает или задает тип проекции ряда для диаграммы с картой региона. Для чтения и записи.|
|[Чартмапоптионслоадоптионс](/javascript/api/excel/excel.chartmapoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartmapoptionsloadoptions#$all)||
||[labelStrategy](/javascript/api/excel/excel.chartmapoptionsloadoptions#labelstrategy)|Возвращает или задает стратегию подписей карт ряда для диаграммы с картой региона. Для чтения и записи.|
||[level](/javascript/api/excel/excel.chartmapoptionsloadoptions#level)|Возвращает или задает уровень карты ряда для диаграммы с картой региона. Для чтения и записи.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsloadoptions#projectiontype)|Возвращает или задает тип проекции ряда для диаграммы с картой региона. Для чтения и записи.|
|[Чартмапоптионсупдатедата](/javascript/api/excel/excel.chartmapoptionsupdatedata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsupdatedata#labelstrategy)|Возвращает или задает стратегию подписей карт ряда для диаграммы с картой региона. Для чтения и записи.|
||[level](/javascript/api/excel/excel.chartmapoptionsupdatedata#level)|Возвращает или задает уровень карты ряда для диаграммы с картой региона. Для чтения и записи.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsupdatedata#projectiontype)|Возвращает или задает тип проекции ряда для диаграммы с картой региона. Для чтения и записи.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[Set (Properties: Excel. Чартпивотоптионс)](/javascript/api/excel/excel.chartpivotoptions#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Чартпивотоптионсупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.chartpivotoptions#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Указывает, следует ли отображать кнопки поля оси в сводной диаграмме. Свойство ShowAxisFieldButtons соответствует команде "Показать кнопки поля оси" в раскрывающемся списке "Кнопки полей" вкладки "Анализировать", доступной при выделении сводной диаграммы.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Указывает, следует ли отображать кнопки поля легенды в сводной диаграмме.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Указывает, следует ли отображать кнопки поля фильтра отчета в сводной диаграмме.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Указывает, следует ли отображать кнопки поля значения в сводной диаграмме.|
|[Чартпивотоптионсдата](/javascript/api/excel/excel.chartpivotoptionsdata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showaxisfieldbuttons)|Указывает, следует ли отображать кнопки поля оси в сводной диаграмме. Свойство ShowAxisFieldButtons соответствует команде "Показать кнопки поля оси" в раскрывающемся списке "Кнопки полей" вкладки "Анализировать", доступной при выделении сводной диаграммы.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showlegendfieldbuttons)|Указывает, следует ли отображать кнопки поля легенды в сводной диаграмме.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showreportfilterfieldbuttons)|Указывает, следует ли отображать кнопки поля фильтра отчета в сводной диаграмме.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showvaluefieldbuttons)|Указывает, следует ли отображать кнопки поля значения в сводной диаграмме.|
|[Чартпивотоптионслоадоптионс](/javascript/api/excel/excel.chartpivotoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartpivotoptionsloadoptions#$all)||
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showaxisfieldbuttons)|Указывает, следует ли отображать кнопки поля оси в сводной диаграмме. Свойство ShowAxisFieldButtons соответствует команде "Показать кнопки поля оси" в раскрывающемся списке "Кнопки полей" вкладки "Анализировать", доступной при выделении сводной диаграммы.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showlegendfieldbuttons)|Указывает, следует ли отображать кнопки поля легенды в сводной диаграмме.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showreportfilterfieldbuttons)|Указывает, следует ли отображать кнопки поля фильтра отчета в сводной диаграмме.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showvaluefieldbuttons)|Указывает, следует ли отображать кнопки поля значения в сводной диаграмме.|
|[Чартпивотоптионсупдатедата](/javascript/api/excel/excel.chartpivotoptionsupdatedata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showaxisfieldbuttons)|Указывает, следует ли отображать кнопки поля оси в сводной диаграмме. Свойство ShowAxisFieldButtons соответствует команде "Показать кнопки поля оси" в раскрывающемся списке "Кнопки полей" вкладки "Анализировать", доступной при выделении сводной диаграммы.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showlegendfieldbuttons)|Указывает, следует ли отображать кнопки поля легенды в сводной диаграмме.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showreportfilterfieldbuttons)|Указывает, следует ли отображать кнопки поля фильтра отчета в сводной диаграмме.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showvaluefieldbuttons)|Указывает, следует ли отображать кнопки поля значения в сводной диаграмме.|
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
|[Чартсериесколлектионлоадоптионс](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#binoptions)|Для каждого элемента в коллекции: инкапсулирует параметры ячейки для гистограмм и диаграмм Парето.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#boxwhiskeroptions)|Для каждого элемента в коллекции: инкапсулирует параметры для полей и диаграмм вхискер.|
||[bubbleScale](/javascript/api/excel/excel.chartseriescollectionloadoptions#bubblescale)|Для каждого элемента в коллекции: это может быть целое число от 0 (ноль) до 300, представляющее процентное значение размера по умолчанию. Это свойство применяется только к пузырьковым диаграммам. Для чтения и записи.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumcolor)|Для каждого элемента в коллекции: Возвращает или задает цвет максимального значения ряда диаграммы карты областей. Для чтения и записи.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumtype)|Для каждого элемента в коллекции: Возвращает или задает тип максимального значения для ряда диаграммы карты областей. Для чтения и записи.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumvalue)|Для каждого элемента в коллекции: Возвращает или задает максимальное значение ряда диаграммы карты областей. Для чтения и записи.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointcolor)|Для каждого элемента в коллекции: Возвращает или задает цвет для среднего значения ряда диаграммы карты областей. Для чтения и записи.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointtype)|Для каждого элемента в коллекции: Возвращает или задает тип среднего значения для ряда диаграммы карты областей. Для чтения и записи.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointvalue)|Для каждого элемента в коллекции: Возвращает или задает среднее значение ряда диаграммы карты областей. Для чтения и записи.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumcolor)|Для каждого элемента в коллекции: Возвращает или задает цвет для минимального значения ряда диаграммы карты областей. Для чтения и записи.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumtype)|Для каждого элемента в коллекции: Возвращает или задает тип для минимального значения ряда диаграммы карты областей. Для чтения и записи.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumvalue)|Для каждого элемента в коллекции: Возвращает или задает минимальное значение ряда диаграммы карты областей. Для чтения и записи.|
||[gradientStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientstyle)|Для каждого элемента в коллекции: Возвращает или задает стиль градиента ряда для диаграммы с областью. Для чтения и записи.|
||[invertColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertcolor)|Для каждого элемента в коллекции: Возвращает или задает цвет заливки для отрицательных точек данных в ряду. Для чтения и записи.|
||[mapOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#mapoptions)|Для каждого элемента в коллекции: инкапсулирует параметры диаграммы с картой областей.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriescollectionloadoptions#parentlabelstrategy)|Для каждого элемента в коллекции: Возвращает или задает родительскую стратегию меток меток для диаграммы эта. Для чтения и записи.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showconnectorlines)|Для каждого элемента в коллекции: указывает, отображаются ли соединительные линии в каскадных диаграммах. Для чтения и записи.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showleaderlines)|Для каждого элемента в коллекции: указывает, отображаются ли линии выноски для каждой подписи данных в ряду. Для чтения и записи.|
||[splitValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#splitvalue)|Для каждого элемента в коллекции: Возвращает или задает пороговое значение, которое разделяет два раздела круговой диаграммы или гистограммы. Для чтения и записи.|
||[xErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#xerrorbars)|Для каждого элемента в коллекции: представляет объект области погрешностей ряда диаграммы.|
||[yErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#yerrorbars)|Для каждого элемента в коллекции: представляет объект области погрешностей ряда диаграммы.|
|[Чартсериесдата](/javascript/api/excel/excel.chartseriesdata)|[binOptions](/javascript/api/excel/excel.chartseriesdata#binoptions)|Объединяет параметры интервалов для гистограмм и диаграмм Парето. Только для чтения.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesdata#boxwhiskeroptions)|Объединяет параметры для диаграмм "ящик с усами" Только для чтения.|
||[bubbleScale](/javascript/api/excel/excel.chartseriesdata#bubblescale)|Может быть целым числом от 0 (нуля) до 300, представляющим процентное значение от размера по умолчанию. Это свойство применяется только к пузырьковым диаграммам. Для чтения и записи.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesdata#gradientmaximumcolor)|Возвращает или задает цвет максимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesdata#gradientmaximumtype)|Возвращает или задает тип максимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesdata#gradientmaximumvalue)|Возвращает или задает максимальное значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesdata#gradientmidpointcolor)|Возвращает или задает цвет среднего значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesdata#gradientmidpointtype)|Возвращает или задает тип среднего значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesdata#gradientmidpointvalue)|Возвращает или задает среднее значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesdata#gradientminimumcolor)|Возвращает или задает цвет минимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesdata#gradientminimumtype)|Возвращает или задает тип минимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesdata#gradientminimumvalue)|Возвращает или задает минимальное значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesdata#gradientstyle)|Возвращает или задает стиль градиента ряда для диаграммы с картой региона. Для чтения и записи.|
||[invertColor](/javascript/api/excel/excel.chartseriesdata#invertcolor)|Возвращает или задает цвет заливки для точек отрицательных данных в ряду. Для чтения и записи.|
||[mapOptions](/javascript/api/excel/excel.chartseriesdata#mapoptions)|Объединяет параметры для диаграммы с картой региона. Только для чтения.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesdata#parentlabelstrategy)|Возвращает или задает область стратегии родительских подписей ряда для диаграммы "дерево". Для чтения и записи.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesdata#showconnectorlines)|Указывает, отображаются ли соединительные линии в каскадных диаграммах. Для чтения и записи.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesdata#showleaderlines)|Указывает, отображаются ли линии выноски для каждой подписи данных в ряду. Для чтения и записи.|
||[splitValue](/javascript/api/excel/excel.chartseriesdata#splitvalue)|Возвращает или задает пороговое значение, разделяющее два раздела вторичной круговой диаграммы или вторичной гистограммы. Для чтения и записи.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesdata#xerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesdata#yerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
|[Чартсериеслоадоптионс](/javascript/api/excel/excel.chartseriesloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriesloadoptions#binoptions)|Объединяет параметры интервалов для гистограмм и диаграмм Парето.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesloadoptions#boxwhiskeroptions)|Объединяет параметры для диаграмм "ящик с усами"|
||[bubbleScale](/javascript/api/excel/excel.chartseriesloadoptions#bubblescale)|Может быть целым числом от 0 (нуля) до 300, представляющим процентное значение от размера по умолчанию. Это свойство применяется только к пузырьковым диаграммам. Для чтения и записи.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumcolor)|Возвращает или задает цвет максимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumtype)|Возвращает или задает тип максимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumvalue)|Возвращает или задает максимальное значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointcolor)|Возвращает или задает цвет среднего значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointtype)|Возвращает или задает тип среднего значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointvalue)|Возвращает или задает среднее значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumcolor)|Возвращает или задает цвет минимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumtype)|Возвращает или задает тип минимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumvalue)|Возвращает или задает минимальное значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesloadoptions#gradientstyle)|Возвращает или задает стиль градиента ряда для диаграммы с картой региона. Для чтения и записи.|
||[invertColor](/javascript/api/excel/excel.chartseriesloadoptions#invertcolor)|Возвращает или задает цвет заливки для точек отрицательных данных в ряду. Для чтения и записи.|
||[mapOptions](/javascript/api/excel/excel.chartseriesloadoptions#mapoptions)|Объединяет параметры для диаграммы с картой региона.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesloadoptions#parentlabelstrategy)|Возвращает или задает область стратегии родительских подписей ряда для диаграммы "дерево". Для чтения и записи.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesloadoptions#showconnectorlines)|Указывает, отображаются ли соединительные линии в каскадных диаграммах. Для чтения и записи.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesloadoptions#showleaderlines)|Указывает, отображаются ли линии выноски для каждой подписи данных в ряду. Для чтения и записи.|
||[splitValue](/javascript/api/excel/excel.chartseriesloadoptions#splitvalue)|Возвращает или задает пороговое значение, разделяющее два раздела вторичной круговой диаграммы или вторичной гистограммы. Для чтения и записи.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#xerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#yerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
|[Чартсериесупдатедата](/javascript/api/excel/excel.chartseriesupdatedata)|[binOptions](/javascript/api/excel/excel.chartseriesupdatedata#binoptions)|Объединяет параметры интервалов для гистограмм и диаграмм Парето.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesupdatedata#boxwhiskeroptions)|Объединяет параметры для диаграмм "ящик с усами"|
||[bubbleScale](/javascript/api/excel/excel.chartseriesupdatedata#bubblescale)|Может быть целым числом от 0 (нуля) до 300, представляющим процентное значение от размера по умолчанию. Это свойство применяется только к пузырьковым диаграммам. Для чтения и записи.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumcolor)|Возвращает или задает цвет максимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumtype)|Возвращает или задает тип максимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumvalue)|Возвращает или задает максимальное значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointcolor)|Возвращает или задает цвет среднего значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointtype)|Возвращает или задает тип среднего значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointvalue)|Возвращает или задает среднее значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumcolor)|Возвращает или задает цвет минимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumtype)|Возвращает или задает тип минимального значения для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumvalue)|Возвращает или задает минимальное значение для ряда диаграммы с картой региона. Для чтения и записи.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesupdatedata#gradientstyle)|Возвращает или задает стиль градиента ряда для диаграммы с картой региона. Для чтения и записи.|
||[invertColor](/javascript/api/excel/excel.chartseriesupdatedata#invertcolor)|Возвращает или задает цвет заливки для точек отрицательных данных в ряду. Для чтения и записи.|
||[mapOptions](/javascript/api/excel/excel.chartseriesupdatedata#mapoptions)|Объединяет параметры для диаграммы с картой региона.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesupdatedata#parentlabelstrategy)|Возвращает или задает область стратегии родительских подписей ряда для диаграммы "дерево". Для чтения и записи.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesupdatedata#showconnectorlines)|Указывает, отображаются ли соединительные линии в каскадных диаграммах. Для чтения и записи.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesupdatedata#showleaderlines)|Указывает, отображаются ли линии выноски для каждой подписи данных в ряду. Для чтения и записи.|
||[splitValue](/javascript/api/excel/excel.chartseriesupdatedata#splitvalue)|Возвращает или задает пороговое значение, разделяющее два раздела вторичной круговой диаграммы или вторичной гистограммы. Для чтения и записи.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#xerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#yerrorbars)|Представляет объект планки погрешностей для ряда диаграммы.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[Чарттрендлинелабелдата](/javascript/api/excel/excel.charttrendlinelabeldata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[Чарттрендлинелабеллоадоптионс](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[Чарттрендлинелабелупдатедата](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#linknumberformat)|Логическое значение, которое указывает, связан ли числовой формат с ячейками (с изменением числового формата в подписях при его изменении в ячейках).|
|[Чартупдатедата](/javascript/api/excel/excel.chartupdatedata)|[pivotOptions](/javascript/api/excel/excel.chartupdatedata#pivotoptions)|Объединяет параметры для сводной диаграммы.|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Представляет свойство `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Представляет свойство `columnIndex`.|
|[Колумнпропертиеслоадоптионс](/javascript/api/excel/excel.columnpropertiesloadoptions)|[columnHidden](/javascript/api/excel/excel.columnpropertiesloadoptions#columnhidden)|Указывает, следует ли загружать `columnHidden` свойство.|
||[columnIndex](/javascript/api/excel/excel.columnpropertiesloadoptions#columnindex)|Указывает, следует ли загружать `columnIndex` свойство.|
||[columnWidth](/javascript/api/excel/excel.columnpropertiesloadoptions#columnwidth)||
||[Format: Excel. Целлпропертиесформатлоадоптионс & {
            columnWidth?] (/жаваскрипт/АПИ/ексцел/ексцел.колумнпропертиеслоадоптионс # Format)|Указывает, следует ли загружать `format` свойство.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, к которым применено условное форматирование. Только для чтения.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, с недопустимыми значениями ячеек. Если все значения ячеек являются допустимыми, эта функция выдаст ошибку ItemNotFound.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Возвращает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, с недопустимыми значениями ячеек. Если все значения ячеек являются допустимыми, эта функция вернет значение null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|Свойство, используемое фильтром для расширенной фильтрации по объектам richvalue.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Возвращает идентификатор фигуры. Только для чтения.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Возвращает объект Shape для геометрической фигуры. Только для чтения.|
|[Жеометрикшапедата](/javascript/api/excel/excel.geometricshapedata)|[id](/javascript/api/excel/excel.geometricshapedata#id)|Возвращает идентификатор фигуры. Только для чтения.|
|[Жеометрикшапелоадоптионс](/javascript/api/excel/excel.geometricshapeloadoptions)|[$all](/javascript/api/excel/excel.geometricshapeloadoptions#$all)||
||[id](/javascript/api/excel/excel.geometricshapeloadoptions#id)|Возвращает идентификатор фигуры. Только для чтения.|
||[shape](/javascript/api/excel/excel.geometricshapeloadoptions#shape)|Возвращает объект Shape для геометрической фигуры.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Возвращает количество фигур в группе фигур. Только для чтения.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Получает фигуру по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Получает фигуру на основе ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Граупшапеколлектионлоадоптионс](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[$all](/javascript/api/excel/excel.groupshapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttextdescription)|Для каждого элемента в коллекции: Возвращает или задает текст альтернативного описания для объекта Shape.|
||[altTextTitle](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttexttitle)|Для каждого элемента в коллекции: Возвращает или задает текст альтернативного заголовка для объекта Shape.|
||[connectionSiteCount](/javascript/api/excel/excel.groupshapecollectionloadoptions#connectionsitecount)|Для каждого элемента в коллекции: Возвращает число сайтов подключения на этой фигуре. Только для чтения.|
||[fill](/javascript/api/excel/excel.groupshapecollectionloadoptions#fill)|Для каждого элемента в коллекции: возвращает форматирование заливки данной фигуры.|
||[geometricShape](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshape)|Для каждого элемента в коллекции: возвращает геометрическую фигуру, связанную с фигурой. Если тип фигуры отличается от GeometricShape, возникает ошибка.|
||[geometricShapeType](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshapetype)|Для каждого элемента в коллекции: представляет тип геометрической фигуры для этой геометрической фигуры. Дополнительные сведения см. в статье Excel.GeometricShapeType. Возвращает значение null, если тип фигуры отличается от GeometricShape.|
||[group](/javascript/api/excel/excel.groupshapecollectionloadoptions#group)|Для каждого элемента в коллекции: Возвращает группу фигур, связанную с фигурой. Если тип фигуры отличается от GroupShape, возникает ошибка.|
||[height](/javascript/api/excel/excel.groupshapecollectionloadoptions#height)|Для каждого элемента в коллекции: представляет высоту фигуры в пунктах.|
||[id](/javascript/api/excel/excel.groupshapecollectionloadoptions#id)|Для каждого элемента в коллекции: представляет идентификатор фигуры. Только для чтения.|
||[image](/javascript/api/excel/excel.groupshapecollectionloadoptions#image)|Для каждого элемента в коллекции: возвращает изображение, связанное с фигурой. Если тип фигуры отличается от Image, возникает ошибка.|
||[left](/javascript/api/excel/excel.groupshapecollectionloadoptions#left)|Для каждого элемента в коллекции: расстояние (в пунктах) от левой стороны фигуры до левой стороны листа.|
||[level](/javascript/api/excel/excel.groupshapecollectionloadoptions#level)|Для каждого элемента в коллекции: представляет уровень указанной фигуры. Например, уровень 0 означает, что фигура не является частью групп; уровень 1 означает, что фигура является частью группы верхнего уровня; уровень 2 означает, что фигура является частью подгруппы верхнего уровня.|
||[line](/javascript/api/excel/excel.groupshapecollectionloadoptions#line)|Для каждого элемента в коллекции: Возвращает строку, связанную с фигурой. Если тип фигуры отличается от Line, возникает ошибка.|
||[lineFormat](/javascript/api/excel/excel.groupshapecollectionloadoptions#lineformat)|Для каждого элемента в коллекции: возвращает форматирование строки этой фигуры.|
||[lockAspectRatio](/javascript/api/excel/excel.groupshapecollectionloadoptions#lockaspectratio)|Для каждого элемента в коллекции: указывает, заблокировано ли пропорции данной фигуры.|
||[name](/javascript/api/excel/excel.groupshapecollectionloadoptions#name)|Для каждого элемента в коллекции: представляет имя фигуры.|
||[parentGroup](/javascript/api/excel/excel.groupshapecollectionloadoptions#parentgroup)|Для каждого элемента в коллекции: представляет родительскую группу этой фигуры.|
||[rotation](/javascript/api/excel/excel.groupshapecollectionloadoptions#rotation)|Для каждого элемента в коллекции — представляет Поворот фигуры в градусах.|
||[textFrame](/javascript/api/excel/excel.groupshapecollectionloadoptions#textframe)|Для каждого элемента в коллекции: Возвращает объект текстового фрейма этой фигуры. Только для чтения.|
||[top](/javascript/api/excel/excel.groupshapecollectionloadoptions#top)|Для каждого элемента в коллекции: расстояние (в пунктах) от верхнего края фигуры до верхнего края листа.|
||[type](/javascript/api/excel/excel.groupshapecollectionloadoptions#type)|Для каждого элемента в коллекции: Возвращает тип этой фигуры. Дополнительные сведения см. в статье Excel.ShapeType. Только для чтения.|
||[visible](/javascript/api/excel/excel.groupshapecollectionloadoptions#visible)|Для каждого элемента в коллекции: представляет видимость этой фигуры.|
||[width](/javascript/api/excel/excel.groupshapecollectionloadoptions#width)|Для каждого элемента в коллекции: представляет ширину фигуры в пунктах.|
||[zOrderPosition](/javascript/api/excel/excel.groupshapecollectionloadoptions#zorderposition)|Для каждого элемента в коллекции: Возвращает позицию указанной фигуры в z-порядке, где 0 представляет нижнюю часть стека заказов. Только для чтения.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Получает или задает центральный нижний колонтитул листа.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Получает или задает центральный верхний колонтитул листа.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Получает или задает левый нижний колонтитул листа.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Получает или задает левый верхний колонтитул листа.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Получает или задает правый нижний колонтитул листа.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Получает или задает правый верхний колонтитул листа.|
||[Set (Properties: Excel. HeaderFooter)](/javascript/api/excel/excel.headerfooter#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Хеадерфутерупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.headerfooter#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Хеадерфутердата](/javascript/api/excel/excel.headerfooterdata)|[centerFooter](/javascript/api/excel/excel.headerfooterdata#centerfooter)|Получает или задает центральный нижний колонтитул листа.|
||[centerHeader](/javascript/api/excel/excel.headerfooterdata#centerheader)|Получает или задает центральный верхний колонтитул листа.|
||[leftFooter](/javascript/api/excel/excel.headerfooterdata#leftfooter)|Получает или задает левый нижний колонтитул листа.|
||[leftHeader](/javascript/api/excel/excel.headerfooterdata#leftheader)|Получает или задает левый верхний колонтитул листа.|
||[rightFooter](/javascript/api/excel/excel.headerfooterdata#rightfooter)|Получает или задает правый нижний колонтитул листа.|
||[rightHeader](/javascript/api/excel/excel.headerfooterdata#rightheader)|Получает или задает правый верхний колонтитул листа.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|Общий колонтитул, используемый для всех страниц, если не указан колонтитул четных и нечетных страниц или первой страницы.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|Колонтитул для четных страниц, для нечетных страниц нужно указывать отдельный колонтитул.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|Колонтитул первой страницы, для остальных страниц используется общий или четный и нечетный колонтитулы.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|Колонтитул для нечетных страниц, для четных страниц нужно указывать отдельный колонтитул.|
||[Set (Properties: Excel. Хеадерфутерграуп)](/javascript/api/excel/excel.headerfootergroup#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Хеадерфутерграупупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.headerfootergroup#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Получает или задает состояние, в котором находятся колонтитулы. Дополнительные сведения см. в статье Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Получает или задает отметку, которая указывает, выровнены ли колонтитулы относительно полей страницы, установленных в параметрах макета страницы для листа.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Получает или задает отметку, которая указывает, нужно ли масштабировать колонтитулы с помощью процентных значений, установленных в параметрах макета страницы для листа.|
|[Хеадерфутерграупдата](/javascript/api/excel/excel.headerfootergroupdata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupdata#defaultforallpages)|Общий колонтитул, используемый для всех страниц, если не указан колонтитул четных и нечетных страниц или первой страницы.|
||[evenPages](/javascript/api/excel/excel.headerfootergroupdata#evenpages)|Колонтитул для четных страниц, для нечетных страниц нужно указывать отдельный колонтитул.|
||[firstPage](/javascript/api/excel/excel.headerfootergroupdata#firstpage)|Колонтитул первой страницы, для остальных страниц используется общий или четный и нечетный колонтитулы.|
||[oddPages](/javascript/api/excel/excel.headerfootergroupdata#oddpages)|Колонтитул для нечетных страниц, для четных страниц нужно указывать отдельный колонтитул.|
||[state](/javascript/api/excel/excel.headerfootergroupdata#state)|Получает или задает состояние, в котором находятся колонтитулы. Дополнительные сведения см. в статье Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupdata#usesheetmargins)|Получает или задает отметку, которая указывает, выровнены ли колонтитулы относительно полей страницы, установленных в параметрах макета страницы для листа.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupdata#usesheetscale)|Получает или задает отметку, которая указывает, нужно ли масштабировать колонтитулы с помощью процентных значений, установленных в параметрах макета страницы для листа.|
|[Хеадерфутерграуплоадоптионс](/javascript/api/excel/excel.headerfootergrouploadoptions)|[$all](/javascript/api/excel/excel.headerfootergrouploadoptions#$all)||
||[defaultForAllPages](/javascript/api/excel/excel.headerfootergrouploadoptions#defaultforallpages)|Общий колонтитул, используемый для всех страниц, если не указан колонтитул четных и нечетных страниц или первой страницы.|
||[evenPages](/javascript/api/excel/excel.headerfootergrouploadoptions#evenpages)|Колонтитул для четных страниц, для нечетных страниц нужно указывать отдельный колонтитул.|
||[firstPage](/javascript/api/excel/excel.headerfootergrouploadoptions#firstpage)|Колонтитул первой страницы, для остальных страниц используется общий или четный и нечетный колонтитулы.|
||[oddPages](/javascript/api/excel/excel.headerfootergrouploadoptions#oddpages)|Колонтитул для нечетных страниц, для четных страниц нужно указывать отдельный колонтитул.|
||[state](/javascript/api/excel/excel.headerfootergrouploadoptions#state)|Получает или задает состояние, в котором находятся колонтитулы. Дополнительные сведения см. в статье Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetmargins)|Получает или задает отметку, которая указывает, выровнены ли колонтитулы относительно полей страницы, установленных в параметрах макета страницы для листа.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetscale)|Получает или задает отметку, которая указывает, нужно ли масштабировать колонтитулы с помощью процентных значений, установленных в параметрах макета страницы для листа.|
|[Хеадерфутерграупупдатедата](/javascript/api/excel/excel.headerfootergroupupdatedata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupupdatedata#defaultforallpages)|Общий колонтитул, используемый для всех страниц, если не указан колонтитул четных и нечетных страниц или первой страницы.|
||[evenPages](/javascript/api/excel/excel.headerfootergroupupdatedata#evenpages)|Колонтитул для четных страниц, для нечетных страниц нужно указывать отдельный колонтитул.|
||[firstPage](/javascript/api/excel/excel.headerfootergroupupdatedata#firstpage)|Колонтитул первой страницы, для остальных страниц используется общий или четный и нечетный колонтитулы.|
||[oddPages](/javascript/api/excel/excel.headerfootergroupupdatedata#oddpages)|Колонтитул для нечетных страниц, для четных страниц нужно указывать отдельный колонтитул.|
||[state](/javascript/api/excel/excel.headerfootergroupupdatedata#state)|Получает или задает состояние, в котором находятся колонтитулы. Дополнительные сведения см. в статье Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetmargins)|Получает или задает отметку, которая указывает, выровнены ли колонтитулы относительно полей страницы, установленных в параметрах макета страницы для листа.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetscale)|Получает или задает отметку, которая указывает, нужно ли масштабировать колонтитулы с помощью процентных значений, установленных в параметрах макета страницы для листа.|
|[Хеадерфутерлоадоптионс](/javascript/api/excel/excel.headerfooterloadoptions)|[$all](/javascript/api/excel/excel.headerfooterloadoptions#$all)||
||[centerFooter](/javascript/api/excel/excel.headerfooterloadoptions#centerfooter)|Получает или задает центральный нижний колонтитул листа.|
||[centerHeader](/javascript/api/excel/excel.headerfooterloadoptions#centerheader)|Получает или задает центральный верхний колонтитул листа.|
||[leftFooter](/javascript/api/excel/excel.headerfooterloadoptions#leftfooter)|Получает или задает левый нижний колонтитул листа.|
||[leftHeader](/javascript/api/excel/excel.headerfooterloadoptions#leftheader)|Получает или задает левый верхний колонтитул листа.|
||[rightFooter](/javascript/api/excel/excel.headerfooterloadoptions#rightfooter)|Получает или задает правый нижний колонтитул листа.|
||[rightHeader](/javascript/api/excel/excel.headerfooterloadoptions#rightheader)|Получает или задает правый верхний колонтитул листа.|
|[Хеадерфутерупдатедата](/javascript/api/excel/excel.headerfooterupdatedata)|[centerFooter](/javascript/api/excel/excel.headerfooterupdatedata#centerfooter)|Получает или задает центральный нижний колонтитул листа.|
||[centerHeader](/javascript/api/excel/excel.headerfooterupdatedata#centerheader)|Получает или задает центральный верхний колонтитул листа.|
||[leftFooter](/javascript/api/excel/excel.headerfooterupdatedata#leftfooter)|Получает или задает левый нижний колонтитул листа.|
||[leftHeader](/javascript/api/excel/excel.headerfooterupdatedata#leftheader)|Получает или задает левый верхний колонтитул листа.|
||[rightFooter](/javascript/api/excel/excel.headerfooterupdatedata#rightfooter)|Получает или задает правый нижний колонтитул листа.|
||[rightHeader](/javascript/api/excel/excel.headerfooterupdatedata#rightheader)|Получает или задает правый верхний колонтитул листа.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Возвращает формат изображения. Только для чтения.|
||[id](/javascript/api/excel/excel.image#id)|Представляет идентификатор фигуры для объекта image. Только для чтения.|
||[shape](/javascript/api/excel/excel.image#shape)|Возвращает объект Shape, связанный с изображением. Только для чтения.|
|[Имажедата](/javascript/api/excel/excel.imagedata)|[format](/javascript/api/excel/excel.imagedata#format)|Возвращает формат изображения. Только для чтения.|
||[id](/javascript/api/excel/excel.imagedata#id)|Представляет идентификатор фигуры для объекта image. Только для чтения.|
|[Имажелоадоптионс](/javascript/api/excel/excel.imageloadoptions)|[$all](/javascript/api/excel/excel.imageloadoptions#$all)||
||[format](/javascript/api/excel/excel.imageloadoptions#format)|Возвращает формат изображения. Только для чтения.|
||[id](/javascript/api/excel/excel.imageloadoptions#id)|Представляет идентификатор фигуры для объекта image. Только для чтения.|
||[shape](/javascript/api/excel/excel.imageloadoptions#shape)|Возвращает объект Shape, связанный с изображением.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Значение true, если в Excel используется итерация для разрешения циклических ссылок.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Возвращает или задает максимальное изменение между итерациями при разрешении в Excel циклических ссылок.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Возвращает или задает максимальное количество итераций, которое можно использовать в Excel для разрешения циклической ссылки.|
||[Set (Properties: Excel. Итеративекалкулатион)](/javascript/api/excel/excel.iterativecalculation#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Итеративекалкулатионупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.iterativecalculation#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Итеративекалкулатиондата](/javascript/api/excel/excel.iterativecalculationdata)|[enabled](/javascript/api/excel/excel.iterativecalculationdata#enabled)|Значение true, если в Excel используется итерация для разрешения циклических ссылок.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationdata#maxchange)|Возвращает или задает максимальное изменение между итерациями при разрешении в Excel циклических ссылок.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationdata#maxiteration)|Возвращает или задает максимальное количество итераций, которое можно использовать в Excel для разрешения циклической ссылки.|
|[Итеративекалкулатионлоадоптионс](/javascript/api/excel/excel.iterativecalculationloadoptions)|[$all](/javascript/api/excel/excel.iterativecalculationloadoptions#$all)||
||[enabled](/javascript/api/excel/excel.iterativecalculationloadoptions#enabled)|Значение true, если в Excel используется итерация для разрешения циклических ссылок.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationloadoptions#maxchange)|Возвращает или задает максимальное изменение между итерациями при разрешении в Excel циклических ссылок.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationloadoptions#maxiteration)|Возвращает или задает максимальное количество итераций, которое можно использовать в Excel для разрешения циклической ссылки.|
|[Итеративекалкулатионупдатедата](/javascript/api/excel/excel.iterativecalculationupdatedata)|[enabled](/javascript/api/excel/excel.iterativecalculationupdatedata#enabled)|Значение true, если в Excel используется итерация для разрешения циклических ссылок.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationupdatedata#maxchange)|Возвращает или задает максимальное изменение между итерациями при разрешении в Excel циклических ссылок.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationupdatedata#maxiteration)|Возвращает или задает максимальное количество итераций, которое можно использовать в Excel для разрешения циклической ссылки.|
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
||[Set (Properties: Excel. line)](/javascript/api/excel/excel.line#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Линеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.line#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[Линедата](/javascript/api/excel/excel.linedata)|[beginArrowheadLength](/javascript/api/excel/excel.linedata#beginarrowheadlength)|Представляет длину наконечника в начале указанной линии.|
||[beginArrowheadStyle](/javascript/api/excel/excel.linedata#beginarrowheadstyle)|Представляет стиль наконечника в начале указанной линии.|
||[beginArrowheadWidth](/javascript/api/excel/excel.linedata#beginarrowheadwidth)|Представляет ширину наконечника в начале указанной линии.|
||[beginConnectedSite](/javascript/api/excel/excel.linedata#beginconnectedsite)|Представляет точку соединения, к которой привязано начало соединительной линии. Только для чтения. Возвращает значение null, если начало линии не привязано к фигуре.|
||[connectorType](/javascript/api/excel/excel.linedata#connectortype)|Представляет тип соединительной линии.|
||[endArrowheadLength](/javascript/api/excel/excel.linedata#endarrowheadlength)|Представляет длину наконечника в конце указанной линии.|
||[endArrowheadStyle](/javascript/api/excel/excel.linedata#endarrowheadstyle)|Представляет стиль наконечника в конце указанной линии.|
||[endArrowheadWidth](/javascript/api/excel/excel.linedata#endarrowheadwidth)|Представляет ширину наконечника в конце указанной линии.|
||[endConnectedSite](/javascript/api/excel/excel.linedata#endconnectedsite)|Представляет точку соединения, к которой привязан конец соединительной линии. Только для чтения. Возвращает значение null, если конец линии не привязан к фигуре.|
||[id](/javascript/api/excel/excel.linedata#id)|Представляет идентификатор фигуры. Только для чтения.|
||[isBeginConnected](/javascript/api/excel/excel.linedata#isbeginconnected)|Указывает, привязано ли начало указанной линии к фигуре. Только для чтения.|
||[isEndConnected](/javascript/api/excel/excel.linedata#isendconnected)|Указывает, привязан ли конец указанной линии к фигуре. Только для чтения.|
|[Линелоадоптионс](/javascript/api/excel/excel.lineloadoptions)|[$all](/javascript/api/excel/excel.lineloadoptions#$all)||
||[beginArrowheadLength](/javascript/api/excel/excel.lineloadoptions#beginarrowheadlength)|Представляет длину наконечника в начале указанной линии.|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#beginarrowheadstyle)|Представляет стиль наконечника в начале указанной линии.|
||[beginArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#beginarrowheadwidth)|Представляет ширину наконечника в начале указанной линии.|
||[beginConnectedShape](/javascript/api/excel/excel.lineloadoptions#beginconnectedshape)|Представляет фигуру, к которой привязано начало указанной линии.|
||[beginConnectedSite](/javascript/api/excel/excel.lineloadoptions#beginconnectedsite)|Представляет точку соединения, к которой привязано начало соединительной линии. Только для чтения. Возвращает значение null, если начало линии не привязано к фигуре.|
||[connectorType](/javascript/api/excel/excel.lineloadoptions#connectortype)|Представляет тип соединительной линии.|
||[endArrowheadLength](/javascript/api/excel/excel.lineloadoptions#endarrowheadlength)|Представляет длину наконечника в конце указанной линии.|
||[endArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#endarrowheadstyle)|Представляет стиль наконечника в конце указанной линии.|
||[endArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#endarrowheadwidth)|Представляет ширину наконечника в конце указанной линии.|
||[endConnectedShape](/javascript/api/excel/excel.lineloadoptions#endconnectedshape)|Представляет фигуру, к которой привязан конец указанной линии.|
||[endConnectedSite](/javascript/api/excel/excel.lineloadoptions#endconnectedsite)|Представляет точку соединения, к которой привязан конец соединительной линии. Только для чтения. Возвращает значение null, если конец линии не привязан к фигуре.|
||[id](/javascript/api/excel/excel.lineloadoptions#id)|Представляет идентификатор фигуры. Только для чтения.|
||[isBeginConnected](/javascript/api/excel/excel.lineloadoptions#isbeginconnected)|Указывает, привязано ли начало указанной линии к фигуре. Только для чтения.|
||[isEndConnected](/javascript/api/excel/excel.lineloadoptions#isendconnected)|Указывает, привязан ли конец указанной линии к фигуре. Только для чтения.|
||[shape](/javascript/api/excel/excel.lineloadoptions#shape)|Возвращает объект Shape, связанный с линией.|
|[Линеупдатедата](/javascript/api/excel/excel.lineupdatedata)|[beginArrowheadLength](/javascript/api/excel/excel.lineupdatedata#beginarrowheadlength)|Представляет длину наконечника в начале указанной линии.|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#beginarrowheadstyle)|Представляет стиль наконечника в начале указанной линии.|
||[beginArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#beginarrowheadwidth)|Представляет ширину наконечника в начале указанной линии.|
||[connectorType](/javascript/api/excel/excel.lineupdatedata#connectortype)|Представляет тип соединительной линии.|
||[endArrowheadLength](/javascript/api/excel/excel.lineupdatedata#endarrowheadlength)|Представляет длину наконечника в конце указанной линии.|
||[endArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#endarrowheadstyle)|Представляет стиль наконечника в конце указанной линии.|
||[endArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#endarrowheadwidth)|Представляет ширину наконечника в конце указанной линии.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Удаляет объект разрыва страницы.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Получает первую ячейку после разрыва страницы.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Представляет индекс столбца для разрыва страницы|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Представляет индекс строки для разрыва страницы|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Добавляет разрыв страницы перед левой верхней ячейкой указанного диапазона.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Получает количество разрывов страниц в коллекции.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Получает объект разрыва страницы по индексу.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Сбрасывает все добавленные вручную разрывы страниц в коллекции.|
|[Пажебреакколлектионлоадоптионс](/javascript/api/excel/excel.pagebreakcollectionloadoptions)|[$all](/javascript/api/excel/excel.pagebreakcollectionloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#columnindex)|Для каждого элемента в коллекции: представляет индекс столбца для разрыва страницы|
||[rowIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#rowindex)|Для каждого элемента в коллекции: представляет индекс строки для разрыва страницы|
|[Пажебреакдата](/javascript/api/excel/excel.pagebreakdata)|[columnIndex](/javascript/api/excel/excel.pagebreakdata#columnindex)|Представляет индекс столбца для разрыва страницы|
||[rowIndex](/javascript/api/excel/excel.pagebreakdata#rowindex)|Представляет индекс строки для разрыва страницы|
|[Пажебреаклоадоптионс](/javascript/api/excel/excel.pagebreakloadoptions)|[$all](/javascript/api/excel/excel.pagebreakloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakloadoptions#columnindex)|Представляет индекс столбца для разрыва страницы|
||[rowIndex](/javascript/api/excel/excel.pagebreakloadoptions#rowindex)|Представляет индекс строки для разрыва страницы|
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
||[Set (Properties: Excel. PageLayout)](/javascript/api/excel/excel.pagelayout#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Пажелайаутупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.pagelayout#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Задает область печати листа.|
||[setPrintMargins(unit: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Задает поля страницы с единицами измерения для листа.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Задает поля страницы с единицами измерения для листа.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Задает столбцы, содержащие ячейки, которые должны повторяться слева на каждой странице при печати листа.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Задает строки, содержащие ячейки, которые должны повторяться сверху каждой страницы при печати листа.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Получает или задает верхнее поле листа (в пунктах) для использования при печати.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Получает или задает параметры масштабирования при печати листа.|
|[Пажелайаутдата](/javascript/api/excel/excel.pagelayoutdata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutdata#blackandwhite)|Получает или задает параметр черно-белой печати листа.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutdata#bottommargin)|Получает или задает нижнее поле страницы листа, чтобы использовать для печати в пунктах.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutdata#centerhorizontally)|Получает или задает отметку выравнивания листа по горизонтали относительно центра. Эта отметка определяет, выравнивается ли лист по горизонтали относительно центра при печати.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutdata#centervertically)|Получает или задает отметку выравнивания листа по вертикали относительно центра. Эта отметка определяет, выравнивается ли лист по вертикали относительно центра при печати.|
||[draftMode](/javascript/api/excel/excel.pagelayoutdata#draftmode)|Получает или задает параметр режима черновика листа. Если присвоено значение true, лист будет печататься без рисунков.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutdata#firstpagenumber)|Получает или задает номер первой страницы листа для печати. Значение null представляет автоматическую нумерацию страниц.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutdata#footermargin)|Получает или задает поле нижнего колонтитула листа (в пунктах) для использования при печати.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutdata#headermargin)|Получает или задает поле верхнего колонтитула листа (в пунктах) для использования при печати.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutdata#headersfooters)|Настройка колонтитулов для листа.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutdata#leftmargin)|Получает или задает левое поле листа (в пунктах) для использования при печати.|
||[orientation](/javascript/api/excel/excel.pagelayoutdata#orientation)|Получает или задает ориентацию страницы для листа.|
||[paperSize](/javascript/api/excel/excel.pagelayoutdata#papersize)|Получает или задает размер бумаги для листа.|
||[printComments](/javascript/api/excel/excel.pagelayoutdata#printcomments)|Получает или задает, должны ли отображаться примечания листа при печати.|
||[printErrors](/javascript/api/excel/excel.pagelayoutdata#printerrors)|Получает или задает параметр ошибок печати листа.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutdata#printgridlines)|Получает или задает отметку печати линий сетки листа. Эта отметка определяет, печатаются ли линии сетки.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutdata#printheadings)|Получает или задает отметку печати заголовков листа. Эта отметка определяет, печатаются ли заголовки.|
||[printOrder](/javascript/api/excel/excel.pagelayoutdata#printorder)|Получает или задает параметр порядка печати листа. Определяет порядок, использующийся при обработке распечатываемых номеров страниц.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutdata#rightmargin)|Получает или задает правое поле листа (в пунктах) для использования при печати.|
||[topMargin](/javascript/api/excel/excel.pagelayoutdata#topmargin)|Получает или задает верхнее поле листа (в пунктах) для использования при печати.|
||[zoom](/javascript/api/excel/excel.pagelayoutdata#zoom)|Получает или задает параметры масштабирования при печати листа.|
|[Пажелайаутлоадоптионс](/javascript/api/excel/excel.pagelayoutloadoptions)|[$all](/javascript/api/excel/excel.pagelayoutloadoptions#$all)||
||[blackAndWhite](/javascript/api/excel/excel.pagelayoutloadoptions#blackandwhite)|Получает или задает параметр черно-белой печати листа.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutloadoptions#bottommargin)|Получает или задает нижнее поле страницы листа, чтобы использовать для печати в пунктах.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutloadoptions#centerhorizontally)|Получает или задает отметку выравнивания листа по горизонтали относительно центра. Эта отметка определяет, выравнивается ли лист по горизонтали относительно центра при печати.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutloadoptions#centervertically)|Получает или задает отметку выравнивания листа по вертикали относительно центра. Эта отметка определяет, выравнивается ли лист по вертикали относительно центра при печати.|
||[draftMode](/javascript/api/excel/excel.pagelayoutloadoptions#draftmode)|Получает или задает параметр режима черновика листа. Если присвоено значение true, лист будет печататься без рисунков.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutloadoptions#firstpagenumber)|Получает или задает номер первой страницы листа для печати. Значение null представляет автоматическую нумерацию страниц.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#footermargin)|Получает или задает поле нижнего колонтитула листа (в пунктах) для использования при печати.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#headermargin)|Получает или задает поле верхнего колонтитула листа (в пунктах) для использования при печати.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutloadoptions#headersfooters)|Настройка колонтитулов для листа.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutloadoptions#leftmargin)|Получает или задает левое поле листа (в пунктах) для использования при печати.|
||[orientation](/javascript/api/excel/excel.pagelayoutloadoptions#orientation)|Получает или задает ориентацию страницы для листа.|
||[paperSize](/javascript/api/excel/excel.pagelayoutloadoptions#papersize)|Получает или задает размер бумаги для листа.|
||[printComments](/javascript/api/excel/excel.pagelayoutloadoptions#printcomments)|Получает или задает, должны ли отображаться примечания листа при печати.|
||[printErrors](/javascript/api/excel/excel.pagelayoutloadoptions#printerrors)|Получает или задает параметр ошибок печати листа.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutloadoptions#printgridlines)|Получает или задает отметку печати линий сетки листа. Эта отметка определяет, печатаются ли линии сетки.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutloadoptions#printheadings)|Получает или задает отметку печати заголовков листа. Эта отметка определяет, печатаются ли заголовки.|
||[printOrder](/javascript/api/excel/excel.pagelayoutloadoptions#printorder)|Получает или задает параметр порядка печати листа. Определяет порядок, использующийся при обработке распечатываемых номеров страниц.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutloadoptions#rightmargin)|Получает или задает правое поле листа (в пунктах) для использования при печати.|
||[topMargin](/javascript/api/excel/excel.pagelayoutloadoptions#topmargin)|Получает или задает верхнее поле листа (в пунктах) для использования при печати.|
||[zoom](/javascript/api/excel/excel.pagelayoutloadoptions#zoom)|Получает или задает параметры масштабирования при печати листа.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Представляет нижнее поле макета страницы в указанных единицах измерения для использования при печати.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Представляет поле нижнего колонтитула макета страницы в указанных единицах измерения для использования при печати.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Представляет поле верхнего колонтитула макета страницы в указанных единицах измерения для использования при печати.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Представляет левое поле макета страницы в указанных единицах измерения для использования при печати.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Представляет правое поле макета страницы в указанных единицах измерения для использования при печати.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Представляет верхнее поле макета страницы в указанных единицах измерения для использования при печати.|
|[Пажелайаутупдатедата](/javascript/api/excel/excel.pagelayoutupdatedata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutupdatedata#blackandwhite)|Получает или задает параметр черно-белой печати листа.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutupdatedata#bottommargin)|Получает или задает нижнее поле страницы листа, чтобы использовать для печати в пунктах.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutupdatedata#centerhorizontally)|Получает или задает отметку выравнивания листа по горизонтали относительно центра. Эта отметка определяет, выравнивается ли лист по горизонтали относительно центра при печати.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutupdatedata#centervertically)|Получает или задает отметку выравнивания листа по вертикали относительно центра. Эта отметка определяет, выравнивается ли лист по вертикали относительно центра при печати.|
||[draftMode](/javascript/api/excel/excel.pagelayoutupdatedata#draftmode)|Получает или задает параметр режима черновика листа. Если присвоено значение true, лист будет печататься без рисунков.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutupdatedata#firstpagenumber)|Получает или задает номер первой страницы листа для печати. Значение null представляет автоматическую нумерацию страниц.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#footermargin)|Получает или задает поле нижнего колонтитула листа (в пунктах) для использования при печати.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#headermargin)|Получает или задает поле верхнего колонтитула листа (в пунктах) для использования при печати.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutupdatedata#headersfooters)|Настройка колонтитулов для листа.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutupdatedata#leftmargin)|Получает или задает левое поле листа (в пунктах) для использования при печати.|
||[orientation](/javascript/api/excel/excel.pagelayoutupdatedata#orientation)|Получает или задает ориентацию страницы для листа.|
||[paperSize](/javascript/api/excel/excel.pagelayoutupdatedata#papersize)|Получает или задает размер бумаги для листа.|
||[printComments](/javascript/api/excel/excel.pagelayoutupdatedata#printcomments)|Получает или задает, должны ли отображаться примечания листа при печати.|
||[printErrors](/javascript/api/excel/excel.pagelayoutupdatedata#printerrors)|Получает или задает параметр ошибок печати листа.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutupdatedata#printgridlines)|Получает или задает отметку печати линий сетки листа. Эта отметка определяет, печатаются ли линии сетки.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutupdatedata#printheadings)|Получает или задает отметку печати заголовков листа. Эта отметка определяет, печатаются ли заголовки.|
||[printOrder](/javascript/api/excel/excel.pagelayoutupdatedata#printorder)|Получает или задает параметр порядка печати листа. Определяет порядок, использующийся при обработке распечатываемых номеров страниц.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutupdatedata#rightmargin)|Получает или задает правое поле листа (в пунктах) для использования при печати.|
||[topMargin](/javascript/api/excel/excel.pagelayoutupdatedata#topmargin)|Получает или задает верхнее поле листа (в пунктах) для использования при печати.|
||[zoom](/javascript/api/excel/excel.pagelayoutupdatedata#zoom)|Получает или задает параметры масштабирования при печати листа.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Количество страниц, размещаемых по горизонтали. Это значение может быть равно null, если используется процентный масштаб.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|Значение масштаба печатной страницы может быть равным от 10 до 400. Это значение может быть равно null, если указано размещение по высоте или ширине страницы.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Количество страниц, размещаемых по вертикали. Это значение может быть равно null, если используется процентный масштаб.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[Сортбивалуес (sortBy: "Ascending \| " "Descending", Валуешиерарчи: Excel. DataPivotHierarchy, пивотитемскопе?: Array<PivotItem \| String>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Сортирует сводную таблицу по указанным значениям в определенной области. Область определяет, какие конкретные значения будут использоваться для сортировки|
||[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Сортирует сводную таблицу по указанным значениям в определенной области. Область определяет, какие конкретные значения будут использоваться для сортировки|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Указывает, применяется ли форматирование автоматически при его обновлении или перемещении полей|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Получает объект DataHierarchy, использующийся для вычисления значения в указанном диапазоне сводной таблицы.|
||[getPivotItems(axis: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Получает объекты PivotItem с оси, образующие значение в указанном диапазоне сводной таблицы.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Получает объекты PivotItem с оси, образующие значение в указанном диапазоне сводной таблицы.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Указывает, сохраняется ли форматирование при обновлении или пересчете отчета с помощью таких операций, как сведение, сортировка или изменение элементов полей страницы.|
||[Сетаутосортонцелл (ячейка: \| строка Range, SortBy: "по возрастанию \| " "по убыванию")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Задает для сводной таблицы автоматическую сортировку, используя указанную ячейку, чтобы автоматически выбрать все необходимые условия и контекст. Это работает аналогично применению автоматической сортировки из пользовательского интерфейса.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Задает для сводной таблицы автоматическую сортировку, используя указанную ячейку, чтобы автоматически выбрать все необходимые условия и контекст. Это работает аналогично применению автоматической сортировки из пользовательского интерфейса.|
|[Пивотлайаутдата](/javascript/api/excel/excel.pivotlayoutdata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutdata#autoformat)|Указывает, применяется ли форматирование автоматически при его обновлении или перемещении полей|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutdata#preserveformatting)|Указывает, сохраняется ли форматирование при обновлении или пересчете отчета с помощью таких операций, как сведение, сортировка или изменение элементов полей страницы.|
|[Пивотлайаутлоадоптионс](/javascript/api/excel/excel.pivotlayoutloadoptions)|[autoFormat](/javascript/api/excel/excel.pivotlayoutloadoptions#autoformat)|Указывает, применяется ли форматирование автоматически при его обновлении или перемещении полей|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutloadoptions#preserveformatting)|Указывает, сохраняется ли форматирование при обновлении или пересчете отчета с помощью таких операций, как сведение, сортировка или изменение элементов полей страницы.|
|[Пивотлайаутупдатедата](/javascript/api/excel/excel.pivotlayoutupdatedata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutupdatedata#autoformat)|Указывает, применяется ли форматирование автоматически при его обновлении или перемещении полей|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutupdatedata#preserveformatting)|Указывает, сохраняется ли форматирование при обновлении или пересчете отчета с помощью таких операций, как сведение, сортировка или изменение элементов полей страницы.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Указывает, разрешается ли пользователю изменять значения данных сводной таблицы.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Указывает, используются ли при сортировке в сводной таблице настраиваемые списки.|
|[Пивоттаблеколлектионлоадоптионс](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottablecollectionloadoptions#enabledatavalueediting)|Для каждого элемента в коллекции: указывает, позволяет ли Сводная таблица изменять значения в тексте данных, которые пользователь может редактировать.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottablecollectionloadoptions#usecustomsortlists)|Для каждого элемента в коллекции: указывает, использует ли Сводная таблица настраиваемые списки при сортировке.|
|[Пивоттабледата](/javascript/api/excel/excel.pivottabledata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottabledata#enabledatavalueediting)|Указывает, разрешается ли пользователю изменять значения данных сводной таблицы.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottabledata#usecustomsortlists)|Указывает, используются ли при сортировке в сводной таблице настраиваемые списки.|
|[Пивоттаблелоадоптионс](/javascript/api/excel/excel.pivottableloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableloadoptions#enabledatavalueediting)|Указывает, разрешается ли пользователю изменять значения данных сводной таблицы.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableloadoptions#usecustomsortlists)|Указывает, используются ли при сортировке в сводной таблице настраиваемые списки.|
|[Пивоттаблеупдатедата](/javascript/api/excel/excel.pivottableupdatedata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableupdatedata#enabledatavalueediting)|Указывает, разрешается ли пользователю изменять значения данных сводной таблицы.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableupdatedata#usecustomsortlists)|Указывает, используются ли при сортировке в сводной таблице настраиваемые списки.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Заполняет конечный диапазон из текущего диапазона.|
||[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Заполняет конечный диапазон из текущего диапазона.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Преобразует диапазон ячеек с типами данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Преобразует ячейки диапазона в связанный тип данных на листе.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Копирует данные ячейки или форматирование из исходного диапазона или объекта RangeAreas в текущий диапазон.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Копирует данные ячейки или форматирование из исходного диапазона или объекта RangeAreas в текущий диапазон.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Находит определенную строку на основе указанных условий.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Находит определенную строку на основе указанных условий.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Выполняет мгновенное заполнение текущего диапазона. Функция мгновенного заполнения автоматически подставляет данные, когда обнаруживает закономерность, поэтому диапазон должен состоять из одного столбца со смежными данными, чтобы выявить закономерность.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Возвращает двумерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой ячейки.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждого столбца.  Для свойств, не являющихся одинаковыми в каждой ячейке определенного столбца, возвращается значение null.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Возвращает одномерный массив, в который включены данные для шрифта, заливки, границ, выравнивания и других свойств каждой строки.  Для свойств, не являющихся одинаковыми в каждой ячейке определенной строки, возвращается значение null.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Получает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, представляющих все ячейки, которые соответствуют указанному типу и значению.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Получает объект RangeAreas, состоящий из одного или нескольких прямоугольных диапазонов, представляющих все ячейки, которые соответствуют указанному типу и значению.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Получает объект RangeAreas, состоящий из одного или нескольких диапазонов, представляющих все ячейки, которые соответствуют указанному типу и значению.|
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
||[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Удаляет значения, формат, заливку, границу и т. д. для каждой области, входящей в этот объект RangeAreas.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Удаляет значения, формат, заливку, границу и т. д. для каждой области, входящей в этот объект RangeAreas.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Преобразует все ячейки в объекте RangeAreas с типами данных в текст.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Преобразует все ячейки в объекте RangeAreas в связанный тип данных.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Копирует данные ячейки или форматирование из исходного диапазона или объекта RangeAreas в текущий объект RangeAreas.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Копирует данные ячейки или форматирование из исходного диапазона или объекта RangeAreas в текущий объект RangeAreas.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Возвращает объект RangeAreas, представляющий все столбцы объекта RangeAreas (например, если текущий объект RangeAreas представляет ячейки "B4:E11, H2", возвращается объект RangeAreas, представляющий столбцы "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Возвращает объект RangeAreas, представляющий все строки объекта RangeAreas (например, если текущий объект RangeAreas представляет ячейки "B4:E11", возвращается объект RangeAreas, представляющий строки "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Возвращает объект RangeAreas, представляющий пересечение заданных диапазонов или RangeAreas. Если пересечение не найдено, возвращается сообщение об ошибке ItemNotFound.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Возвращает объект RangeAreas, представляющий пересечение заданных диапазонов или RangeAreas. Если пересечение не найдено, возвращается пустой объект.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Возвращает объект RangeAreas, смещенный на определенное количество строк и столбцов. Измерение возвращаемого объекта RangeAreas будет соответствовать исходному объекту. Если результирующий объект RangeAreas выходит за пределы таблицы листа, возникнет ошибка.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Возвращает объект RangeAreas, представляющий все ячейки, которые соответствуют указанному типу и значению. Выдает ошибку, если не найдено специальных ячеек, соответствующих условиям. |
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Возвращает объект RangeAreas, представляющий все ячейки, которые соответствуют указанному типу и значению. Выдает ошибку, если не найдено специальных ячеек, соответствующих условиям. |
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Возвращает объект RangeAreas, представляющий все ячейки, которые соответствуют указанному типу и значению. Возвращает пустой объект, если не найдено специальных ячеек, соответствующих условиям. |
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
||[Set (Properties: Excel. RangeAreas)](/javascript/api/excel/excel.rangeareas#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Ранжеареасупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.rangeareas#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Устанавливает объект RangeAreas, предназначенный для пересчета при выполнении следующего пересчета.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Представляет стиль всех диапазонов в этом объекте RangeAreas.|
||[track()](/javascript/api/excel/excel.rangeareas#track--)|Отслеживает объект для автоматической корректировки с учетом окружающих изменений в документе. Этот вызов является сокращением для context.trackedObjects.add(thisObject). Если этот объект используется в вызовах .sync и вне последовательного выполнения пакета .run с возникновением ошибки InvalidObjectPath при установке свойства или вызове метода для объекта, необходимо было добавить объект в коллекцию отслеживаемых объектов при первоначальном создании объекта.|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|Освобождает память, связанную с этим объектом, если он ранее отслеживался. Этот вызов является сокращением для context.trackedObjects.remove(thisObject). Наличие большого количества отслеживаемых объектов замедляет ведущее приложение, поэтому не забывайте освобождать любые добавленные объекты после завершения их использования. Перед фактическим освобождением памяти потребуется вызвать метод context.sync().|
|[Ранжеареасдата](/javascript/api/excel/excel.rangeareasdata)|[address](/javascript/api/excel/excel.rangeareasdata#address)|Возвращает ссылку на RageAreas в стиле A1. Значение адреса содержит имя листа для каждого прямоугольного блока или ячейки (например, "Лист1!A1:B4, Лист1!D1:D4"). Только для чтения.|
||[addressLocal](/javascript/api/excel/excel.rangeareasdata#addresslocal)|Возвращает ссылку на RageAreas в языковом стандарте пользователя. Только для чтения.|
||[areaCount](/javascript/api/excel/excel.rangeareasdata#areacount)|Возвращает количество прямоугольных диапазонов, составляющих этот объект RangeAreas.|
||[areas](/javascript/api/excel/excel.rangeareasdata#areas)|Возвращает коллекцию прямоугольных диапазонов, составляющих этот объект RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareasdata#cellcount)|Возвращает число ячеек в объекте RangeAreas с суммированием количества ячеек всех отдельных прямоугольных диапазонов. Возвращает значение -1, если количество ячеек превышает 2^31-1 (2 147 483 647). Только для чтения.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareasdata#conditionalformats)|Возвращает коллекцию объектов ConditionalFormat, пересекающихся с любыми ячейками в этом объекте RangeAreas. Только для чтения.|
||[dataValidation](/javascript/api/excel/excel.rangeareasdata#datavalidation)|Возвращает объект dataValidation для всех диапазонов в объекте RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasdata#format)|Возвращает объект rangeFormat, в который включены шрифт, заливка, границы, выравнивание и другие свойства всех диапазонов в объекте RangeAreas. Только для чтения.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasdata#isentirecolumn)|Указывает, представляют ли все диапазоны в объекте RangeAreas целые столбцы (например, "A:C, Q:Z"). Только для чтения.|
||[isEntireRow](/javascript/api/excel/excel.rangeareasdata#isentirerow)|Указывает, представляют ли все диапазоны в объекте RangeAreas целые строки (например, "1:3, 5:7"). Только для чтения.|
||[style](/javascript/api/excel/excel.rangeareasdata#style)|Представляет стиль всех диапазонов в этом объекте RangeAreas.|
|[Ранжеареаслоадоптионс](/javascript/api/excel/excel.rangeareasloadoptions)|[$all](/javascript/api/excel/excel.rangeareasloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeareasloadoptions#address)|Возвращает ссылку на RageAreas в стиле A1. Значение адреса содержит имя листа для каждого прямоугольного блока или ячейки (например, "Лист1!A1:B4, Лист1!D1:D4"). Только для чтения.|
||[addressLocal](/javascript/api/excel/excel.rangeareasloadoptions#addresslocal)|Возвращает ссылку на RageAreas в языковом стандарте пользователя. Только для чтения.|
||[areaCount](/javascript/api/excel/excel.rangeareasloadoptions#areacount)|Возвращает количество прямоугольных диапазонов, составляющих этот объект RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareasloadoptions#cellcount)|Возвращает число ячеек в объекте RangeAreas с суммированием количества ячеек всех отдельных прямоугольных диапазонов. Возвращает значение -1, если количество ячеек превышает 2^31-1 (2 147 483 647). Только для чтения.|
||[dataValidation](/javascript/api/excel/excel.rangeareasloadoptions#datavalidation)|Возвращает объект dataValidation для всех диапазонов в объекте RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasloadoptions#format)|Возвращает объект rangeFormat, в который включены шрифт, заливка, границы, выравнивание и другие свойства всех диапазонов в объекте RangeAreas.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasloadoptions#isentirecolumn)|Указывает, представляют ли все диапазоны в объекте RangeAreas целые столбцы (например, "A:C, Q:Z"). Только для чтения.|
||[isEntireRow](/javascript/api/excel/excel.rangeareasloadoptions#isentirerow)|Указывает, представляют ли все диапазоны в объекте RangeAreas целые строки (например, "1:3, 5:7"). Только для чтения.|
||[style](/javascript/api/excel/excel.rangeareasloadoptions#style)|Представляет стиль всех диапазонов в этом объекте RangeAreas.|
||[worksheet](/javascript/api/excel/excel.rangeareasloadoptions#worksheet)|Возвращает лист для текущего объекта RangeAreas.|
|[Ранжеареасупдатедата](/javascript/api/excel/excel.rangeareasupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeareasupdatedata#datavalidation)|Возвращает объект dataValidation для всех диапазонов в объекте RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasupdatedata#format)|Возвращает объект rangeFormat, в который включены шрифт, заливка, границы, выравнивание и другие свойства всех диапазонов в объекте RangeAreas.|
||[style](/javascript/api/excel/excel.rangeareasupdatedata#style)|Представляет стиль всех диапазонов в этом объекте RangeAreas.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для границы диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для границ диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжебордерколлектионлоадоптионс](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionloadoptions#tintandshade)|Для каждого элемента в коллекции: Возвращает или задает Double, который осветляет или затемняет цвет границы диапазона, значение находится в пределах от-1 (самая темная) и 1 (самое яркое) с 0 для исходного цвета.|
|[Ранжебордерколлектионупдатедата](/javascript/api/excel/excel.rangebordercollectionupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionupdatedata#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для границ диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжебордердата](/javascript/api/excel/excel.rangeborderdata)|[tintAndShade](/javascript/api/excel/excel.rangeborderdata#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для границы диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжебордерлоадоптионс](/javascript/api/excel/excel.rangeborderloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangeborderloadoptions#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для границы диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжебордерупдатедата](/javascript/api/excel/excel.rangeborderupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangeborderupdatedata#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для границы диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Возвращает количество диапазонов в объекте RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Возвращает объект диапазона в зависимости от его позиции в объекте RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Ранжеколлектионлоадоптионс](/javascript/api/excel/excel.rangecollectionloadoptions)|[$all](/javascript/api/excel/excel.rangecollectionloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangecollectionloadoptions#address)|Для каждого элемента в коллекции: представляет ссылку на диапазон в стиле a1. Значение Address будет содержать ссылку на лист (например, "Лист1! A1: B4). Только для чтения.|
||[addressLocal](/javascript/api/excel/excel.rangecollectionloadoptions#addresslocal)|Для каждого элемента в коллекции: представляет ссылку на диапазон для указанного диапазона на языке пользователя. Только для чтения.|
||[cellCount](/javascript/api/excel/excel.rangecollectionloadoptions#cellcount)|Для каждого элемента в коллекции: количество ячеек в диапазоне. Этот API возвращает значение -1, если количество ячеек превышает 2^31-1 (2,147,483,647). Только для чтения.|
||[Число](/javascript/api/excel/excel.rangecollectionloadoptions#columncount)|Для каждого элемента в коллекции: представляет общее число столбцов в диапазоне. Только для чтения.|
||[columnHidden](/javascript/api/excel/excel.rangecollectionloadoptions#columnhidden)|Для каждого элемента в коллекции: указывает, скрыты ли все столбцы текущего диапазона.|
||[columnIndex](/javascript/api/excel/excel.rangecollectionloadoptions#columnindex)|Для каждого элемента в коллекции — представляет номер столбца первой ячейки в диапазоне. Используется нулевой индекс. Только для чтения.|
||[dataValidation](/javascript/api/excel/excel.rangecollectionloadoptions#datavalidation)|Для каждого элемента в коллекции: Возвращает объект проверки данных.|
||[format](/javascript/api/excel/excel.rangecollectionloadoptions#format)|Для каждого элемента в коллекции: Возвращает объект Format, который инкапсулирует шрифт, заливку, границы, выравнивание и другие свойства диапазона.|
||[formulas](/javascript/api/excel/excel.rangecollectionloadoptions#formulas)|Для каждого элемента в коллекции: представляет формулу в нотации стиля a1.|
||[formulasLocal](/javascript/api/excel/excel.rangecollectionloadoptions#formulaslocal)|Для каждого элемента в коллекции: представляет формулу в нотации стиля a1 в языке пользователя и в языковом стандартном форматировании.  Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[formulasR1C1](/javascript/api/excel/excel.rangecollectionloadoptions#formulasr1c1)|Для каждого элемента в коллекции: представляет формулу в нотации стиля R1C1.|
||[hidden](/javascript/api/excel/excel.rangecollectionloadoptions#hidden)|Для каждого элемента в коллекции: указывает, скрыты ли все ячейки текущего диапазона. Только для чтения.|
||[hyperlink](/javascript/api/excel/excel.rangecollectionloadoptions#hyperlink)|Для каждого элемента в коллекции — представляет гиперссылку для текущего диапазона.|
||[isEntireColumn](/javascript/api/excel/excel.rangecollectionloadoptions#isentirecolumn)|Для каждого элемента в коллекции: указывает, является ли текущий диапазон столбцом целиком. Только для чтения.|
||[isEntireRow](/javascript/api/excel/excel.rangecollectionloadoptions#isentirerow)|Для каждого элемента в коллекции: указывает, является ли текущий диапазон целой строкой. Только для чтения.|
||[linkedDataTypeState](/javascript/api/excel/excel.rangecollectionloadoptions#linkeddatatypestate)|Для каждого элемента в коллекции: представляет состояние типа данных для каждой ячейки. Только для чтения.|
||[numberFormat](/javascript/api/excel/excel.rangecollectionloadoptions#numberformat)|Для каждого элемента в коллекции: представляет код числового формата Excel для заданного диапазона.|
||[numberFormatLocal](/javascript/api/excel/excel.rangecollectionloadoptions#numberformatlocal)|Для каждого элемента в коллекции: представляет код числового формата Excel для заданного диапазона в виде строки на языке пользователя.|
||[Стро](/javascript/api/excel/excel.rangecollectionloadoptions#rowcount)|Для каждого элемента в коллекции: Возвращает общее число строк в диапазоне. Только для чтения.|
||[rowHidden](/javascript/api/excel/excel.rangecollectionloadoptions#rowhidden)|Для каждого элемента в коллекции: указывает, скрыты ли все строки текущего диапазона.|
||[rowIndex](/javascript/api/excel/excel.rangecollectionloadoptions#rowindex)|Для каждого элемента в коллекции: Возвращает номер строки первой ячейки в диапазоне. Используется нулевой индекс. Только для чтения.|
||[style](/javascript/api/excel/excel.rangecollectionloadoptions#style)|Для каждого элемента в коллекции: представляет стиль текущего диапазона.|
||[text](/javascript/api/excel/excel.rangecollectionloadoptions#text)|Для каждого элемента в коллекции: текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.rangecollectionloadoptions#valuetypes)|Для каждого элемента в коллекции: представляет тип данных каждой ячейки. Только для чтения.|
||[values](/javascript/api/excel/excel.rangecollectionloadoptions#values)|Для каждого элемента в коллекции: представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
||[worksheet](/javascript/api/excel/excel.rangecollectionloadoptions#worksheet)|Для каждого элемента в коллекции: лист, содержащий текущий диапазон.|
|[Ранжедата](/javascript/api/excel/excel.rangedata)|[linkedDataTypeState](/javascript/api/excel/excel.rangedata#linkeddatatypestate)|Представляет состояние типа данных каждой ячейки. Только для чтения.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Получает или задает шаблон объекта Range. Дополнительные сведения см. в статье Excel.FillPattern. LinearGradient и RectangularGradient не поддерживаются.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Задает HTML-код, представляющий шаблон объекта Range в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет шаблона для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжефиллдата](/javascript/api/excel/excel.rangefilldata)|[pattern](/javascript/api/excel/excel.rangefilldata#pattern)|Получает или задает шаблон объекта Range. Дополнительные сведения см. в статье Excel.FillPattern. LinearGradient и RectangularGradient не поддерживаются.|
||[patternColor](/javascript/api/excel/excel.rangefilldata#patterncolor)|Задает HTML-код, представляющий шаблон объекта Range в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[patternTintAndShade](/javascript/api/excel/excel.rangefilldata#patterntintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет шаблона для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
||[tintAndShade](/javascript/api/excel/excel.rangefilldata#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжефилллоадоптионс](/javascript/api/excel/excel.rangefillloadoptions)|[pattern](/javascript/api/excel/excel.rangefillloadoptions#pattern)|Получает или задает шаблон объекта Range. Дополнительные сведения см. в статье Excel.FillPattern. LinearGradient и RectangularGradient не поддерживаются.|
||[patternColor](/javascript/api/excel/excel.rangefillloadoptions#patterncolor)|Задает HTML-код, представляющий шаблон объекта Range в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillloadoptions#patterntintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет шаблона для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
||[tintAndShade](/javascript/api/excel/excel.rangefillloadoptions#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжефиллупдатедата](/javascript/api/excel/excel.rangefillupdatedata)|[pattern](/javascript/api/excel/excel.rangefillupdatedata#pattern)|Получает или задает шаблон объекта Range. Дополнительные сведения см. в статье Excel.FillPattern. LinearGradient и RectangularGradient не поддерживаются.|
||[patternColor](/javascript/api/excel/excel.rangefillupdatedata#patterncolor)|Задает HTML-код, представляющий шаблон объекта Range в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillupdatedata#patterntintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет шаблона для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
||[tintAndShade](/javascript/api/excel/excel.rangefillupdatedata#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для заливки диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Указывает, зачеркнут ли шрифт. Значение null указывает, что для всего диапазона не применяется единый параметр зачеркивания.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Указывает, является ли шрифт подстрочным.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Указывает, является ли шрифт надстрочным.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для шрифта диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжефонтдата](/javascript/api/excel/excel.rangefontdata)|[strikethrough](/javascript/api/excel/excel.rangefontdata#strikethrough)|Указывает, зачеркнут ли шрифт. Значение null указывает, что для всего диапазона не применяется единый параметр зачеркивания.|
||[subscript](/javascript/api/excel/excel.rangefontdata#subscript)|Указывает, является ли шрифт подстрочным.|
||[superscript](/javascript/api/excel/excel.rangefontdata#superscript)|Указывает, является ли шрифт надстрочным.|
||[tintAndShade](/javascript/api/excel/excel.rangefontdata#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для шрифта диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжефонтлоадоптионс](/javascript/api/excel/excel.rangefontloadoptions)|[strikethrough](/javascript/api/excel/excel.rangefontloadoptions#strikethrough)|Указывает, зачеркнут ли шрифт. Значение null указывает, что для всего диапазона не применяется единый параметр зачеркивания.|
||[subscript](/javascript/api/excel/excel.rangefontloadoptions#subscript)|Указывает, является ли шрифт подстрочным.|
||[superscript](/javascript/api/excel/excel.rangefontloadoptions#superscript)|Указывает, является ли шрифт надстрочным.|
||[tintAndShade](/javascript/api/excel/excel.rangefontloadoptions#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для шрифта диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[Ранжефонтупдатедата](/javascript/api/excel/excel.rangefontupdatedata)|[strikethrough](/javascript/api/excel/excel.rangefontupdatedata#strikethrough)|Указывает, зачеркнут ли шрифт. Значение null указывает, что для всего диапазона не применяется единый параметр зачеркивания.|
||[subscript](/javascript/api/excel/excel.rangefontupdatedata#subscript)|Указывает, является ли шрифт подстрочным.|
||[superscript](/javascript/api/excel/excel.rangefontupdatedata#superscript)|Указывает, является ли шрифт надстрочным.|
||[tintAndShade](/javascript/api/excel/excel.rangefontupdatedata#tintandshade)|Возвращает или задает значение типа double, осветляющее или затемняющее цвет для шрифта диапазона. Значение: от -1 (самый темный) до 1 (самый светлый). Исходному цвету соответствует значение 0.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста установлено на равномерное распределение.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|Направление чтения для диапазона.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
|[Ранжеформатдата](/javascript/api/excel/excel.rangeformatdata)|[autoIndent](/javascript/api/excel/excel.rangeformatdata#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста установлено на равномерное распределение.|
||[indentLevel](/javascript/api/excel/excel.rangeformatdata#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа.|
||[readingOrder](/javascript/api/excel/excel.rangeformatdata#readingorder)|Направление чтения для диапазона.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatdata#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
|[Ранжеформатлоадоптионс](/javascript/api/excel/excel.rangeformatloadoptions)|[autoIndent](/javascript/api/excel/excel.rangeformatloadoptions#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста установлено на равномерное распределение.|
||[indentLevel](/javascript/api/excel/excel.rangeformatloadoptions#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа.|
||[readingOrder](/javascript/api/excel/excel.rangeformatloadoptions#readingorder)|Направление чтения для диапазона.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatloadoptions#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
|[Ранжеформатупдатедата](/javascript/api/excel/excel.rangeformatupdatedata)|[autoIndent](/javascript/api/excel/excel.rangeformatupdatedata#autoindent)|Указывает, будет ли выполнен автоматический отступ для текста, если выравнивание текста установлено на равномерное распределение.|
||[indentLevel](/javascript/api/excel/excel.rangeformatupdatedata#indentlevel)|Целое число от 0 до 250, указывающее уровень отступа.|
||[readingOrder](/javascript/api/excel/excel.rangeformatupdatedata#readingorder)|Направление чтения для диапазона.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatupdatedata#shrinktofit)|Указывает, сжимается ли автоматически текст для соответствия имеющейся ширине столбца.|
|[Ранжелоадоптионс](/javascript/api/excel/excel.rangeloadoptions)|[linkedDataTypeState](/javascript/api/excel/excel.rangeloadoptions#linkeddatatypestate)|Представляет состояние типа данных каждой ячейки. Только для чтения.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Количество повторяющихся строк, удаленных операцией.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Количество оставшихся уникальных строк, присутствующих в получившемся диапазоне.|
|[Ремоведупликатесресултдата](/javascript/api/excel/excel.removeduplicatesresultdata)|[removed](/javascript/api/excel/excel.removeduplicatesresultdata#removed)|Количество повторяющихся строк, удаленных операцией.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultdata#uniqueremaining)|Количество оставшихся уникальных строк, присутствующих в получившемся диапазоне.|
|[Ремоведупликатесресултлоадоптионс](/javascript/api/excel/excel.removeduplicatesresultloadoptions)|[$all](/javascript/api/excel/excel.removeduplicatesresultloadoptions#$all)||
||[removed](/javascript/api/excel/excel.removeduplicatesresultloadoptions#removed)|Количество повторяющихся строк, удаленных операцией.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultloadoptions#uniqueremaining)|Количество оставшихся уникальных строк, присутствующих в получившемся диапазоне.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Указывает, должно ли совпадение быть полным или частичным. Значение по умолчанию: false (частичное).|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Указывает, учитывается ли регистр при сопоставлении. Значение по умолчанию: false (без учета регистра).|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Представляет свойство `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Представляет свойство `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Представляет свойство `rowIndex`.|
|[Ровпропертиеслоадоптионс](/javascript/api/excel/excel.rowpropertiesloadoptions)|[Format: Excel. Целлпропертиесформатлоадоптионс & {
            rowHeight?] (/жаваскрипт/АПИ/ексцел/ексцел.ровпропертиеслоадоптионс # Format)|Указывает, следует ли загружать `format` свойство.|
||[rowHeight](/javascript/api/excel/excel.rowpropertiesloadoptions#rowheight)||
||[rowHidden](/javascript/api/excel/excel.rowpropertiesloadoptions#rowhidden)|Указывает, следует ли загружать `rowHidden` свойство.|
||[rowIndex](/javascript/api/excel/excel.rowpropertiesloadoptions#rowindex)|Указывает, следует ли загружать `rowIndex` свойство.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Указывает, должно ли совпадение быть полным или частичным. Полное совпадение соответствует всему содержимому ячейки. Значение по умолчанию: false (частичное).|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Указывает, учитывается ли регистр при сопоставлении. Значение по умолчанию: false (без учета регистра).|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Указывает направление поиска. Значение по умолчанию: вперед. См. статью Excel.SearchDirection.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Представляет свойство `format`.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Представляет свойство `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Представляет свойство `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Представляет свойство `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[Format: Excel. Целлпропертиесформат & {
            columnWidth?] (/жаваскрипт/АПИ/ексцел/ексцел.сеттаблеколумнпропертиес # Format)|Представляет свойство `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[Format: Excel. Целлпропертиесформат & {
            rowHeight?] (/жаваскрипт/АПИ/ексцел/ексцел.сеттаблеровпропертиес # Format)|Представляет свойство `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Представляет свойство `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Возвращает или задает замещающий текст описания для объекта Shape.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Возвращает или задает замещающий текст заголовка для объекта Shape.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Удаляет фигуру с листа.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Представляет геометрический тип фигуры. Дополнительные сведения см. в статье Excel.GeometricShapeType. Возвращает значение null, если тип фигуры отличается от GeometricShape.|
||[getAsImage(format: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#getasimage-format-)|Преобразует фигуру в изображение и возвращает изображение в виде строки в кодировке base64. Число точек на дюйм: 96. Единственные поддерживаемые форматы: `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG` и `Excel.PictureFormat.GIF`.|
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
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Масштабирует высоту фигуры с применением указанного коэффициента. Для изображений можно указать изменение масштаба фигуры относительно исходного или текущего размера. Фигуры, не являющиеся изображениями, всегда масштабируются относительно их текущей высоты.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Масштабирует высоту фигуры с применением указанного коэффициента. Для изображений можно указать изменение масштаба фигуры относительно исходного или текущего размера. Фигуры, не являющиеся изображениями, всегда масштабируются относительно их текущей высоты.|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Масштабирует ширину фигуры с применением указанного коэффициента. Для изображений можно указать изменение масштаба фигуры относительно исходного или текущего размера. Фигуры, не являющиеся изображениями, всегда масштабируются относительно их текущей ширины.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Масштабирует ширину фигуры с применением указанного коэффициента. Для изображений можно указать изменение масштаба фигуры относительно исходного или текущего размера. Фигуры, не являющиеся изображениями, всегда масштабируются относительно их текущей ширины.|
||[Set (Properties: Excel. Shape)](/javascript/api/excel/excel.shape#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Шапеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.shape#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[setZOrder(position: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-position-)|Перемещает указанную фигуру вверх или вниз по оси Z в коллекции, что переносит ее вперед или назад относительно других фигур.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Перемещает указанную фигуру вверх или вниз по оси Z в коллекции, что переносит ее вперед или назад относительно других фигур.|
||[top](/javascript/api/excel/excel.shape#top)|Расстояние в пунктах от верхнего края фигуры до верхнего края листа.|
||[visible](/javascript/api/excel/excel.shape#visible)|Представляет видимость фигуры.|
||[width](/javascript/api/excel/excel.shape#width)|Представляет ширину фигуры (в пунктах).|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Получает идентификатор активированной фигуры.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором активирована фигура.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus")](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Добавляет геометрическую фигуру на лист. Возвращает объект Shape, представляющий новую фигуру.|
||[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Добавляет геометрическую фигуру на лист. Возвращает объект Shape, представляющий новую фигуру.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Группирует подмножество фигур на листе этой коллекции. Возвращает объект Shape, представляющий новую группу фигур.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Создает изображение из строки в кодировке base64 и добавляет его на лист. Возвращает объект Shape, представляющий новое изображение.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Добавляет линию на лист. Возвращает объект Shape, представляющий новую линию.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Добавляет линию на лист. Возвращает объект Shape, представляющий новую линию.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Добавляет текстовое поле на лист с указанным текстом в качестве содержимого. Возвращает объект Shape, представляющий новое текстовое поле.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Возвращает количество фигур на листе. Только для чтения.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Получает фигуру по имени или идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Получает фигуру с помощью ее позиции в коллекции.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Шапеколлектионлоадоптионс](/javascript/api/excel/excel.shapecollectionloadoptions)|[$all](/javascript/api/excel/excel.shapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapecollectionloadoptions#alttextdescription)|Для каждого элемента в коллекции: Возвращает или задает текст альтернативного описания для объекта Shape.|
||[altTextTitle](/javascript/api/excel/excel.shapecollectionloadoptions#alttexttitle)|Для каждого элемента в коллекции: Возвращает или задает текст альтернативного заголовка для объекта Shape.|
||[connectionSiteCount](/javascript/api/excel/excel.shapecollectionloadoptions#connectionsitecount)|Для каждого элемента в коллекции: Возвращает число сайтов подключения на этой фигуре. Только для чтения.|
||[fill](/javascript/api/excel/excel.shapecollectionloadoptions#fill)|Для каждого элемента в коллекции: возвращает форматирование заливки данной фигуры.|
||[geometricShape](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshape)|Для каждого элемента в коллекции: возвращает геометрическую фигуру, связанную с фигурой. Если тип фигуры отличается от GeometricShape, возникает ошибка.|
||[geometricShapeType](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshapetype)|Для каждого элемента в коллекции: представляет тип геометрической фигуры для этой геометрической фигуры. Дополнительные сведения см. в статье Excel.GeometricShapeType. Возвращает значение null, если тип фигуры отличается от GeometricShape.|
||[group](/javascript/api/excel/excel.shapecollectionloadoptions#group)|Для каждого элемента в коллекции: Возвращает группу фигур, связанную с фигурой. Если тип фигуры отличается от GroupShape, возникает ошибка.|
||[height](/javascript/api/excel/excel.shapecollectionloadoptions#height)|Для каждого элемента в коллекции: представляет высоту фигуры в пунктах.|
||[id](/javascript/api/excel/excel.shapecollectionloadoptions#id)|Для каждого элемента в коллекции: представляет идентификатор фигуры. Только для чтения.|
||[image](/javascript/api/excel/excel.shapecollectionloadoptions#image)|Для каждого элемента в коллекции: возвращает изображение, связанное с фигурой. Если тип фигуры отличается от Image, возникает ошибка.|
||[left](/javascript/api/excel/excel.shapecollectionloadoptions#left)|Для каждого элемента в коллекции: расстояние (в пунктах) от левой стороны фигуры до левой стороны листа.|
||[level](/javascript/api/excel/excel.shapecollectionloadoptions#level)|Для каждого элемента в коллекции: представляет уровень указанной фигуры. Например, уровень 0 означает, что фигура не является частью групп; уровень 1 означает, что фигура является частью группы верхнего уровня; уровень 2 означает, что фигура является частью подгруппы верхнего уровня.|
||[line](/javascript/api/excel/excel.shapecollectionloadoptions#line)|Для каждого элемента в коллекции: Возвращает строку, связанную с фигурой. Если тип фигуры отличается от Line, возникает ошибка.|
||[lineFormat](/javascript/api/excel/excel.shapecollectionloadoptions#lineformat)|Для каждого элемента в коллекции: возвращает форматирование строки этой фигуры.|
||[lockAspectRatio](/javascript/api/excel/excel.shapecollectionloadoptions#lockaspectratio)|Для каждого элемента в коллекции: указывает, заблокировано ли пропорции данной фигуры.|
||[name](/javascript/api/excel/excel.shapecollectionloadoptions#name)|Для каждого элемента в коллекции: представляет имя фигуры.|
||[parentGroup](/javascript/api/excel/excel.shapecollectionloadoptions#parentgroup)|Для каждого элемента в коллекции: представляет родительскую группу этой фигуры.|
||[rotation](/javascript/api/excel/excel.shapecollectionloadoptions#rotation)|Для каждого элемента в коллекции — представляет Поворот фигуры в градусах.|
||[textFrame](/javascript/api/excel/excel.shapecollectionloadoptions#textframe)|Для каждого элемента в коллекции: Возвращает объект текстового фрейма этой фигуры. Только для чтения.|
||[top](/javascript/api/excel/excel.shapecollectionloadoptions#top)|Для каждого элемента в коллекции: расстояние (в пунктах) от верхнего края фигуры до верхнего края листа.|
||[type](/javascript/api/excel/excel.shapecollectionloadoptions#type)|Для каждого элемента в коллекции: Возвращает тип этой фигуры. Дополнительные сведения см. в статье Excel.ShapeType. Только для чтения.|
||[visible](/javascript/api/excel/excel.shapecollectionloadoptions#visible)|Для каждого элемента в коллекции: представляет видимость этой фигуры.|
||[width](/javascript/api/excel/excel.shapecollectionloadoptions#width)|Для каждого элемента в коллекции: представляет ширину фигуры в пунктах.|
||[zOrderPosition](/javascript/api/excel/excel.shapecollectionloadoptions#zorderposition)|Для каждого элемента в коллекции: Возвращает позицию указанной фигуры в z-порядке, где 0 представляет нижнюю часть стека заказов. Только для чтения.|
|[Шапедата](/javascript/api/excel/excel.shapedata)|[altTextDescription](/javascript/api/excel/excel.shapedata#alttextdescription)|Возвращает или задает замещающий текст описания для объекта Shape.|
||[altTextTitle](/javascript/api/excel/excel.shapedata#alttexttitle)|Возвращает или задает замещающий текст заголовка для объекта Shape.|
||[connectionSiteCount](/javascript/api/excel/excel.shapedata#connectionsitecount)|Возвращает количество точек соединения на фигуре. Только для чтения.|
||[fill](/javascript/api/excel/excel.shapedata#fill)|Возвращает формат заливки фигуры. Только для чтения.|
||[geometricShapeType](/javascript/api/excel/excel.shapedata#geometricshapetype)|Представляет геометрический тип фигуры. Дополнительные сведения см. в статье Excel.GeometricShapeType. Возвращает значение null, если тип фигуры отличается от GeometricShape.|
||[height](/javascript/api/excel/excel.shapedata#height)|Представляет высоту фигуры (в пунктах).|
||[id](/javascript/api/excel/excel.shapedata#id)|Представляет идентификатор фигуры. Только для чтения.|
||[left](/javascript/api/excel/excel.shapedata#left)|Расстояние в пунктах от левого края фигуры до левого края листа.|
||[level](/javascript/api/excel/excel.shapedata#level)|Представляет уровень указанной фигуры. Например, уровень 0 означает, что фигура не является частью групп; уровень 1 означает, что фигура является частью группы верхнего уровня; уровень 2 означает, что фигура является частью подгруппы верхнего уровня.|
||[lineFormat](/javascript/api/excel/excel.shapedata#lineformat)|Возвращает формат линии для фигуры. Только для чтения.|
||[lockAspectRatio](/javascript/api/excel/excel.shapedata#lockaspectratio)|Указывает, заблокированы ли пропорции фигуры.|
||[name](/javascript/api/excel/excel.shapedata#name)|Представляет название фигуры.|
||[rotation](/javascript/api/excel/excel.shapedata#rotation)|Представляет поворот фигуры в градусах.|
||[top](/javascript/api/excel/excel.shapedata#top)|Расстояние в пунктах от верхнего края фигуры до верхнего края листа.|
||[type](/javascript/api/excel/excel.shapedata#type)|Возвращает тип фигуры. Дополнительные сведения см. в статье Excel.ShapeType. Только для чтения.|
||[visible](/javascript/api/excel/excel.shapedata#visible)|Представляет видимость фигуры.|
||[width](/javascript/api/excel/excel.shapedata#width)|Представляет ширину фигуры (в пунктах).|
||[zOrderPosition](/javascript/api/excel/excel.shapedata#zorderposition)|Возвращает положение указанной фигуры по оси Z. Значение 0 представляет нижнее положение по оси. Только для чтения.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Получает идентификатор деактивированной фигуры.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Получает идентификатор листа, в котором деактивирована фигура.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Очищает формат заливки фигуры.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Представляет цвет переднего плана заливки фигуры в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[type](/javascript/api/excel/excel.shapefill#type)|Возвращает тип заливки фигуры. Только для чтения. Дополнительные сведения см. в статье Excel.ShapeFillType.|
||[Set (Properties: Excel. Шапефилл)](/javascript/api/excel/excel.shapefill#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Шапефиллупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.shapefill#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Задает заливку одним цветом для фигуры. При этом тип заливки изменяется на сплошную.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Возвращает или задает процентное значение прозрачности заливки как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если тип фигуры не поддерживает прозрачность или заливка фигуры имеет несогласованную прозрачность, например при использовании градиентной заливки.|
|[Шапефиллдата](/javascript/api/excel/excel.shapefilldata)|[foregroundColor](/javascript/api/excel/excel.shapefilldata#foregroundcolor)|Представляет цвет переднего плана заливки фигуры в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[transparency](/javascript/api/excel/excel.shapefilldata#transparency)|Возвращает или задает процентное значение прозрачности заливки как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если тип фигуры не поддерживает прозрачность или заливка фигуры имеет несогласованную прозрачность, например при использовании градиентной заливки.|
||[type](/javascript/api/excel/excel.shapefilldata#type)|Возвращает тип заливки фигуры. Только для чтения. Дополнительные сведения см. в статье Excel.ShapeFillType.|
|[Шапефилллоадоптионс](/javascript/api/excel/excel.shapefillloadoptions)|[$all](/javascript/api/excel/excel.shapefillloadoptions#$all)||
||[foregroundColor](/javascript/api/excel/excel.shapefillloadoptions#foregroundcolor)|Представляет цвет переднего плана заливки фигуры в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[transparency](/javascript/api/excel/excel.shapefillloadoptions#transparency)|Возвращает или задает процентное значение прозрачности заливки как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если тип фигуры не поддерживает прозрачность или заливка фигуры имеет несогласованную прозрачность, например при использовании градиентной заливки.|
||[type](/javascript/api/excel/excel.shapefillloadoptions#type)|Возвращает тип заливки фигуры. Только для чтения. Дополнительные сведения см. в статье Excel.ShapeFillType.|
|[Шапефиллупдатедата](/javascript/api/excel/excel.shapefillupdatedata)|[foregroundColor](/javascript/api/excel/excel.shapefillupdatedata#foregroundcolor)|Представляет цвет переднего плана заливки фигуры в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[transparency](/javascript/api/excel/excel.shapefillupdatedata#transparency)|Возвращает или задает процентное значение прозрачности заливки как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если тип фигуры не поддерживает прозрачность или заливка фигуры имеет несогласованную прозрачность, например при использовании градиентной заливки.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Указывает, является ли шрифт полужирным. Возвращает значение null, если объект TextRange включает фрагменты как с полужирным, так и без полужирного текста.|
||[color](/javascript/api/excel/excel.shapefont#color)|HTML-код цвета текста (например, значение #FF0000 обозначает красный цвет). Возвращает значение null, если объект TextRange включает фрагменты текста с разными цветами.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Указывает, применяется ли курсив. Возвращает значение null, если объект TextRange включает фрагменты текста как выделенные, так и не выделенные курсивом.|
||[name](/javascript/api/excel/excel.shapefont#name)|Представляет имя шрифта (например, Calibri). Если текст является набором сложных знаков или написан на восточноазиатских языках, этот параметр является соответствующим именем шрифта. В противном случае это имя шрифта на латинице.|
||[Set (Properties: Excel. Шапефонт)](/javascript/api/excel/excel.shapefont#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Шапефонтупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.shapefont#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[size](/javascript/api/excel/excel.shapefont#size)|Представляет размер шрифта в пунктах (например, 11). Возвращает значение null, если объект TextRange включает фрагменты текста с разными размерами шрифта.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Тип подчеркивания, применяемый для шрифта. Возвращает значение null, если объект TextRange включает фрагменты текста с разными стилями подчеркивания. Дополнительные сведения см. в статье Excel.ShapeFontUnderlineStyle.|
|[Шапефонтдата](/javascript/api/excel/excel.shapefontdata)|[bold](/javascript/api/excel/excel.shapefontdata#bold)|Указывает, является ли шрифт полужирным. Возвращает значение null, если объект TextRange включает фрагменты как с полужирным, так и без полужирного текста.|
||[color](/javascript/api/excel/excel.shapefontdata#color)|HTML-код цвета текста (например, значение #FF0000 обозначает красный цвет). Возвращает значение null, если объект TextRange включает фрагменты текста с разными цветами.|
||[italic](/javascript/api/excel/excel.shapefontdata#italic)|Указывает, применяется ли курсив. Возвращает значение null, если объект TextRange включает фрагменты текста как выделенные, так и не выделенные курсивом.|
||[name](/javascript/api/excel/excel.shapefontdata#name)|Представляет имя шрифта (например, Calibri). Если текст является набором сложных знаков или написан на восточноазиатских языках, этот параметр является соответствующим именем шрифта. В противном случае это имя шрифта на латинице.|
||[size](/javascript/api/excel/excel.shapefontdata#size)|Представляет размер шрифта в пунктах (например, 11). Возвращает значение null, если объект TextRange включает фрагменты текста с разными размерами шрифта.|
||[underline](/javascript/api/excel/excel.shapefontdata#underline)|Тип подчеркивания, применяемый для шрифта. Возвращает значение null, если объект TextRange включает фрагменты текста с разными стилями подчеркивания. Дополнительные сведения см. в статье Excel.ShapeFontUnderlineStyle.|
|[Шапефонтлоадоптионс](/javascript/api/excel/excel.shapefontloadoptions)|[$all](/javascript/api/excel/excel.shapefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.shapefontloadoptions#bold)|Указывает, является ли шрифт полужирным. Возвращает значение null, если объект TextRange включает фрагменты как с полужирным, так и без полужирного текста.|
||[color](/javascript/api/excel/excel.shapefontloadoptions#color)|HTML-код цвета текста (например, значение #FF0000 обозначает красный цвет). Возвращает значение null, если объект TextRange включает фрагменты текста с разными цветами.|
||[italic](/javascript/api/excel/excel.shapefontloadoptions#italic)|Указывает, применяется ли курсив. Возвращает значение null, если объект TextRange включает фрагменты текста как выделенные, так и не выделенные курсивом.|
||[name](/javascript/api/excel/excel.shapefontloadoptions#name)|Представляет имя шрифта (например, Calibri). Если текст является набором сложных знаков или написан на восточноазиатских языках, этот параметр является соответствующим именем шрифта. В противном случае это имя шрифта на латинице.|
||[size](/javascript/api/excel/excel.shapefontloadoptions#size)|Представляет размер шрифта в пунктах (например, 11). Возвращает значение null, если объект TextRange включает фрагменты текста с разными размерами шрифта.|
||[underline](/javascript/api/excel/excel.shapefontloadoptions#underline)|Тип подчеркивания, применяемый для шрифта. Возвращает значение null, если объект TextRange включает фрагменты текста с разными стилями подчеркивания. Дополнительные сведения см. в статье Excel.ShapeFontUnderlineStyle.|
|[Шапефонтупдатедата](/javascript/api/excel/excel.shapefontupdatedata)|[bold](/javascript/api/excel/excel.shapefontupdatedata#bold)|Указывает, является ли шрифт полужирным. Возвращает значение null, если объект TextRange включает фрагменты как с полужирным, так и без полужирного текста.|
||[color](/javascript/api/excel/excel.shapefontupdatedata#color)|HTML-код цвета текста (например, значение #FF0000 обозначает красный цвет). Возвращает значение null, если объект TextRange включает фрагменты текста с разными цветами.|
||[italic](/javascript/api/excel/excel.shapefontupdatedata#italic)|Указывает, применяется ли курсив. Возвращает значение null, если объект TextRange включает фрагменты текста как выделенные, так и не выделенные курсивом.|
||[name](/javascript/api/excel/excel.shapefontupdatedata#name)|Представляет имя шрифта (например, Calibri). Если текст является набором сложных знаков или написан на восточноазиатских языках, этот параметр является соответствующим именем шрифта. В противном случае это имя шрифта на латинице.|
||[size](/javascript/api/excel/excel.shapefontupdatedata#size)|Представляет размер шрифта в пунктах (например, 11). Возвращает значение null, если объект TextRange включает фрагменты текста с разными размерами шрифта.|
||[underline](/javascript/api/excel/excel.shapefontupdatedata#underline)|Тип подчеркивания, применяемый для шрифта. Возвращает значение null, если объект TextRange включает фрагменты текста с разными стилями подчеркивания. Дополнительные сведения см. в статье Excel.ShapeFontUnderlineStyle.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Представляет идентификатор фигуры. Только для чтения.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Возвращает объект Shape, связанный с группой. Только для чтения.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Возвращает коллекцию объектов Shape. Только для чтения.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Отменяет группировку любых сгруппированных фигур в указанной группе фигур.|
|[Шапеграупдата](/javascript/api/excel/excel.shapegroupdata)|[id](/javascript/api/excel/excel.shapegroupdata#id)|Представляет идентификатор фигуры. Только для чтения.|
||[shapes](/javascript/api/excel/excel.shapegroupdata#shapes)|Возвращает коллекцию объектов Shape. Только для чтения.|
|[Шапеграуплоадоптионс](/javascript/api/excel/excel.shapegrouploadoptions)|[$all](/javascript/api/excel/excel.shapegrouploadoptions#$all)||
||[id](/javascript/api/excel/excel.shapegrouploadoptions#id)|Представляет идентификатор фигуры. Только для чтения.|
||[shape](/javascript/api/excel/excel.shapegrouploadoptions#shape)|Возвращает объект Shape, связанный с группой.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Представляет цвет линии в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные типы штриха. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[Set (Properties: Excel. Шапелинеформат)](/javascript/api/excel/excel.shapelineformat#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Шапелинеформатупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.shapelineformat#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные стили. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Представляет степень прозрачности указанной линии как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если в фигуре используются несогласованные параметры прозрачности.|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Указывает, отображается ли форматирование линии элемента фигуры. Возвращает значение null, если в фигуре используются несогласованные параметры видимости.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Представляет толщину линии (в пунктах). Возвращает значение null, если линия является невидимой или используются линии с несогласованной толщиной.|
|[Шапелинеформатдата](/javascript/api/excel/excel.shapelineformatdata)|[color](/javascript/api/excel/excel.shapelineformatdata#color)|Представляет цвет линии в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[dashStyle](/javascript/api/excel/excel.shapelineformatdata#dashstyle)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные типы штриха. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformatdata#style)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные стили. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformatdata#transparency)|Представляет степень прозрачности указанной линии как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если в фигуре используются несогласованные параметры прозрачности.|
||[visible](/javascript/api/excel/excel.shapelineformatdata#visible)|Указывает, отображается ли форматирование линии элемента фигуры. Возвращает значение null, если в фигуре используются несогласованные параметры видимости.|
||[weight](/javascript/api/excel/excel.shapelineformatdata#weight)|Представляет толщину линии (в пунктах). Возвращает значение null, если линия является невидимой или используются линии с несогласованной толщиной.|
|[Шапелинеформатлоадоптионс](/javascript/api/excel/excel.shapelineformatloadoptions)|[$all](/javascript/api/excel/excel.shapelineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.shapelineformatloadoptions#color)|Представляет цвет линии в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[dashStyle](/javascript/api/excel/excel.shapelineformatloadoptions#dashstyle)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные типы штриха. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformatloadoptions#style)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные стили. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformatloadoptions#transparency)|Представляет степень прозрачности указанной линии как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если в фигуре используются несогласованные параметры прозрачности.|
||[visible](/javascript/api/excel/excel.shapelineformatloadoptions#visible)|Указывает, отображается ли форматирование линии элемента фигуры. Возвращает значение null, если в фигуре используются несогласованные параметры видимости.|
||[weight](/javascript/api/excel/excel.shapelineformatloadoptions#weight)|Представляет толщину линии (в пунктах). Возвращает значение null, если линия является невидимой или используются линии с несогласованной толщиной.|
|[Шапелинеформатупдатедата](/javascript/api/excel/excel.shapelineformatupdatedata)|[color](/javascript/api/excel/excel.shapelineformatupdatedata#color)|Представляет цвет линии в формате HTML в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|
||[dashStyle](/javascript/api/excel/excel.shapelineformatupdatedata#dashstyle)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные типы штриха. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformatupdatedata#style)|Представляет тип линии фигуры. Возвращает значение null, если линия является невидимой или используются несогласованные стили. Дополнительные сведения см. в статье Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformatupdatedata#transparency)|Представляет степень прозрачности указанной линии как значение от 0,0 (непрозрачная) до 1,0 (полностью прозрачная). Возвращает значение null, если в фигуре используются несогласованные параметры прозрачности.|
||[visible](/javascript/api/excel/excel.shapelineformatupdatedata#visible)|Указывает, отображается ли форматирование линии элемента фигуры. Возвращает значение null, если в фигуре используются несогласованные параметры видимости.|
||[weight](/javascript/api/excel/excel.shapelineformatupdatedata#weight)|Представляет толщину линии (в пунктах). Возвращает значение null, если линия является невидимой или используются линии с несогласованной толщиной.|
|[Шапелоадоптионс](/javascript/api/excel/excel.shapeloadoptions)|[$all](/javascript/api/excel/excel.shapeloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapeloadoptions#alttextdescription)|Возвращает или задает замещающий текст описания для объекта Shape.|
||[altTextTitle](/javascript/api/excel/excel.shapeloadoptions#alttexttitle)|Возвращает или задает замещающий текст заголовка для объекта Shape.|
||[connectionSiteCount](/javascript/api/excel/excel.shapeloadoptions#connectionsitecount)|Возвращает количество точек соединения на фигуре. Только для чтения.|
||[fill](/javascript/api/excel/excel.shapeloadoptions#fill)|Возвращает формат заливки фигуры.|
||[geometricShape](/javascript/api/excel/excel.shapeloadoptions#geometricshape)|Возвращает геометрическую фигуру, связанную с линией. Если тип фигуры отличается от GeometricShape, возникает ошибка.|
||[geometricShapeType](/javascript/api/excel/excel.shapeloadoptions#geometricshapetype)|Представляет геометрический тип фигуры. Дополнительные сведения см. в статье Excel.GeometricShapeType. Возвращает значение null, если тип фигуры отличается от GeometricShape.|
||[group](/javascript/api/excel/excel.shapeloadoptions#group)|Возвращает группу фигур, связанную с фигурой. Если тип фигуры отличается от GroupShape, возникает ошибка.|
||[height](/javascript/api/excel/excel.shapeloadoptions#height)|Представляет высоту фигуры (в пунктах).|
||[id](/javascript/api/excel/excel.shapeloadoptions#id)|Представляет идентификатор фигуры. Только для чтения.|
||[image](/javascript/api/excel/excel.shapeloadoptions#image)|Возвращает изображение, связанное с фигурой. Если тип фигуры отличается от Image, возникает ошибка.|
||[left](/javascript/api/excel/excel.shapeloadoptions#left)|Расстояние в пунктах от левого края фигуры до левого края листа.|
||[level](/javascript/api/excel/excel.shapeloadoptions#level)|Представляет уровень указанной фигуры. Например, уровень 0 означает, что фигура не является частью групп; уровень 1 означает, что фигура является частью группы верхнего уровня; уровень 2 означает, что фигура является частью подгруппы верхнего уровня.|
||[line](/javascript/api/excel/excel.shapeloadoptions#line)|Возвращает линию, связанную с фигурой. Если тип фигуры отличается от Line, возникает ошибка.|
||[lineFormat](/javascript/api/excel/excel.shapeloadoptions#lineformat)|Возвращает формат линии для фигуры.|
||[lockAspectRatio](/javascript/api/excel/excel.shapeloadoptions#lockaspectratio)|Указывает, заблокированы ли пропорции фигуры.|
||[name](/javascript/api/excel/excel.shapeloadoptions#name)|Представляет название фигуры.|
||[parentGroup](/javascript/api/excel/excel.shapeloadoptions#parentgroup)|Представляет родительскую группу фигуры.|
||[rotation](/javascript/api/excel/excel.shapeloadoptions#rotation)|Представляет поворот фигуры в градусах.|
||[textFrame](/javascript/api/excel/excel.shapeloadoptions#textframe)|Возвращает объект рамки с текстом для фигуры. Только для чтения.|
||[top](/javascript/api/excel/excel.shapeloadoptions#top)|Расстояние в пунктах от верхнего края фигуры до верхнего края листа.|
||[type](/javascript/api/excel/excel.shapeloadoptions#type)|Возвращает тип фигуры. Дополнительные сведения см. в статье Excel.ShapeType. Только для чтения.|
||[visible](/javascript/api/excel/excel.shapeloadoptions#visible)|Представляет видимость фигуры.|
||[width](/javascript/api/excel/excel.shapeloadoptions#width)|Представляет ширину фигуры (в пунктах).|
||[zOrderPosition](/javascript/api/excel/excel.shapeloadoptions#zorderposition)|Возвращает положение указанной фигуры по оси Z. Значение 0 представляет нижнее положение по оси. Только для чтения.|
|[Шапеупдатедата](/javascript/api/excel/excel.shapeupdatedata)|[altTextDescription](/javascript/api/excel/excel.shapeupdatedata#alttextdescription)|Возвращает или задает замещающий текст описания для объекта Shape.|
||[altTextTitle](/javascript/api/excel/excel.shapeupdatedata#alttexttitle)|Возвращает или задает замещающий текст заголовка для объекта Shape.|
||[fill](/javascript/api/excel/excel.shapeupdatedata#fill)|Возвращает формат заливки фигуры.|
||[geometricShapeType](/javascript/api/excel/excel.shapeupdatedata#geometricshapetype)|Представляет геометрический тип фигуры. Дополнительные сведения см. в статье Excel.GeometricShapeType. Возвращает значение null, если тип фигуры отличается от GeometricShape.|
||[height](/javascript/api/excel/excel.shapeupdatedata#height)|Представляет высоту фигуры (в пунктах).|
||[left](/javascript/api/excel/excel.shapeupdatedata#left)|Расстояние в пунктах от левого края фигуры до левого края листа.|
||[lineFormat](/javascript/api/excel/excel.shapeupdatedata#lineformat)|Возвращает формат линии для фигуры.|
||[lockAspectRatio](/javascript/api/excel/excel.shapeupdatedata#lockaspectratio)|Указывает, заблокированы ли пропорции фигуры.|
||[name](/javascript/api/excel/excel.shapeupdatedata#name)|Представляет название фигуры.|
||[rotation](/javascript/api/excel/excel.shapeupdatedata#rotation)|Представляет поворот фигуры в градусах.|
||[top](/javascript/api/excel/excel.shapeupdatedata#top)|Расстояние в пунктах от верхнего края фигуры до верхнего края листа.|
||[visible](/javascript/api/excel/excel.shapeupdatedata#visible)|Представляет видимость фигуры.|
||[width](/javascript/api/excel/excel.shapeupdatedata#width)|Представляет ширину фигуры (в пунктах).|
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
|[Таблеколлектионлоадоптионс](/javascript/api/excel/excel.tablecollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.tablecollectionloadoptions#autofilter)|Для каждого элемента в коллекции: представляет объект автофильтра таблицы.|
|[TableData](/javascript/api/excel/excel.tabledata)|[autoFilter](/javascript/api/excel/excel.tabledata#autofilter)|Представляет объект AutoFilter таблицы. Только для чтения.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Указывает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Указывает идентификатор удаленной таблицы.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Указывает имя удаленной таблицы.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Указывает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Указывает идентификатор листа, в котором удаляется таблица.|
|[Таблелоадоптионс](/javascript/api/excel/excel.tableloadoptions)|[autoFilter](/javascript/api/excel/excel.tableloadoptions#autofilter)|Представляет объект AutoFilter таблицы.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Получает количество таблиц в коллекции.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Получает первую таблицу в коллекции. Таблицы в коллекции сортируются сверху вниз и слева направо, поэтому верхняя левая таблица является первой в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Получает таблицу по имени или идентификатору.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Таблескопедколлектионлоадоптионс](/javascript/api/excel/excel.tablescopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablescopedcollectionloadoptions#$all)||
||[autoFilter](/javascript/api/excel/excel.tablescopedcollectionloadoptions#autofilter)|Для каждого элемента в коллекции: представляет объект автофильтра таблицы.|
||[столбцы](/javascript/api/excel/excel.tablescopedcollectionloadoptions#columns)|Для каждого элемента в коллекции: представляет коллекцию всех столбцов в таблице.|
||[highlightFirstColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightfirstcolumn)|Для каждого элемента в коллекции: указывает, содержит ли первый столбец специальное форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightlastcolumn)|Для каждого элемента в коллекции: указывает, содержит ли последний столбец специальное форматирование.|
||[id](/javascript/api/excel/excel.tablescopedcollectionloadoptions#id)|Для каждого элемента в коллекции: Возвращает значение, однозначно идентифицирующее таблицу в заданной книге. Значение идентификатора остается прежним, даже если переименовать таблицу. Только для чтения.|
||[legacyId](/javascript/api/excel/excel.tablescopedcollectionloadoptions#legacyid)|Для каждого элемента в коллекции: Возвращает числовой идентификатор.|
||[name](/javascript/api/excel/excel.tablescopedcollectionloadoptions#name)|Для каждого элемента в коллекции: имя таблицы.|
||[строки](/javascript/api/excel/excel.tablescopedcollectionloadoptions#rows)|Для каждого элемента в коллекции: представляет коллекцию всех строк в таблице.|
||[showBandedColumns](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedcolumns)|Для каждого элемента в коллекции: указывает, отображаются ли в столбцах полоснее форматирование, в результате которой нечетные столбцы выделяются не так, как даже для упрощения чтения таблицы.|
||[showBandedRows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedrows)|Для каждого элемента в коллекции: указывает, отображаются ли в строках форматирование с чередованием, в результате чего нечетные строки выделяются иначе, чтобы упростить чтение таблицы.|
||[showFilterButton](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showfilterbutton)|Для каждого элемента в коллекции: указывает, отображаются ли кнопки фильтра в верхней части каждого заголовка столбца. Это свойство можно использовать, только если таблица содержит строку заголовков.|
||[Шовхеадерс](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showheaders)|Для каждого элемента в коллекции: указывает, видима ли строка заголовков. Можно задать это значение, чтобы отобразить или скрыть строку заголовков.|
||[Шовтоталс](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showtotals)|Для каждого элемента в коллекции: указывает, видима ли строка итогов. Можно задать это значение, чтобы отобразить или скрыть строку итогов.|
||[sort](/javascript/api/excel/excel.tablescopedcollectionloadoptions#sort)|Для каждого элемента в коллекции: представляет сортировку для таблицы.|
||[style](/javascript/api/excel/excel.tablescopedcollectionloadoptions#style)|Для каждого элемента в коллекции: значение константы, представляющее стиль таблицы. Возможные значения: от TableStyleLight1 до TableStyleLight21, от TableStyleMedium1 до TableStyleMedium28, от TableStyleStyleDark1 до TableStyleStyleDark11. Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
||[worksheet](/javascript/api/excel/excel.tablescopedcollectionloadoptions#worksheet)|Для каждого элемента в коллекции: лист, содержащий текущую таблицу.|
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
||[Set (Properties: Excel. TextFrame)](/javascript/api/excel/excel.textframe#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Текстфрамеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.textframe#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Представляет вертикальное выравнивание для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalAlignment.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Представляет действие вертикального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalOverflow.|
|[Текстфрамедата](/javascript/api/excel/excel.textframedata)|[autoSizeSetting](/javascript/api/excel/excel.textframedata#autosizesetting)|Возвращает или задает параметры автоматического подбора размера для рамки с текстом. Для рамки с текстом можно настроить автоматический подбор размера текста в соответствии с размером рамки, автоматический подбор размера рамки в соответствии с содержимым или не выполнять автоматический подбор размера.|
||[bottomMargin](/javascript/api/excel/excel.textframedata#bottommargin)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[hasText](/javascript/api/excel/excel.textframedata#hastext)|Указывает, содержится ли в текстовой рамке текст.|
||[horizontalAlignment](/javascript/api/excel/excel.textframedata#horizontalalignment)|Представляет горизонтальное выравнивание рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextHorizontalAlignment.|
||[horizontalOverflow](/javascript/api/excel/excel.textframedata#horizontaloverflow)|Представляет действие горизонтального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextHorizontalOverflow.|
||[leftMargin](/javascript/api/excel/excel.textframedata#leftmargin)|Представляет левое поле рамки с текстом (в пунктах).|
||[orientation](/javascript/api/excel/excel.textframedata#orientation)|Представляет ориентацию текста для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextOrientation.|
||[readingOrder](/javascript/api/excel/excel.textframedata#readingorder)|Представляет направление чтения рамки с текстом (слева направо или справа налево). Дополнительные сведения см. в статье Excel.ShapeTextReadingOrder.|
||[rightMargin](/javascript/api/excel/excel.textframedata#rightmargin)|Представляет правое поле рамки с текстом (в пунктах).|
||[topMargin](/javascript/api/excel/excel.textframedata#topmargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.textframedata#verticalalignment)|Представляет вертикальное выравнивание для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalAlignment.|
||[verticalOverflow](/javascript/api/excel/excel.textframedata#verticaloverflow)|Представляет действие вертикального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalOverflow.|
|[Текстфрамелоадоптионс](/javascript/api/excel/excel.textframeloadoptions)|[$all](/javascript/api/excel/excel.textframeloadoptions#$all)||
||[autoSizeSetting](/javascript/api/excel/excel.textframeloadoptions#autosizesetting)|Возвращает или задает параметры автоматического подбора размера для рамки с текстом. Для рамки с текстом можно настроить автоматический подбор размера текста в соответствии с размером рамки, автоматический подбор размера рамки в соответствии с содержимым или не выполнять автоматический подбор размера.|
||[bottomMargin](/javascript/api/excel/excel.textframeloadoptions#bottommargin)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[hasText](/javascript/api/excel/excel.textframeloadoptions#hastext)|Указывает, содержится ли в текстовой рамке текст.|
||[horizontalAlignment](/javascript/api/excel/excel.textframeloadoptions#horizontalalignment)|Представляет горизонтальное выравнивание рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextHorizontalAlignment.|
||[horizontalOverflow](/javascript/api/excel/excel.textframeloadoptions#horizontaloverflow)|Представляет действие горизонтального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextHorizontalOverflow.|
||[leftMargin](/javascript/api/excel/excel.textframeloadoptions#leftmargin)|Представляет левое поле рамки с текстом (в пунктах).|
||[orientation](/javascript/api/excel/excel.textframeloadoptions#orientation)|Представляет ориентацию текста для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextOrientation.|
||[readingOrder](/javascript/api/excel/excel.textframeloadoptions#readingorder)|Представляет направление чтения рамки с текстом (слева направо или справа налево). Дополнительные сведения см. в статье Excel.ShapeTextReadingOrder.|
||[rightMargin](/javascript/api/excel/excel.textframeloadoptions#rightmargin)|Представляет правое поле рамки с текстом (в пунктах).|
||[textRange](/javascript/api/excel/excel.textframeloadoptions#textrange)|Представляет текст, присоединенный к фигуре в текстовой рамке, а также свойства и методы для операций с текстом. Дополнительные сведения см. в статье Excel.TextRange.|
||[topMargin](/javascript/api/excel/excel.textframeloadoptions#topmargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.textframeloadoptions#verticalalignment)|Представляет вертикальное выравнивание для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalAlignment.|
||[verticalOverflow](/javascript/api/excel/excel.textframeloadoptions#verticaloverflow)|Представляет действие вертикального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalOverflow.|
|[Текстфрамеупдатедата](/javascript/api/excel/excel.textframeupdatedata)|[autoSizeSetting](/javascript/api/excel/excel.textframeupdatedata#autosizesetting)|Возвращает или задает параметры автоматического подбора размера для рамки с текстом. Для рамки с текстом можно настроить автоматический подбор размера текста в соответствии с размером рамки, автоматический подбор размера рамки в соответствии с содержимым или не выполнять автоматический подбор размера.|
||[bottomMargin](/javascript/api/excel/excel.textframeupdatedata#bottommargin)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[horizontalAlignment](/javascript/api/excel/excel.textframeupdatedata#horizontalalignment)|Представляет горизонтальное выравнивание рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextHorizontalAlignment.|
||[horizontalOverflow](/javascript/api/excel/excel.textframeupdatedata#horizontaloverflow)|Представляет действие горизонтального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextHorizontalOverflow.|
||[leftMargin](/javascript/api/excel/excel.textframeupdatedata#leftmargin)|Представляет левое поле рамки с текстом (в пунктах).|
||[orientation](/javascript/api/excel/excel.textframeupdatedata#orientation)|Представляет ориентацию текста для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextOrientation.|
||[readingOrder](/javascript/api/excel/excel.textframeupdatedata#readingorder)|Представляет направление чтения рамки с текстом (слева направо или справа налево). Дополнительные сведения см. в статье Excel.ShapeTextReadingOrder.|
||[rightMargin](/javascript/api/excel/excel.textframeupdatedata#rightmargin)|Представляет правое поле рамки с текстом (в пунктах).|
||[topMargin](/javascript/api/excel/excel.textframeupdatedata#topmargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.textframeupdatedata#verticalalignment)|Представляет вертикальное выравнивание для рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalAlignment.|
||[verticalOverflow](/javascript/api/excel/excel.textframeupdatedata#verticaloverflow)|Представляет действие вертикального переполнения рамки с текстом. Дополнительные сведения см. в статье Excel.ShapeTextVerticalOverflow.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Возвращает объект TextRange для подстроки в указанном диапазоне.|
||[font](/javascript/api/excel/excel.textrange#font)|Возвращает объект ShapeFont, представляющий атрибуты шрифта для диапазона текста. Только для чтения.|
||[Set (Properties: Excel. TextRange)](/javascript/api/excel/excel.textrange#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Текстранжеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.textrange#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[text](/javascript/api/excel/excel.textrange#text)|Представляет содержимое с обычным текстом в диапазоне текста.|
|[Текстранжедата](/javascript/api/excel/excel.textrangedata)|[font](/javascript/api/excel/excel.textrangedata#font)|Возвращает объект ShapeFont, представляющий атрибуты шрифта для диапазона текста. Только для чтения.|
||[text](/javascript/api/excel/excel.textrangedata#text)|Представляет содержимое с обычным текстом в диапазоне текста.|
|[Текстранжелоадоптионс](/javascript/api/excel/excel.textrangeloadoptions)|[$all](/javascript/api/excel/excel.textrangeloadoptions#$all)||
||[font](/javascript/api/excel/excel.textrangeloadoptions#font)|Возвращает объект ShapeFont, представляющий атрибуты шрифта для диапазона текста.|
||[text](/javascript/api/excel/excel.textrangeloadoptions#text)|Представляет содержимое с обычным текстом в диапазоне текста.|
|[Текстранжеупдатедата](/javascript/api/excel/excel.textrangeupdatedata)|[font](/javascript/api/excel/excel.textrangeupdatedata#font)|Возвращает объект ShapeFont, представляющий атрибуты шрифта для диапазона текста.|
||[text](/javascript/api/excel/excel.textrangeupdatedata#text)|Представляет содержимое с обычным текстом в диапазоне текста.|
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
|[Воркбукаутосавесеттингчанжедевентаргс](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Представляет тип события. Дополнительные сведения см. в статье Excel.EventType.|
|[Воркбукдата](/javascript/api/excel/excel.workbookdata)|[autoSave](/javascript/api/excel/excel.workbookdata#autosave)|Указывает, применяется ли в книге режим автосохранения. Только для чтения.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookdata#calculationengineversion)|Возвращает номер версии модуля вычислений Excel. Только для чтения.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookdata#chartdatapointtrack)|Значение true, если все диаграммы в книге отслеживают точки фактических данных, с которыми они связаны.|
||[isDirty](/javascript/api/excel/excel.workbookdata#isdirty)|Указывает, внесены ли изменения с момента последнего сохранении книги.|
||[previouslySaved](/javascript/api/excel/excel.workbookdata#previouslysaved)|Указывает, сохранялась ли книга ранее (локально или в Интернете). Только для чтения.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookdata#useprecisionasdisplayed)|Значение true, если вычисления в книге выполняются только с той точностью чисел, с которой они отображаются.|
|[Воркбуклоадоптионс](/javascript/api/excel/excel.workbookloadoptions)|[autoSave](/javascript/api/excel/excel.workbookloadoptions#autosave)|Указывает, применяется ли в книге режим автосохранения. Только для чтения.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookloadoptions#calculationengineversion)|Возвращает номер версии модуля вычислений Excel. Только для чтения.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookloadoptions#chartdatapointtrack)|Значение true, если все диаграммы в книге отслеживают точки фактических данных, с которыми они связаны.|
||[isDirty](/javascript/api/excel/excel.workbookloadoptions#isdirty)|Указывает, внесены ли изменения с момента последнего сохранении книги.|
||[previouslySaved](/javascript/api/excel/excel.workbookloadoptions#previouslysaved)|Указывает, сохранялась ли книга ранее (локально или в Интернете). Только для чтения.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookloadoptions#useprecisionasdisplayed)|Значение true, если вычисления в книге выполняются только с той точностью чисел, с которой они отображаются.|
|[Воркбукупдатедата](/javascript/api/excel/excel.workbookupdatedata)|[chartDataPointTrack](/javascript/api/excel/excel.workbookupdatedata#chartdatapointtrack)|Значение true, если все диаграммы в книге отслеживают точки фактических данных, с которыми они связаны.|
||[isDirty](/javascript/api/excel/excel.workbookupdatedata#isdirty)|Указывает, внесены ли изменения с момента последнего сохранении книги.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookupdatedata#useprecisionasdisplayed)|Значение true, если вычисления в книге выполняются только с той точностью чисел, с которой они отображаются.|
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
|[Воркшитколлектионлоадоптионс](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetcollectionloadoptions#autofilter)|Для каждого элемента в коллекции: представляет объект автофильтра на листе.|
||[enableCalculation](/javascript/api/excel/excel.worksheetcollectionloadoptions#enablecalculation)|Для каждого элемента в коллекции: Получает или задает свойство Енаблекалкулатион рабочего листа.|
||[pageLayout](/javascript/api/excel/excel.worksheetcollectionloadoptions#pagelayout)|Для каждого элемента в коллекции: получает объект PageLayout рабочего листа.|
|[Воркшитдата](/javascript/api/excel/excel.worksheetdata)|[autoFilter](/javascript/api/excel/excel.worksheetdata#autofilter)|Представляет объект AutoFilter листа. Только для чтения.|
||[enableCalculation](/javascript/api/excel/excel.worksheetdata#enablecalculation)|Получает или задает свойство enableCalculation для листа.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheetdata#horizontalpagebreaks)|Получает коллекцию горизонтальных разрывов страницы для листа. Эта коллекция содержит только добавленные вручную разрывы страниц.|
||[pageLayout](/javascript/api/excel/excel.worksheetdata#pagelayout)|Получает объект PageLayout листа.|
||[shapes](/javascript/api/excel/excel.worksheetdata#shapes)|Возвращает коллекцию всех объектов Shape на листе. Только для чтения.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheetdata#verticalpagebreaks)|Получает коллекцию вертикальных разрывов страницы для листа. Эта коллекция содержит только добавленные вручную разрывы страниц.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область конкретного листа. Может возвращать пустой объект.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[Воркшитлоадоптионс](/javascript/api/excel/excel.worksheetloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetloadoptions#autofilter)|Представляет объект AutoFilter листа.|
||[enableCalculation](/javascript/api/excel/excel.worksheetloadoptions#enablecalculation)|Получает или задает свойство enableCalculation для листа.|
||[pageLayout](/javascript/api/excel/excel.worksheetloadoptions#pagelayout)|Получает объект PageLayout листа.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Указывает, должно ли совпадение быть полным или частичным. Полное совпадение соответствует всему содержимому ячейки. Значение по умолчанию: false (частичное).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Указывает, учитывается ли регистр при сопоставлении. Значение по умолчанию: false (без учета регистра).|
|[Воркшитупдатедата](/javascript/api/excel/excel.worksheetupdatedata)|[enableCalculation](/javascript/api/excel/excel.worksheetupdatedata#enablecalculation)|Получает или задает свойство enableCalculation для листа.|
||[pageLayout](/javascript/api/excel/excel.worksheetupdatedata#pagelayout)|Получает объект PageLayout листа.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)

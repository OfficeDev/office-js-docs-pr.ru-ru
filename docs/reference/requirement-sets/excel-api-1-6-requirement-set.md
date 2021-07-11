---
title: Excel Набор API JavaScript 1.6
description: Сведения о наборе требований ExcelApi 1.6.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: bc2eb8f182a329808a46f172868b818027f5e367
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350108"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Новые возможности API JavaScript для Excel 1.6

## <a name="conditional-formatting"></a>Условное форматирование

Вводится условное форматирование диапазона. Позволяет использовать следующие типы условного форматирования.

- Цветовая шкала
- Гистограмма
- Набор значков
- Настраиваемый

Дополнительно:

- Возврат диапазона, к которому применено условное форматирование.
- Удаление условного форматирования.
- Обеспечивает приоритет и `stopifTrue` возможности.
- Получение полной коллекции условного форматирования для определенного диапазона.
- Полное удаление условного форматирование в указанном диапазоне.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.6. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.6 или ранее, см. в Excel API в наборе требований [1.6](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Приостанавливает вычисление до вызова следующего "context.sync()".|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Указывает объект Правило в этом условном формате.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Критерии цветовой шкалы.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Если значение true, то цветовая шкала будет иметь три точки (минимальная, средней точки, максимум), в противном случае она будет иметь два (минимум, максимум).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Оператор условного формата значения ячейки.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Условие цветовой шкалы "максимальная точка".|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Условие цветовой шкалы "средняя точка", если используется трехцветная цветовая шкала.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Условие цветовой шкалы "минимальная точка".|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Представление цветового кода HTML цвета (например, #FF0000 представляет красный цвет).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Число, формула или значение NULL (если указан тип LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|На чем должна основываться условная формула критерия.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|HTML-код, представляющий цвет линии границы в формате #RRGGBB (например, "FFA500") или в виде ключевого слова (например, "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|HTML-цветовой код, представляющий цвет заполнения, формы #RRGGBB (например, "FFA500") или как названный HTML-цвет (например, "оранжевый").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Указывает, имеет ли отрицательный DataBar тот же цвет границы, что и положительный DataBar.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Указывает, имеет ли отрицательный DataBar тот же цвет заполнения, что и положительный DataBar.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|HTML-код, представляющий цвет линии границы в формате #RRGGBB (например, "FFA500") или в виде ключевого слова (например, "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|HTML-цветовой код, представляющий цвет заполнения, формы #RRGGBB (например, "FFA500") или как названный HTML-цвет (например, "оранжевый").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Указывает, имеет ли градиент DataBar.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Формула, с помощью которой при необходимости оценивается правило гистограммы.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Тип правила для панели данных.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Удаляет это условное форматирование.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Возврат диапазона, к которому применено условное форматирование.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Возвращает диапазон, к который применяется кондитональный формат, или объект null, если условный формат применяется к нескольким диапазонам.|
||[приоритет](/javascript/api/excel/excel.conditionalformat#priority)|Приоритет (или индекс) в условном наборе форматов, в который в настоящее время существует этот условный формат.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Возвращает свойства условного формата значения ячейки, если текущий условный формат является типом CellValue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Возвращает свойства условного формата значения ячейки, если текущий условный формат является типом CellValue.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Возвращает свойства условного формата ColorScale, если текущий условный формат — это тип ColorScale.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Возвращает свойства условного формата ColorScale, если текущий условный формат — это тип ColorScale.|
||[настраиваемый](/javascript/api/excel/excel.conditionalformat#custom)|Возвращает настраиваемые свойства условного формата, если текущий условный формат является пользовательским типом.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Возвращает настраиваемые свойства условного формата, если текущий условный формат является пользовательским типом.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Возвращает свойства панели данных, если текущий условный формат является панели данных.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Возвращает свойства панели данных, если текущий условный формат является панели данных.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Возвращает свойства условного формата IconSet, если текущий условный формат — это тип IconSet.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Возвращает свойства условного формата IconSet, если текущий условный формат — это тип IconSet.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|Приоритет условного форматирования в пределах текущего класса ConditionalFormatCollection.|
||[предустановка](/javascript/api/excel/excel.conditionalformat#preset)|Возвращает условный формат предварительных критериев.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Возвращает условный формат предварительных критериев.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Возвращает определенные свойства условного формата текста, если текущий условный формат — это текстовый тип.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Возвращает определенные свойства условного формата текста, если текущий условный формат — это текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Возвращает свойства условного формата Top/Bottom, если текущий условный формат — тип TopBottom.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Возвращает свойства условного формата Top/Bottom, если текущий условный формат — тип TopBottom.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Тип условного формата.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Добавляет новый условный формат в коллекцию с первого и верхнего приоритета.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Полное удаление условного форматирование в указанном диапазоне.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Возвращает количество условных форматов в книге.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Возвращает условное форматирование для указанного идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Возвращает условное форматирование по индексу.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Формула, с помощью которой при необходимости оценивается правило условного форматирования на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Формула, с помощью которой при необходимости оценивается правило условного форматирования в формате R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Специальный значок для текущего условия, если он отличается от набора значков по умолчанию, в противном случае возвращается значение NULL.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Число или формула в зависимости от типа.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan или GreaterThanOrEqual для каждого типа правила для условного формата Icon.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|На чем должна основываться условная формула значка.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[критерий](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Критерий условного формата.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML-код, представляющий цвет линии границы в формате #RRGGBB (например, "FFA500") или в виде ключевого слова (например, "orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Постоянное значение, указывающее определенную сторону границы.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Одна из констант стиля линии, определяющая стиль линии границы.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Возвращает объект границы по его индексу.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Получает нижнюю границу.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Количество объектов границы в коллекции.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Получает левую границу.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Получает правую границу.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Получает верхнюю границу.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Удаляет заливку.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML-цветовой код, представляющий цвет заполнения, формы #RRGGBB (например, "FFA500") или как имя HTML-цвета (например, "оранжевый").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Указывает, является ли шрифт смелым.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Удаляет форматирование шрифтов.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Представление цветового кода HTML текстового цвета (например, #FF0000 представляет красный цвет).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Указывает, является ли шрифт italic.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Указывает состояние забастовки шрифта.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Тип подчеркнутого, примененного к шрифту.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Представляет Excel формата номеров для данного диапазона.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Коллекция пограничных объектов, применимых к общему диапазону условного формата.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Возвращает объект заполнения, определенный в общем диапазоне условного формата.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Возвращает объект шрифта, определенный в общем диапазоне условного формата.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Оператор текстового условного формата.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Текстовое значение условного форматирования.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|От 1 до 1000 для числовых рейтингов или от 1 до 100 для процентных рейтингов.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Значения формата на основе верхнего или нижнего ранга.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.customconditionalformat#rule)|Указывает объект Правило в этом условном формате.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|HTML-цветовой код, представляющий цвет линии Axis, формы #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Представление того, как ось определяется для Excel панели данных.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Указывает, в каком направлении должна основываться графика панели данных.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Представление всех значений слева от оси в панели Excel данных.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Представление всех значений справа от оси в панели Excel данных.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Значение true скрывает значения ячеек, где применяется гистограмма.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Массив критериев и iconSets для правил и потенциальных пользовательских значков для условных значков.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Если верно, отменяет заказы значка для IconSet.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Значение true скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Если установлено, отображается параметр IconSet для условного формата.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Правило условного форматирования.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Вычисляет диапазон ячеек на листе.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Коллекция условныхформатов, пересекает диапазон.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.textconditionalformat#rule)|Правило условного форматирования.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Критерии условного формата Top/Bottom.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Вычисляет все ячейки на листе.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

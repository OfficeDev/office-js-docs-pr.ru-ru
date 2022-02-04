---
title: Excel API JavaScript установлено 1.6
description: Сведения о наборе требований ExcelApi 1.6.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
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

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.6. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.6 или ранее, см. Excel API в наборе требований [1.6 или ранее](/javascript/api/excel?view=excel-js-1.6&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendapicalculationuntilnextsync-member(1))|Приостанавливать вычисление, пока не будет `context.sync()` вызван следующий.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#excel-excel-cellvalueconditionalformat-format-member)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.cellvalueconditionalformat#excel-excel-cellvalueconditionalformat-rule-member)|Указывает объект правила в этом условном формате.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#excel-excel-colorscaleconditionalformat-criteria-member)|Критерии цветовой шкалы.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#excel-excel-colorscaleconditionalformat-threecolorscale-member)|Если `true`цветовая шкала будет иметь три точки (минимальная, средней точки, максимум), в противном случае она будет иметь два (минимум, максимум).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-formula1-member)|Формула, если требуется, для оценки правила условного формата.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-formula2-member)|Формула, если требуется, для оценки правила условного формата.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-operator-member)|Оператор условного формата значения ячейки.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-maximum-member)|Максимальная точка критерия цветовой шкалы.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-midpoint-member)|Середина критерия цветовой шкалы, если цветовая шкала — это трехцветная шкала.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-minimum-member)|Минимальная точка критерия цветовой шкалы.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-color-member)|Представление цветового кода HTML цвета (например, #FF0000 представляет красный цвет).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-formula-member)|Число, формула или `null` (если `type` есть `lowestValue`).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-type-member)|На чем должна основываться условная формула критерия.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-bordercolor-member)|ЦВЕТОВой код HTML, представляющий цвет пограничной строки, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-fillcolor-member)|ЦВЕТОВой код HTML, представляющий цвет заполнения, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-matchpositivebordercolor-member)|Указывает, имеет ли отрицательная планка данных тот же цвет границы, что и положительная планка данных.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-matchpositivefillcolor-member)|Указывает, имеет ли отрицательная планка данных тот же цвет заполнения, что и положительный.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-bordercolor-member)|ЦВЕТОВой код HTML, представляющий цвет пограничной строки, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-fillcolor-member)|ЦВЕТОВой код HTML, представляющий цвет заполнения, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-gradientfill-member)|Указывает, есть ли в панели данных градиент.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#excel-excel-conditionaldatabarrule-formula-member)|Формула, если требуется, для оценки правила панели данных.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#excel-excel-conditionaldatabarrule-type-member)|Тип правила для панели данных.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[cellValue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalue-member)|Возвращает свойства условного формата значения ячейки, если текущий условный формат является типом `CellValue` .|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalueornullobject-member)|Возвращает свойства условного формата значения ячейки, если текущий условный формат является типом `CellValue` .|
||[colorScale](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscale-member)|Возвращает свойства условного формата цветовой шкалы, если текущий условный формат является типом `ColorScale` .|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscaleornullobject-member)|Возвращает свойства условного формата цветовой шкалы, если текущий условный формат является типом `ColorScale` .|
||[настраиваемый](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-custom-member)|Возвращает настраиваемые свойства условного формата, если текущий условный формат является пользовательским типом.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-customornullobject-member)|Возвращает настраиваемые свойства условного формата, если текущий условный формат является пользовательским типом.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databar-member)|Возвращает свойства панели данных, если текущий условный формат является панели данных.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databarornullobject-member)|Возвращает свойства панели данных, если текущий условный формат является панели данных.|
||[delete()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-delete-member(1))|Удаляет это условное форматирование.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrange-member(1))|Возврат диапазона, к которому применено условное форматирование.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrangeornullobject-member(1))|Возвращает диапазон, к которому применяется кондитональный формат.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconset-member)|Возвращает свойства условного формата набора значков, если текущий условный формат является типом `IconSet` .|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconsetornullobject-member)|Возвращает свойства условного формата набора значков, если текущий условный формат является типом `IconSet` .|
||[id](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-id-member)|Приоритет условного формата в текущем `ConditionalFormatCollection`.|
||[предустановка](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-preset-member)|Возвращает условный формат предварительных критериев.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-presetornullobject-member)|Возвращает условный формат предварительных критериев.|
||[приоритет](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-priority-member)|Приоритет (или индекс) в условном наборе форматов, в который в настоящее время существует этот условный формат.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-stopiftrue-member)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparison-member)|Возвращает определенные свойства условного формата текста, если текущий условный формат — это текстовый тип.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparisonornullobject-member)|Возвращает определенные свойства условного формата текста, если текущий условный формат — это текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottom-member)|Возвращает свойства верхнего и нижнего условного формата, если текущий условный формат является типом `TopBottom` .|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottomornullobject-member)|Возвращает свойства верхнего и нижнего условного формата, если текущий условный формат является типом `TopBottom` .|
||[type](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-type-member)|Тип условного формата.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-add-member(1))|Добавляет новый условный формат в коллекцию с первого и верхнего приоритета.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-clearall-member(1))|Полное удаление условного форматирование в указанном диапазоне.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getcount-member(1))|Возвращает количество условных форматов в книге.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitem-member(1))|Возвращает условное форматирование для указанного идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemat-member(1))|Возвращает условное форматирование по индексу.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formula-member)|Формула, если требуется, для оценки правила условного формата.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formulalocal-member)|Формула, если требуется, для оценки правила условного формата на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formular1c1-member)|Формула, если требуется, для оценки правила условного формата в нотации в стиле R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-customicon-member)|Пользовательский значок для текущего критерия, если он отличается от набора значков по умолчанию, будет `null` возвращен.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-formula-member)|Число или формула в зависимости от типа.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-operator-member)|`greaterThan` или `greaterThanOrEqual` для каждого из типов правил для условного формата значка.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-type-member)|На чем должна основываться условная формула значка.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[критерий](/javascript/api/excel/excel.conditionalpresetcriteriarule#excel-excel-conditionalpresetcriteriarule-criterion-member)|Критерий условного формата.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-color-member)|ЦВЕТОВой код HTML, представляющий цвет пограничной строки, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-sideindex-member)|Постоянное значение, указывающее определенную сторону границы.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-style-member)|Одна из констант стиля линии, определяющая стиль линии границы.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-bottom-member)|Получает нижнюю границу.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-count-member)|Количество объектов границы в коллекции.|
||[getItem(index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-getitem-member(1))|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-getitemat-member(1))|Возвращает объект границы по его индексу.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-left-member)|Получает левую границу.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-right-member)|Получает правую границу.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-top-member)|Получает верхнюю границу.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#excel-excel-conditionalrangefill-clear-member(1))|Удаляет заливку.|
||[color](/javascript/api/excel/excel.conditionalrangefill#excel-excel-conditionalrangefill-color-member)|ЦВЕТОВой код HTML, представляющий цвет заполнения, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-bold-member)|Указывает, является ли шрифт смелым.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-clear-member(1))|Удаляет форматирование шрифтов.|
||[color](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-color-member)|Представление цветового кода HTML текстового цвета (например, #FF0000 представляет красный цвет).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-italic-member)|Указывает, является ли шрифт italic.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-strikethrough-member)|Указывает состояние забастовки шрифта.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-underline-member)|Тип подчеркнутого, примененного к шрифту.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[borders](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-borders-member)|Коллекция пограничных объектов, применимых к общему диапазону условного формата.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-fill-member)|Возвращает объект заполнения, определенный в общем диапазоне условного формата.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-font-member)|Возвращает объект шрифта, определенный в общем диапазоне условного формата.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-numberformat-member)|Представляет Excel формата номеров для данного диапазона.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#excel-excel-conditionaltextcomparisonrule-operator-member)|Оператор текстового условного формата.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#excel-excel-conditionaltextcomparisonrule-text-member)|Текстовое значение условного формата.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#excel-excel-conditionaltopbottomrule-rank-member)|От 1 до 1000 для числовых рейтингов или от 1 до 100 для процентных рейтингов.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#excel-excel-conditionaltopbottomrule-type-member)|Значения формата на основе верхнего или нижнего ранга.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#excel-excel-customconditionalformat-format-member)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.customconditionalformat#excel-excel-customconditionalformat-rule-member)|Указывает объект в `Rule` этом условном формате.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-axiscolor-member)|ЦВЕТОВой код HTML, представляющий цвет линии Axis, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-axisformat-member)|Представление того, как ось определяется для Excel панели данных.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-bardirection-member)|Указывает, в каком направлении должна основываться графика панели данных.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-lowerboundrule-member)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-negativeformat-member)|Представление всех значений слева от оси в панели Excel данных.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-positiveformat-member)|Представление всех значений справа от оси в панели Excel данных.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-showdatabaronly-member)|Если `true`, скрывает значения из ячеек, где применяется планка данных.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-upperboundrule-member)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-criteria-member)|Набор критериев и наборов значков для правил и потенциальных пользовательских значков для условных значков.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-reverseiconorder-member)|Если `true`, отменит заказы значка для набора значков.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-showicononly-member)|Если `true`, скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-style-member)|Если установлено, отображается параметр набора значков для условного формата.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#excel-excel-presetcriteriaconditionalformat-format-member)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.presetcriteriaconditionalformat#excel-excel-presetcriteriaconditionalformat-rule-member)|Правило условного форматирования.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#excel-excel-range-calculate-member(1))|Вычисляет диапазон ячеек на листе.|
||[conditionalFormats](/javascript/api/excel/excel.range#excel-excel-range-conditionalformats-member)|Эта коллекция `ConditionalFormats` пересекает диапазон.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#excel-excel-textconditionalformat-format-member)|Возвращает объект формата, инкапсулируя шрифт условного формата, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.textconditionalformat#excel-excel-textconditionalformat-rule-member)|Правило условного форматирования.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#excel-excel-topbottomconditionalformat-format-member)|Возвращает объект формата, инкапсулируя шрифт условного формата, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.topbottomconditionalformat#excel-excel-topbottomconditionalformat-rule-member)|Критерии условного формата верхнего и нижнего.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-calculate-member(1))|Вычисляет все ячейки на листе.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)

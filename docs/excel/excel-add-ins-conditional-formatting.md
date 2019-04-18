---
title: Применение условного форматирования к диапазонам с помощью API JavaScript для Excel
description: ''
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 7c6b5b5433e2dc59259eb937ef553ff265443f75
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914384"
---
# <a name="apply-conditional-formatting-to-excel-ranges"></a><span data-ttu-id="fdd08-102">Применение условного форматирования к диапазонам Excel</span><span class="sxs-lookup"><span data-stu-id="fdd08-102">Apply conditional formatting to Excel ranges</span></span>

<span data-ttu-id="fdd08-103">Библиотека JavaScript Excel предоставляет API для применения условного форматирования к диапазонам данных в книгах.</span><span class="sxs-lookup"><span data-stu-id="fdd08-103">The Excel JavaScript Library provides APIs to apply conditional formatting to data ranges in your worksheets.</span></span> <span data-ttu-id="fdd08-104">Эта функция упрощает визуальный анализ больших наборов данных.</span><span class="sxs-lookup"><span data-stu-id="fdd08-104">This functionality makes large sets of data easy to visually parse.</span></span> <span data-ttu-id="fdd08-105">Форматирование также динамически обновляется с учетом изменений в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="fdd08-105">The formatting also dynamically updates based on changes within the range.</span></span> 

> [!NOTE]
> <span data-ttu-id="fdd08-106">В этой статье рассматривается условное форматирование в контексте надстроек JavaScript для Excel. В указанных ниже статьях представлены подробные сведения о всех возможностях условного форматирования в Excel.</span><span class="sxs-lookup"><span data-stu-id="fdd08-106">This article covers conditional formatting in the context of Excel JavaScript add-ins. The following articles provide detailed information about the full conditional formatting capabilities within Excel.</span></span>
> -  [<span data-ttu-id="fdd08-107">Добавление, изменение или удаление условного форматирования</span><span class="sxs-lookup"><span data-stu-id="fdd08-107">Add, change, or clear conditional formats</span></span>](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
> -  [<span data-ttu-id="fdd08-108">Использование формул с условным форматированием</span><span class="sxs-lookup"><span data-stu-id="fdd08-108">Use formulas with conditional formatting</span></span>](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)

## <a name="programmatic-control-of-conditional-formatting"></a><span data-ttu-id="fdd08-109">Программное управление условным форматированием</span><span class="sxs-lookup"><span data-stu-id="fdd08-109">Programmatic control of conditional formatting</span></span>

<span data-ttu-id="fdd08-110">Свойство `Range.conditionalFormats` — это коллекция объектов [ConditionalFormat](/javascript/api/excel/excel.conditionalformat), применяемых к диапазону.</span><span class="sxs-lookup"><span data-stu-id="fdd08-110">The `Range.conditionalFormats` property is a collection of [ConditionalFormat](/javascript/api/excel/excel.conditionalformat) objects that apply to the range.</span></span>  <span data-ttu-id="fdd08-111">Объект `ConditionalFormat` содержит несколько свойств, определяющих применяемый формат на основе [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).</span><span class="sxs-lookup"><span data-stu-id="fdd08-111">The `ConditionalFormat` object contains several properties that define the format to be applied based on the [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).</span></span> 

-   `cellValue`
-   `colorScale`
-   `custom`
-   `dataBar`
-   `iconSet`
-   `preset`
-   `textComparison`
-   `topBottom`

> [!NOTE]
> <span data-ttu-id="fdd08-112">У каждого из этих свойств форматирования есть соответствующий вариант `*OrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="fdd08-112">Each of these formatting properties has a corresponding `*OrNullObject` variant.</span></span> <span data-ttu-id="fdd08-113">Дополнительные сведения об этом шаблоне см. в разделе [Методы \*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="fdd08-113">Learn more about that pattern in the [\*OrNullObject methods](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods) section.</span></span>

<span data-ttu-id="fdd08-114">Для объекта ConditionalFormat можно установить только один тип формата.</span><span class="sxs-lookup"><span data-stu-id="fdd08-114">Only one format type can be set for the ConditionalFormat object.</span></span> <span data-ttu-id="fdd08-115">Это определено свойством `type`, которое является значением перечисления объекта [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).</span><span class="sxs-lookup"><span data-stu-id="fdd08-115">This is determined by the `type` property, which is a [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype) enum value.</span></span> <span data-ttu-id="fdd08-116">Параметр `type` устанавливается при добавлении условного форматирования к диапазону.</span><span class="sxs-lookup"><span data-stu-id="fdd08-116">`type` is set when adding a conditional format to a range.</span></span> 

## <a name="creating-conditional-formatting-rules"></a><span data-ttu-id="fdd08-117">Создание правил условного форматирования</span><span class="sxs-lookup"><span data-stu-id="fdd08-117">Creating conditional formatting rules</span></span>

<span data-ttu-id="fdd08-118">Условное форматирование добавляется к диапазону с помощью `conditionalFormats.add`.</span><span class="sxs-lookup"><span data-stu-id="fdd08-118">Conditional formats are added to a range by using `conditionalFormats.add`.</span></span> <span data-ttu-id="fdd08-119">После добавления можно задать свойства, относящиеся к условному форматированию.</span><span class="sxs-lookup"><span data-stu-id="fdd08-119">Once added, the properties specific to the conditional format can be set.</span></span> <span data-ttu-id="fdd08-120">В примерах ниже показано создание различных типов форматирования.</span><span class="sxs-lookup"><span data-stu-id="fdd08-120">The following examples show the creation of different formatting types.</span></span>

### <a name="cell-valuejavascriptapiexcelexcelcellvalueconditionalformat"></a>[<span data-ttu-id="fdd08-121">Значение ячейки</span><span class="sxs-lookup"><span data-stu-id="fdd08-121">Cell value</span></span>](/javascript/api/excel/excel.cellvalueconditionalformat)

<span data-ttu-id="fdd08-122">При условном форматировании значения ячейки применяется пользовательский формат на основе результатов одной или двух формул в [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule).</span><span class="sxs-lookup"><span data-stu-id="fdd08-122">Cell value conditional formatting applies a user-defined format based on the results of one or two formulas in the [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule).</span></span> <span data-ttu-id="fdd08-123">Свойство `operator` является оператором [ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator), который определяет, как итоговое выражение связано с форматированием.</span><span class="sxs-lookup"><span data-stu-id="fdd08-123">The `operator` property is a [ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator) defining how the resulting expressions relate to the formatting.</span></span>

<span data-ttu-id="fdd08-124">В приведенном ниже примере показано применение красного шрифта ко всем значениям диапазона, которые меньше нуля.</span><span class="sxs-lookup"><span data-stu-id="fdd08-124">The following example shows red font coloring applied to any value in the range less than zero.</span></span>

![Диапазон с отрицательными числами красного цвета.](../images/excel-conditional-format-cell-value.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B21:E23");
const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.cellValue
);

// set the font of negative numbers to red
conditionalFormat.cellValue.format.font.color = "red";
conditionalFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };

await context.sync();
```

### <a name="color-scalejavascriptapiexcelexcelcolorscaleconditionalformat"></a>[<span data-ttu-id="fdd08-126">Цветовая шкала</span><span class="sxs-lookup"><span data-stu-id="fdd08-126">Color scale</span></span>](/javascript/api/excel/excel.colorscaleconditionalformat)

<span data-ttu-id="fdd08-127">При условном форматировании с использованием цветовой шкалы применяется цветовой градиент в диапазоне данных.</span><span class="sxs-lookup"><span data-stu-id="fdd08-127">Color scale conditional formatting applies a color gradient across the data range.</span></span> <span data-ttu-id="fdd08-128">Свойство `criteria` в `ColorScaleConditionalFormat` определяет три точки [ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum` и (при желании) `midpoint`.</span><span class="sxs-lookup"><span data-stu-id="fdd08-128">The `criteria` property on the `ColorScaleConditionalFormat` defines three [ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`, and, optionally, `midpoint`.</span></span> <span data-ttu-id="fdd08-129">У каждой точки условия есть три свойства:</span><span class="sxs-lookup"><span data-stu-id="fdd08-129">Each of the criterion scale points have three properties:</span></span>

-   <span data-ttu-id="fdd08-130">`color` — HTML-код цвета для конечной точки.</span><span class="sxs-lookup"><span data-stu-id="fdd08-130">`color` - The HTML color code for the endpoint.</span></span>
-   <span data-ttu-id="fdd08-131">`formula` — число или формула, представляющая значение конечной точки.</span><span class="sxs-lookup"><span data-stu-id="fdd08-131">`formula` - A number or formula representing the endpoint.</span></span> <span data-ttu-id="fdd08-132">Оно будет равным `null`, если `type` имеет значение `lowestValue` или `highestValue`.</span><span class="sxs-lookup"><span data-stu-id="fdd08-132">This will be `null` if `type` is `lowestValue` or `highestValue`.</span></span>
-   <span data-ttu-id="fdd08-133">`type` — способ оценки формулы.</span><span class="sxs-lookup"><span data-stu-id="fdd08-133">`type` - How the formula should be evaluated.</span></span> <span data-ttu-id="fdd08-134">`highestValue` и `lowestValue` относятся к значениям в форматируемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="fdd08-134">`highestValue` and `lowestValue` refer to values in the range being formatted.</span></span>

<span data-ttu-id="fdd08-135">В приведенном ниже примере показан диапазон, окрашенный с переходом от синего к желтому и красному цвету.</span><span class="sxs-lookup"><span data-stu-id="fdd08-135">The following example shows a range being colored blue to yellow to red.</span></span> <span data-ttu-id="fdd08-136">Обратите внимание, что `minimum` и `maximum` являются минимальным и максимальным значением соответственно, и для них используются формулы `null`.</span><span class="sxs-lookup"><span data-stu-id="fdd08-136">Note that `minimum` and `maximum` are the lowest and highest values respectively and use `null` formulas.</span></span> <span data-ttu-id="fdd08-137">Для значения `midpoint` используется тип `percentage` с формулой `”=50”`, чтобы самая желтая ячейка соответствовала среднему значению.</span><span class="sxs-lookup"><span data-stu-id="fdd08-137">`midpoint` is using the `percentage` type with a formula of `”=50”` so the yellowest cell is the mean value.</span></span>

![Диапазон с низким значением синего цвета, средним значением желтого цвета, высоким значением красного цвета и применением градиента для промежуточных значений.](../images/excel-conditional-format-color-scale.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
      Excel.ConditionalFormatType.colorScale
);

// color the backgrounds of the cells from blue to yellow to red based on value
const criteria = {
      minimum: {
           formula: null,
           type: Excel.ConditionalFormatColorCriterionType.lowestValue,
           color: "blue"
      },
      midpoint: {
           formula: "50",
           type: Excel.ConditionalFormatColorCriterionType.percent,
           color: "yellow"
      },
      maximum: {
           formula: null,
           type: Excel.ConditionalFormatColorCriterionType.highestValue,
           color: "red"
      }
};
conditionalFormat.colorScale.criteria = criteria;

await context.sync();
```

### <a name="customjavascriptapiexcelexcelcustomconditionalformat"></a>[<span data-ttu-id="fdd08-139">Пользовательское</span><span class="sxs-lookup"><span data-stu-id="fdd08-139">Custom</span></span>](/javascript/api/excel/excel.customconditionalformat)

<span data-ttu-id="fdd08-140">При пользовательском условном форматировании применяется пользовательский формат к ячейкам на основе формулы произвольной сложности.</span><span class="sxs-lookup"><span data-stu-id="fdd08-140">Custom conditional formatting applies a user-defined format to the cells based on a formula of arbitrary complexity.</span></span> <span data-ttu-id="fdd08-141">Объект [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) позволяет определять формулу в разных нотациях:</span><span class="sxs-lookup"><span data-stu-id="fdd08-141">The [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) object lets you define the formula in different notations:</span></span>

-   <span data-ttu-id="fdd08-142">`formula` — стандартная нотация.</span><span class="sxs-lookup"><span data-stu-id="fdd08-142">`formula` - Standard notation.</span></span>
-   <span data-ttu-id="fdd08-143">`formulaLocal` — локализованная на основе языка пользователя.</span><span class="sxs-lookup"><span data-stu-id="fdd08-143">`formulaLocal` - Localized based on the user’s language.</span></span>
-   <span data-ttu-id="fdd08-144">`formulaR1C1` — нотация R1C1.</span><span class="sxs-lookup"><span data-stu-id="fdd08-144">`formulaR1C1` - R1C1-style notation.</span></span>

<span data-ttu-id="fdd08-145">В приведенном ниже примере зеленым цветом окрашен шрифт ячеек с более высокими значениями, чем в ячейках слева.</span><span class="sxs-lookup"><span data-stu-id="fdd08-145">The following example colors the fonts green of cells with higher values than the cell to their left.</span></span>

![Диапазон с числами, окрашенными в зеленый цвет, если значение в предшествующем столбце этой строки ниже.](../images/excel-conditional-format-custom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.custom
);

// if a cell has a higher value than the one to its left, set that cell’s font to green
conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT("RC[-1]",0),TRUE)';
conditionalFormat.custom.format.font.color = "green";

await context.sync();

```
### <a name="data-barjavascriptapiexcelexceldatabarconditionalformat"></a>[<span data-ttu-id="fdd08-147">Гистограмма</span><span class="sxs-lookup"><span data-stu-id="fdd08-147">Data bar</span></span>](/javascript/api/excel/excel.databarconditionalformat)

<span data-ttu-id="fdd08-148">При условном форматировании с использованием гистограмм они добавляются к ячейкам.</span><span class="sxs-lookup"><span data-stu-id="fdd08-148">Data bar conditional formatting adds data bars to the cells.</span></span> <span data-ttu-id="fdd08-149">По умолчанию минимальное и максимальное значения в диапазоне создают границы и пропорциональные размеры гистограмм.</span><span class="sxs-lookup"><span data-stu-id="fdd08-149">By default, the minimum and maximum values in the Range form the bounds and proportional sizes of the data bars.</span></span> <span data-ttu-id="fdd08-150">У объекта `DataBarConditionalFormat` есть несколько свойств, определяющих внешний вид гистограммы.</span><span class="sxs-lookup"><span data-stu-id="fdd08-150">The `DataBarConditionalFormat` object has several properties to control the bar’s appearance.</span></span> 

<span data-ttu-id="fdd08-151">В приведенном ниже примере используется форматирование с помощью гистограмм с заполнением слева направо.</span><span class="sxs-lookup"><span data-stu-id="fdd08-151">The following example formats the range with data bars filling left-to-right.</span></span>

![Диапазон с гистограммами позади значений в ячейках.](../images/excel-conditional-format-databar.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.dataBar
);

// give left-to-right, default-appearance data bars to all the cells
conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
await context.sync();
```

### <a name="icon-setjavascriptapiexcelexceliconsetconditionalformat"></a>[<span data-ttu-id="fdd08-153">Набор значков</span><span class="sxs-lookup"><span data-stu-id="fdd08-153">Icon set</span></span>](/javascript/api/excel/excel.iconsetconditionalformat)

<span data-ttu-id="fdd08-154">При условном форматировании с набором значков используются [значки](/javascript/api/excel/excel.icon) Excel для выделения ячеек.</span><span class="sxs-lookup"><span data-stu-id="fdd08-154">Icon set conditional formatting uses Excel [Icons](/javascript/api/excel/excel.icon) to highlight cells.</span></span> <span data-ttu-id="fdd08-155">Свойство `criteria` — это массив объекта [ConditionalIconCriterion](/javascript/api/excel/excel.ConditionalIconCriterion), определяющий добавляемый символ и условия для добавления.</span><span class="sxs-lookup"><span data-stu-id="fdd08-155">The `criteria` property is an array of [ConditionalIconCriterion](/javascript/api/excel/excel.ConditionalIconCriterion), which define the symbol to be inserted and the condition under which it is inserted.</span></span> <span data-ttu-id="fdd08-156">Этот массив автоматически заполняется элементами условия со свойствами по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="fdd08-156">This array is automatically prepopulated with criterion elements with default properties.</span></span> <span data-ttu-id="fdd08-157">Отдельные свойства не могут быть перезаписаны.</span><span class="sxs-lookup"><span data-stu-id="fdd08-157">Individual properties cannot be overwritten.</span></span> <span data-ttu-id="fdd08-158">Вместо этого необходимо заменить весь объект условия.</span><span class="sxs-lookup"><span data-stu-id="fdd08-158">Instead, the whole criteria object must be replaced.</span></span> 

<span data-ttu-id="fdd08-159">В приведенном ниже примере показано применение в диапазоне набора из трех значков с треугольниками.</span><span class="sxs-lookup"><span data-stu-id="fdd08-159">The following example shows a three-triangle icon set applied across the range.</span></span>

![Диапазон с зелеными треугольными треугольниками для значений выше 1000, желтые линии для значений между 700 и 1000, а треугольники Красного вниз для уменьшения значений.](../images/excel-conditional-format-iconset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.iconSet
);

const iconSetCF = conditionalFormat.iconSet;
iconSetCF.style = Excel.IconSet.threeTriangles;

/*
   With a "three*” icon set style, such as "threeTriangles", the third
    element in the criteria array (criteria[2]) defines the "top" icon;
    e.g., a green triangle. The second (criteria[1]) defines the "middle"
    icon, The first (criteria[0]) defines the "low" icon, but it can often 
    be left empty as this method does below, because every cell that
   does not match the other two criteria always gets the low icon.
*/
iconSetCF.criteria = [
    {} as any,
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=700"
      },
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=1000"
      }
];

await context.sync();
```

### <a name="preset-criteriajavascriptapiexcelexcelpresetcriteriaconditionalformat"></a>[<span data-ttu-id="fdd08-161">Готовые условия</span><span class="sxs-lookup"><span data-stu-id="fdd08-161">Preset criteria</span></span>](/javascript/api/excel/excel.presetcriteriaconditionalformat)

<span data-ttu-id="fdd08-162">При условном форматировании с готовыми условиями применяется пользовательский формат к диапазону на основе выбранного стандартного правила.</span><span class="sxs-lookup"><span data-stu-id="fdd08-162">Preset conditional formatting applies a user-defined format to the range based on a selected standard rule.</span></span> <span data-ttu-id="fdd08-163">Эти правила определяются с помощью [ConditionalFormatPresetCriterion](/javascript/api/excel/excel.ConditionalFormatPresetCriterion) в [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule).</span><span class="sxs-lookup"><span data-stu-id="fdd08-163">These rules are defined by the [ConditionalFormatPresetCriterion](/javascript/api/excel/excel.ConditionalFormatPresetCriterion) in the [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule).</span></span> 

<span data-ttu-id="fdd08-164">В приведенном ниже примере белым цветом окрашивается шрифт во всех ячейках со значениями, превышающими среднее значение диапазона хотя бы на одно стандартное отклонение.</span><span class="sxs-lookup"><span data-stu-id="fdd08-164">The following example colors the font white wherever a cell’s value is at least one standard deviation above the range’s average.</span></span>

![Диапазон с белым шрифтом в ячейках со значениями, превышающими среднее значение хотя бы на одно стандартное отклонение.](../images/excel-conditional-format-preset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.presetCriteria
);

// color every cell’s font white that is one standard deviation above average relative to the range
conditionalFormat.preset.format.font.color = "white";
conditionalFormat.preset.rule = {
     criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage
};

await context.sync();
```

### <a name="text-comparisonjavascriptapiexcelexceltextconditionalformat"></a>[<span data-ttu-id="fdd08-166">Сравнение текста</span><span class="sxs-lookup"><span data-stu-id="fdd08-166">Text comparison</span></span>](/javascript/api/excel/excel.textconditionalformat)

<span data-ttu-id="fdd08-167">При условном форматировании со сравнением текста используется сравнение строк в качестве условия.</span><span class="sxs-lookup"><span data-stu-id="fdd08-167">Text comparison conditional formatting uses string comparisons as the condition.</span></span> <span data-ttu-id="fdd08-168">Свойство `rule` является объектом [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule), определяющим строку для сравнения с ячейкой и оператор для указания типа сравнения.</span><span class="sxs-lookup"><span data-stu-id="fdd08-168">The `rule` property is a [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule) defining a string to compare with the cell and an operator to specify the type of comparison.</span></span> 

<span data-ttu-id="fdd08-169">В приведенном ниже примере применяется форматирование шрифта красным цветом, если текст ячейки содержит слово Delayed.</span><span class="sxs-lookup"><span data-stu-id="fdd08-169">The following example formats the font color red when a cell’s text contains “Delayed”.</span></span>

![Диапазон с ячейками, содержащими слово Delayed красного цвета.](../images/excel-conditional-format-text.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B16:D18");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.containsText
);

// color the font of every cell containing “Delayed”
conditionalFormat.textComparison.format.font.color = "red";
conditionalFormat.textComparison.rule = {
     operator: Excel.ConditionalTextOperator.contains,
     text: "Delayed"
};

await context.sync();
```

### <a name="topbottomjavascriptapiexcelexceltopbottomconditionalformat"></a>[<span data-ttu-id="fdd08-171">Верхнее или нижнее значение</span><span class="sxs-lookup"><span data-stu-id="fdd08-171">Top/bottom</span></span>](/javascript/api/excel/excel.TopBottomconditionalformat)

<span data-ttu-id="fdd08-172">При условном форматировании верхнего или нижнего значения применяется форматирование к наибольшему или наименьшему значению в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="fdd08-172">Top/bottom conditional formatting applies a format to the highest or lowest values in a range.</span></span> <span data-ttu-id="fdd08-173">Свойство `rule`, являющееся типом [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), указывает основание для условия (максимальное или минимальное значение), а также применение ранжированной или процентной оценки.</span><span class="sxs-lookup"><span data-stu-id="fdd08-173">The `rule` property, which is of type [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), sets whether the condition is based on the highest or lowest, as well as whether the evaluation is ranked or percentage-based.</span></span> 

<span data-ttu-id="fdd08-174">В приведенном ниже примере применяется зеленое выделение к ячейке с максимальным значением в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="fdd08-174">The following example applies a green highlight to the highest value cell in the range.</span></span>


![Диапазон с максимальным числом, выделенным зеленым цветом.](../images/excel-conditional-format-topbottom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B21:E23");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.topBottom
);

// for the highest valued cell in the range, make the background green
conditionalFormat.topBottom.format.fill.color = "green"
conditionalFormat.topBottom.rule = { rank: 1, type: "TopItems"}

await context.sync();
```

## <a name="multiple-formats-and-priority"></a><span data-ttu-id="fdd08-176">Разные форматирования и приоритет</span><span class="sxs-lookup"><span data-stu-id="fdd08-176">Multiple formats and priority</span></span>

<span data-ttu-id="fdd08-177">К диапазону можно применять несколько типов условного форматирования.</span><span class="sxs-lookup"><span data-stu-id="fdd08-177">You can apply multiple conditional formats to a range.</span></span> <span data-ttu-id="fdd08-178">Если форматы содержат конфликтующие элементы, например разный цвет шрифта, только один формат применяет этот конкретный элемент.</span><span class="sxs-lookup"><span data-stu-id="fdd08-178">If the formats have conflicting elements, such as differing font colors, only one format applies that particular element.</span></span> <span data-ttu-id="fdd08-179">Приоритет определяется свойством `ConditionalFormat.priority`.</span><span class="sxs-lookup"><span data-stu-id="fdd08-179">Precedence is defined by the `ConditionalFormat.priority` property.</span></span> <span data-ttu-id="fdd08-180">Приоритет — это число (равное индексу в `ConditionalFormatCollection`), которое можно установить при создании формата.</span><span class="sxs-lookup"><span data-stu-id="fdd08-180">Priority is a number (equal to the index in the `ConditionalFormatCollection`) and can be set when creating the format.</span></span> <span data-ttu-id="fdd08-181">Чем ниже значение `priority`, тем выше приоритет формата.</span><span class="sxs-lookup"><span data-stu-id="fdd08-181">The lowerer the `priority` value, the higher the priority of the format is.</span></span>

<span data-ttu-id="fdd08-182">В приведенном ниже примере показан выбор цвета шрифта при конфликте между двумя форматами.</span><span class="sxs-lookup"><span data-stu-id="fdd08-182">The following example shows a conflicting font color choice between the two formats.</span></span> <span data-ttu-id="fdd08-183">Для отрицательных чисел применяется полужирный шрифт, но НЕ красный, так как приоритет получает формат, устанавливающий для них синий цвет шрифта.</span><span class="sxs-lookup"><span data-stu-id="fdd08-183">Negative numbers will get a bold font, but NOT a red font, because priority goes to the format that gives them a blue font.</span></span>

![Диапазон с использованием для низких чисел полужирного красного шрифта, а для отрицательных — синего шрифта с зеленым фоном.](../images/excel-conditional-format-priority.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const temperatureDataRange = sheet.tables.getItem("TemperatureTable").getDataBodyRange();


// Set low numbers to bold, dark red font and assign priority 1.
const presetFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.presetCriteria);
presetFormat.preset.format.font.color = "red";
presetFormat.preset.format.font.bold = true;
presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };
presetFormat.priority = 1;

// Set negative numbers to blue font with green background and set priority 0.
const cellValueFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.cellValue);
cellValueFormat.cellValue.format.font.color = "blue";
cellValueFormat.cellValue.format.fill.color = "lightgreen";
cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
cellValueFormat.priority = 0;

await context.sync();

```

### <a name="mutually-exclusive-conditional-formats"></a><span data-ttu-id="fdd08-185">Взаимоисключающие условные форматирования</span><span class="sxs-lookup"><span data-stu-id="fdd08-185">Mutually exclusive conditional formats</span></span>

<span data-ttu-id="fdd08-186">Свойство `stopIfTrue` объекта `ConditionalFormat` не позволяет применять к диапазону условное форматирование с более низким приоритетом.</span><span class="sxs-lookup"><span data-stu-id="fdd08-186">The `stopIfTrue` property of `ConditionalFormat` prevents lower priority conditional formats from being applied to the range.</span></span> <span data-ttu-id="fdd08-187">Если при сопоставлении с диапазоном применяется условное форматирование со свойством `stopIfTrue === true`, последующие условные форматирования не применяются, даже если их элементы не вступают в противоречие.</span><span class="sxs-lookup"><span data-stu-id="fdd08-187">When a range matching the conditional format with `stopIfTrue === true` is applied, no subsequent conditional formats are applied, even if their formatting details are not contradictory.</span></span>

<span data-ttu-id="fdd08-188">В приведенном ниже примере показано добавление в диапазон двух условных форматов.</span><span class="sxs-lookup"><span data-stu-id="fdd08-188">The following example shows two conditional formats being added to a range.</span></span> <span data-ttu-id="fdd08-189">Для отрицательных чисел будет использоваться синий шрифт со светло-зеленым фоном, независимо от того, выполняются ли условия другого формата.</span><span class="sxs-lookup"><span data-stu-id="fdd08-189">Negative numbers will have a blue font with a light green background, regardless of whether the other format condition is true.</span></span>

![Диапазон с применением полужирного красного шрифта для низких чисел, если они не являются отрицательными, так как в этом случае для них используется синий шрифт (не полужирный) с зеленым фоном.](../images/excel-conditional-format-stopiftrue.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const temperatureDataRange = sheet.tables.getItem("TemperatureTable").getDataBodyRange();

// Set low numbers to bold, dark red font and assign priority 1.
const presetFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.presetCriteria);
presetFormat.preset.format.font.color = "red";
presetFormat.preset.format.font.bold = true;
presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };
presetFormat.priority = 1;

// Set negative numbers to blue font with green background and 
// set priority 0, but set stopIfTrue to true, so none of the 
// formatting of the conditional format with the higher priority
// value will apply, not even the bolding of the font.
const cellValueFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.cellValue);
cellValueFormat.cellValue.format.font.color = "blue";
cellValueFormat.cellValue.format.fill.color = "lightgreen";
cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
cellValueFormat.priority = 0;
cellValueFormat.stopIfTrue = true;

await context.sync();
```

## <a name="see-also"></a><span data-ttu-id="fdd08-191">См. также</span><span class="sxs-lookup"><span data-stu-id="fdd08-191">See also</span></span>

- [<span data-ttu-id="fdd08-192">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="fdd08-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](/office/dev/add-ins/excel/excel-add-ins-core-concepts)
- [<span data-ttu-id="fdd08-193">Работа с диапазонами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="fdd08-193">Work with ranges using the Excel JavaScript API</span></span>](/office/dev/add-ins/excel/excel-add-ins-ranges)
- [<span data-ttu-id="fdd08-194">Объект ConditionalFormat (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="fdd08-194">ConditionalFormat Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.conditionalformat)
- [<span data-ttu-id="fdd08-195">Добавление, изменение или удаление условного форматирования</span><span class="sxs-lookup"><span data-stu-id="fdd08-195">Add, change, or clear conditional formats</span></span>](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
- [<span data-ttu-id="fdd08-196">Использование формул с условным форматированием</span><span class="sxs-lookup"><span data-stu-id="fdd08-196">Use formulas with conditional formatting</span></span>](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)

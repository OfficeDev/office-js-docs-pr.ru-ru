---
title: Пустые значения и значения NULL в надстройках Excel
description: Узнайте, как работать с пустыми значениями NULL в методах и свойствах объектной модели Excel.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 3f38569f7342bb88c52ce424db426bfa7939be5e
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409416"
---
# <a name="blank-and-null-values-in-excel-add-ins"></a><span data-ttu-id="e8819-103">Пустые значения и значения NULL в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="e8819-103">Blank and null values in Excel add-ins</span></span>

<span data-ttu-id="e8819-104">Значения `null` и пустые строки имеют специальные применения в API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="e8819-104">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="e8819-105">Они используются для представления пустых ячеек, отсутствия форматирования или значений по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e8819-105">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="e8819-106">В этом разделе описано использование значения `null` и пустой строки при получении и настройке свойств.</span><span class="sxs-lookup"><span data-stu-id="e8819-106">This section details the use of `null` and empty string when getting and setting properties.</span></span>

## <a name="null-input-in-2-d-array"></a><span data-ttu-id="e8819-107">Входное значение null в двумерном массиве</span><span class="sxs-lookup"><span data-stu-id="e8819-107">null input in 2-D Array</span></span>

<span data-ttu-id="e8819-p102">В Excel диапазон представлен двумерным массивом, в котором первое измерение — это строки, а второе — столбцы. Чтобы задать значения, формат чисел или формулу только для определенных ячеек в диапазоне, укажите значения, формат чисел или формулу для этих ячеек в двумерном массиве, а для всех остальных ячеек в этом массиве укажите значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e8819-p102">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="e8819-p103">Например, чтобы изменить формат чисел только для одной ячейки в диапазоне и сохранить существующий формат чисел для всех остальных ячеек в диапазоне, укажите новый формат чисел для ячейки, которую необходимо изменить, а для всех остальных ячеек укажите значение `null`. Во фрагменте кода ниже показано, как задать новый формат чисел для четвертой ячейки в диапазоне, при этом формат чисел для первых трех ячеек в диапазоне останется неизменным.</span><span class="sxs-lookup"><span data-stu-id="e8819-p103">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a><span data-ttu-id="e8819-112">Входное значение null для свойства</span><span class="sxs-lookup"><span data-stu-id="e8819-112">null input for a property</span></span>

<span data-ttu-id="e8819-p104">`null` не является допустимым входным значением для одного свойства. Например, указанный ниже фрагмент кода не является допустимым, так как свойство `values` диапазона не должно иметь значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e8819-p104">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null; // This is not a valid snippet. 
```

<span data-ttu-id="e8819-115">Аналогично, указанный ниже фрагмент кода не является допустимым, так как `null` — недопустимое значение для свойства `color`.</span><span class="sxs-lookup"><span data-stu-id="e8819-115">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a><span data-ttu-id="e8819-116">Значения свойств null в ответе</span><span class="sxs-lookup"><span data-stu-id="e8819-116">null property values in the response</span></span>

<span data-ttu-id="e8819-p105">Если в указанном диапазоне имеются другие значения, свойства форматирования, например `size` и `color` будут содержать значения `null` в ответе. Например, если вы получаете диапазон и загружаете его свойство `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="e8819-p105">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="e8819-119">Если у всех ячеек в диапазоне один и тот же цвет шрифта, свойство `range.format.font.color` указывает этот цвет.</span><span class="sxs-lookup"><span data-stu-id="e8819-119">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="e8819-120">Если в диапазоне используется несколько цветов шрифтов, свойство `range.format.font.color` имеет значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e8819-120">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

## <a name="blank-input-for-a-property"></a><span data-ttu-id="e8819-121">Пустое входное значение для свойства</span><span class="sxs-lookup"><span data-stu-id="e8819-121">Blank input for a property</span></span>

<span data-ttu-id="e8819-p106">Когда вы указываете пустое значение для свойства (то есть две кавычки подряд без других знаков между `''`), это будет интерпретировано как инструкция по очистке или сбросу свойства. Например:</span><span class="sxs-lookup"><span data-stu-id="e8819-p106">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="e8819-124">Если вы укажете пустое значение для свойства `values` диапазона, содержимое диапазона будет очищено.</span><span class="sxs-lookup"><span data-stu-id="e8819-124">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="e8819-125">Если вы укажете пустое значение для свойства `numberFormat`, формат чисел будет "сброшен" до формата `General`.</span><span class="sxs-lookup"><span data-stu-id="e8819-125">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="e8819-126">Если вы укажете пустое значение для свойств `formula` и `formulaLocale`, значения формул будут очищены.</span><span class="sxs-lookup"><span data-stu-id="e8819-126">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

## <a name="blank-property-values-in-the-response"></a><span data-ttu-id="e8819-127">Значения пустых свойств в ответе</span><span class="sxs-lookup"><span data-stu-id="e8819-127">Blank property values in the response</span></span>

<span data-ttu-id="e8819-p107">Для операций чтения пустое значение свойства в ответе (то есть две кавычки подряд без других знаков между `''`) указывает, что ячейка не содержит данных или значения. В первом примере ниже первая и последняя ячейки в диапазоне не содержат данных. Во втором примере две первые ячейки в диапазоне не содержат формул.</span><span class="sxs-lookup"><span data-stu-id="e8819-p107">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

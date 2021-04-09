---
title: Установите формат диапазона с помощью API JavaScript Excel
description: Узнайте, как использовать API JavaScript Excel для набора формата диапазона.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdd78ea69fc38cbefb9d240dbc61554891c73c21
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652909"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a><span data-ttu-id="d6fef-103">Настройка формата диапазона с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="d6fef-103">Set range format using the Excel JavaScript API</span></span>

<span data-ttu-id="d6fef-104">В этой статье данная статья содержит примеры кода, которые устанавливают цвет шрифта, цвет и формат номеров для ячеек в диапазоне с API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="d6fef-104">This article provides code samples that set font color, fill color, and number format for cells in a range with the Excel JavaScript API.</span></span> <span data-ttu-id="d6fef-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="d6fef-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a><span data-ttu-id="d6fef-106">Задание цвета шрифта и цвета заливки</span><span class="sxs-lookup"><span data-stu-id="d6fef-106">Set font color and fill color</span></span>

<span data-ttu-id="d6fef-107">В примере ниже показано, как задать цвет шрифта и цвет заливки для ячеек в диапазоне **B2: E2**.</span><span class="sxs-lookup"><span data-stu-id="d6fef-107">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a><span data-ttu-id="d6fef-108">Данные в диапазоне перед заданием цвета шрифта и цвета заливки</span><span class="sxs-lookup"><span data-stu-id="d6fef-108">Data in range before font color and fill color are set</span></span>

![Данные в Excel перед заданием формата](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a><span data-ttu-id="d6fef-110">Данные в диапазоне после задания цвета шрифта и цвета заливки</span><span class="sxs-lookup"><span data-stu-id="d6fef-110">Data in range after font color and fill color are set</span></span>

![Данные в Excel после задания формата](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a><span data-ttu-id="d6fef-112">Задание формата чисел</span><span class="sxs-lookup"><span data-stu-id="d6fef-112">Set number format</span></span>

<span data-ttu-id="d6fef-113">В примере ниже показано, как задать формат чисел для ячеек в диапазоне **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="d6fef-113">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-number-format-is-set"></a><span data-ttu-id="d6fef-114">Данные в диапазоне перед заданием формата чисел</span><span class="sxs-lookup"><span data-stu-id="d6fef-114">Data in range before number format is set</span></span>

![Данные в Excel перед набором формата номеров](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a><span data-ttu-id="d6fef-116">Данные в диапазоне после задания формата чисел</span><span class="sxs-lookup"><span data-stu-id="d6fef-116">Data in range after number format is set</span></span>

![Данные в Excel после набора формата номеров](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a><span data-ttu-id="d6fef-118">См. также</span><span class="sxs-lookup"><span data-stu-id="d6fef-118">See also</span></span>

- [<span data-ttu-id="d6fef-119">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="d6fef-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d6fef-120">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="d6fef-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="d6fef-121">Настройка и получения диапазонов с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="d6fef-121">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="d6fef-122">Установите и получите значения диапазона, текст или формулы с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="d6fef-122">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)

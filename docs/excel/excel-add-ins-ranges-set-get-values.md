---
title: Установите и получите значения диапазона, текст или формулы с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для набора и получения значений диапазона, текста или формул.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5d9d1bf3b248585bf27ac591754cfa4eb4dd0fbc
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075749"
---
# <a name="set-and-get-range-values-text-or-formulas-using-the-excel-javascript-api"></a><span data-ttu-id="ad904-103">Установите и получите значения диапазона, текст или формулы с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="ad904-103">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>

<span data-ttu-id="ad904-104">В этой статье данная статья содержит примеры кода, которые устанавливают и получают значения диапазона, текст или формулы с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ad904-104">This article provides code samples that set and get range values, text, or formulas with the Excel JavaScript API.</span></span> <span data-ttu-id="ad904-105">Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="ad904-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-values-or-formulas"></a><span data-ttu-id="ad904-106">Задание значений или формул</span><span class="sxs-lookup"><span data-stu-id="ad904-106">Set values or formulas</span></span>

<span data-ttu-id="ad904-107">В следующих примерах кода устанавливаются значения и формулы для одной ячейки или целого ряда ячеек.</span><span class="sxs-lookup"><span data-stu-id="ad904-107">The following code samples set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="ad904-108">Задание значения для одной ячейки</span><span class="sxs-lookup"><span data-stu-id="ad904-108">Set value for a single cell</span></span>

<span data-ttu-id="ad904-109">В примере кода ниже показано, как присвоить ячейке **C3** значение 5, а затем настроить ширину столбцов для наилучшего размещения данных.</span><span class="sxs-lookup"><span data-stu-id="ad904-109">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-value-is-updated"></a><span data-ttu-id="ad904-110">Данные перед изменением значения ячейки</span><span class="sxs-lookup"><span data-stu-id="ad904-110">Data before cell value is updated</span></span>

![Данные в Excel перед обновлением значения ячейки.](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a><span data-ttu-id="ad904-112">Данные после изменения значения ячейки</span><span class="sxs-lookup"><span data-stu-id="ad904-112">Data after cell value is updated</span></span>

![Данные в Excel после обновления значения ячейки.](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="ad904-114">Задание значений для диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-114">Set values for a range of cells</span></span>

<span data-ttu-id="ad904-115">В примере кода ниже показано, как присвоить значения ячейкам в диапазоне **B5:D5**, а затем настроить ширину столбцов для наилучшего размещения данных.</span><span class="sxs-lookup"><span data-stu-id="ad904-115">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];

    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-values-are-updated"></a><span data-ttu-id="ad904-116">Данные перед изменением значений ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-116">Data before cell values are updated</span></span>

![Данные в Excel перед обновлением значений ячейки.](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a><span data-ttu-id="ad904-118">Данные после изменения значений ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-118">Data after cell values are updated</span></span>

![Данные в Excel после обновления значений ячейки.](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="ad904-120">Задание формулы для одной ячейки</span><span class="sxs-lookup"><span data-stu-id="ad904-120">Set formula for a single cell</span></span>

<span data-ttu-id="ad904-121">В примере кода ниже показано, как задать формулу для ячейки **E3**, а затем настроить ширину столбцов для наилучшего размещения данных.</span><span class="sxs-lookup"><span data-stu-id="ad904-121">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formula-is-set"></a><span data-ttu-id="ad904-122">Данные перед заданием формулы для ячейки</span><span class="sxs-lookup"><span data-stu-id="ad904-122">Data before cell formula is set</span></span>

![Данные в Excel перед набором формулы ячейки.](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a><span data-ttu-id="ad904-124">Данные после задания формулы для ячейки</span><span class="sxs-lookup"><span data-stu-id="ad904-124">Data after cell formula is set</span></span>

![Данные в Excel после набора формулы ячейки.](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="ad904-126">Задание формул для диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-126">Set formulas for a range of cells</span></span>

<span data-ttu-id="ad904-127">В примере кода ниже показано, как задать формулы для ячеек в диапазоне **E2:E6**, а затем настроить ширину столбцов для наилучшего размещения данных.</span><span class="sxs-lookup"><span data-stu-id="ad904-127">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];

    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formulas-are-set"></a><span data-ttu-id="ad904-128">Данные перед заданием формул для ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-128">Data before cell formulas are set</span></span>

![Данные в Excel перед набором формул ячейки.](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a><span data-ttu-id="ad904-130">Данные после задания формул для ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-130">Data after cell formulas are set</span></span>

![Данные в Excel после набора формул ячейки.](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="ad904-132">Получение значений, текста или формул</span><span class="sxs-lookup"><span data-stu-id="ad904-132">Get values, text, or formulas</span></span>

<span data-ttu-id="ad904-133">Эти примеры кода получают значения, текст и формулы из ряда ячеек.</span><span class="sxs-lookup"><span data-stu-id="ad904-133">These code samples get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="ad904-134">Получение значений из диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-134">Get values from a range of cells</span></span>

<span data-ttu-id="ad904-135">В следующем примере кода получает диапазон **B2:E6,** загружается его свойство и записывает значения `values` на консоль.</span><span class="sxs-lookup"><span data-stu-id="ad904-135">The following code sample gets the range **B2:E6**, loads its `values` property, and writes the values to the console.</span></span> <span data-ttu-id="ad904-136">Свойство `values` диапазона указывает необработанные значения, содержащиеся в ячейках.</span><span class="sxs-lookup"><span data-stu-id="ad904-136">The `values` property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="ad904-137">Даже если некоторые ячейки в диапазоне содержат формулы, свойство диапазона указывает необработанные значения для этих ячеек, а не какие-либо `values` формулы.</span><span class="sxs-lookup"><span data-stu-id="ad904-137">Even if some cells in a range contain formulas, the `values` property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="ad904-138">Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)</span><span class="sxs-lookup"><span data-stu-id="ad904-138">Data in range (values in column E are a result of formulas)</span></span>

![Данные в Excel после набора формул ячейки.](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="ad904-140">range.values (как записано в консоль в примере кода выше)</span><span class="sxs-lookup"><span data-stu-id="ad904-140">range.values (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="ad904-141">Получение текста из диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-141">Get text from a range of cells</span></span>

<span data-ttu-id="ad904-142">Следующий пример кода получает диапазон **B2:E6,** загружает его `text` свойство и записывает его на консоль.</span><span class="sxs-lookup"><span data-stu-id="ad904-142">The following code sample gets the range **B2:E6**, loads its `text` property, and writes it to the console.</span></span> <span data-ttu-id="ad904-143">Свойство диапазона указывает значения отображения для `text` ячеек в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="ad904-143">The `text` property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="ad904-144">Даже если некоторые ячейки в диапазоне содержат формулы, свойство диапазона указывает значения отображения для этих ячеек, а не любые `text` формулы.</span><span class="sxs-lookup"><span data-stu-id="ad904-144">Even if some cells in a range contain formulas, the `text` property of the range specifies the display values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="ad904-145">Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)</span><span class="sxs-lookup"><span data-stu-id="ad904-145">Data in range (values in column E are a result of formulas)</span></span>

![Данные в Excel после набора формул ячейки.](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="ad904-147">range.text (как записано в консоль в примере кода выше)</span><span class="sxs-lookup"><span data-stu-id="ad904-147">range.text (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="ad904-148">Получение формул из диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="ad904-148">Get formulas from a range of cells</span></span>

<span data-ttu-id="ad904-149">Следующий пример кода получает диапазон **B2:E6,** загружает его `formulas` свойство и записывает его на консоль.</span><span class="sxs-lookup"><span data-stu-id="ad904-149">The following code sample gets the range **B2:E6**, loads its `formulas` property, and writes it to the console.</span></span> <span data-ttu-id="ad904-150">Свойство диапазона указывает формулы для ячеек в диапазоне, содержащих формулы, и необработанные значения для ячеек в диапазоне, которые не `formulas` содержат формул.</span><span class="sxs-lookup"><span data-stu-id="ad904-150">The `formulas` property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="ad904-151">Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)</span><span class="sxs-lookup"><span data-stu-id="ad904-151">Data in range (values in column E are a result of formulas)</span></span>

![Данные в Excel после набора формул ячейки.](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="ad904-153">range.formulas (как записано в консоль в примере кода выше)</span><span class="sxs-lookup"><span data-stu-id="ad904-153">range.formulas (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="see-also"></a><span data-ttu-id="ad904-154">См. также</span><span class="sxs-lookup"><span data-stu-id="ad904-154">See also</span></span>

- [<span data-ttu-id="ad904-155">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="ad904-155">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ad904-156">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="ad904-156">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="ad904-157">Настройка и получения диапазонов с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="ad904-157">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="ad904-158">Настройка формата диапазона с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="ad904-158">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)

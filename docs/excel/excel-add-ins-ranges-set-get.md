---
title: Установите и получите выбранный диапазон с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для набора и получения выбранного диапазона с Excel API JavaScript.
ms.date: 06/22/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9e4c31f165b39d45fac342cb85577ef737105472
ms.sourcegitcommit: ebb4a22a0bdeb5623c72b9494ebbce3909d0c90c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2021
ms.locfileid: "53126738"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a><span data-ttu-id="7748b-103">Установите и получите выбранный диапазон с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="7748b-103">Set and get the selected range using the Excel JavaScript API</span></span>

<span data-ttu-id="7748b-104">В этой статье данная статья содержит примеры кода, которые устанавливают и получают выбранный диапазон с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7748b-104">This article provides code samples that set and get the selected range with the Excel JavaScript API.</span></span> <span data-ttu-id="7748b-105">Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="7748b-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="7748b-106">Задание выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="7748b-106">Set the selected range</span></span>

<span data-ttu-id="7748b-107">В примере кода ниже показано, как выделить диапазон **B2:E6** на активном листе.</span><span class="sxs-lookup"><span data-stu-id="7748b-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="7748b-108">Выделенный диапазон B2:E6</span><span class="sxs-lookup"><span data-stu-id="7748b-108">Selected range B2:E6</span></span>

![Выбранный диапазон в Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="7748b-110">Получение выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="7748b-110">Get the selected range</span></span>

<span data-ttu-id="7748b-111">Следующий пример кода получает выбранный диапазон, загружает его `address` свойство и пишет сообщение на консоль.</span><span class="sxs-lookup"><span data-stu-id="7748b-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="select-the-edge-of-a-used-range-online-only"></a><span data-ttu-id="7748b-112">Выберите край используемого диапазона (только для сети)</span><span class="sxs-lookup"><span data-stu-id="7748b-112">Select the edge of a used range (online-only)</span></span>

> [!NOTE]
> <span data-ttu-id="7748b-113">В настоящее время эти методы `Range.getRangeEdge` доступны только в `Range.getExtendedRange` ExcelApiOnline 1.1.</span><span class="sxs-lookup"><span data-stu-id="7748b-113">The `Range.getRangeEdge` and `Range.getExtendedRange` methods are currently only available in ExcelApiOnline 1.1.</span></span> <span data-ttu-id="7748b-114">Дополнительные дополнительные [Excel API JavaScript в интернете.](../reference/requirement-sets/excel-api-online-requirement-set.md)</span><span class="sxs-lookup"><span data-stu-id="7748b-114">To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).</span></span>

<span data-ttu-id="7748b-115">Методы [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) и [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) позволяют надстройке реплицировать поведение ярлыков выбора клавиатуры, выбрав край используемого диапазона на основе выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="7748b-115">The [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) and [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) methods let your add-in replicate the behavior of the keyboard selection shortcuts, selecting the edge of the used range based on the currently selected range.</span></span> <span data-ttu-id="7748b-116">Дополнительные дополнительные новости об используемых диапазонах см. в [руб. Get used range.](excel-add-ins-ranges-get.md#get-used-range)</span><span class="sxs-lookup"><span data-stu-id="7748b-116">To learn more about used ranges, see [Get used range](excel-add-ins-ranges-get.md#get-used-range).</span></span>

<span data-ttu-id="7748b-117">На следующем скриншоте используется диапазон таблицы со значениями в каждой ячейке **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="7748b-117">In the following screenshot, the used range is the table with values in each cell, **C5:F12**.</span></span> <span data-ttu-id="7748b-118">Пустые ячейки за пределами этой таблицы находятся за пределами используемого диапазона.</span><span class="sxs-lookup"><span data-stu-id="7748b-118">The empty cells outside this table are outside the used range.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a><span data-ttu-id="7748b-120">Выберите ячейку на краю текущего используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="7748b-120">Select the cell at the edge of the current used range</span></span>

<span data-ttu-id="7748b-121">В следующем примере кода показано, как использовать метод для выбора ячейки на самом дальнем краю используемого диапазона тока `Range.getRangeEdge` в направлении вверх.</span><span class="sxs-lookup"><span data-stu-id="7748b-121">The following code sample shows how use the `Range.getRangeEdge` method to select the cell at the furthest edge of the current used range, in the direction up.</span></span> <span data-ttu-id="7748b-122">Это действие соответствует результату использования клавиши Ctrl+Up при выборе диапазона.</span><span class="sxs-lookup"><span data-stu-id="7748b-122">This action matches the result of using the Ctrl+Up arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    var rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="7748b-123">Перед выбором ячейки на краю используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="7748b-123">Before selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="7748b-124">На следующем скриншоте показан используемый диапазон и выбранный диапазон в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="7748b-124">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="7748b-125">Используемый диапазон — это таблица с данными **на C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="7748b-125">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="7748b-126">В этой таблице выбирается **диапазон D8:E9.**</span><span class="sxs-lookup"><span data-stu-id="7748b-126">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="7748b-127">Этот выбор является *состоянием до* запуска `Range.getRangeEdge` метода.</span><span class="sxs-lookup"><span data-stu-id="7748b-127">This selection is the *before* state, prior to running the `Range.getRangeEdge` method.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="7748b-130">После выбора ячейки на краю используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="7748b-130">After selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="7748b-131">На следующем скриншоте показана та же таблица, что и на предыдущем скриншоте, с данными в диапазоне **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="7748b-131">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="7748b-132">В этой таблице выбирается **диапазон D5.**</span><span class="sxs-lookup"><span data-stu-id="7748b-132">Inside this table, the range **D5** is selected.</span></span> <span data-ttu-id="7748b-133">Этот выбор после *состояния,* после запуска метода, чтобы выбрать ячейку на краю используемого диапазона `Range.getRangeEdge` в направлении вверх.</span><span class="sxs-lookup"><span data-stu-id="7748b-133">This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a><span data-ttu-id="7748b-136">Выберите все ячейки от текущего диапазона до дальнего края используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="7748b-136">Select all cells from current range to furthest edge of used range</span></span>

<span data-ttu-id="7748b-137">В следующем примере кода показано, как использовать метод для выбора всех ячеек из выбранного диапазона до самого дальнего края используемого диапазона в направлении `Range.getExtendedRange` вниз.</span><span class="sxs-lookup"><span data-stu-id="7748b-137">The following code sample shows how use the `Range.getExtendedRange` method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down.</span></span> <span data-ttu-id="7748b-138">Это действие соответствует результату использования клавиши Ctrl+Shift+Down при выборе диапазона.</span><span class="sxs-lookup"><span data-stu-id="7748b-138">This action matches the result of using the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    var extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="7748b-139">Перед выбором всех ячеек от текущего диапазона до края используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="7748b-139">Before selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="7748b-140">На следующем скриншоте показан используемый диапазон и выбранный диапазон в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="7748b-140">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="7748b-141">Используемый диапазон — это таблица с данными **на C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="7748b-141">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="7748b-142">В этой таблице выбирается **диапазон D8:E9.**</span><span class="sxs-lookup"><span data-stu-id="7748b-142">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="7748b-143">Этот выбор является *состоянием до* запуска `Range.getExtendedRange` метода.</span><span class="sxs-lookup"><span data-stu-id="7748b-143">This selection is the *before* state, prior to running the `Range.getExtendedRange` method.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="7748b-146">После выбора всех ячеек от текущего диапазона до края используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="7748b-146">After selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="7748b-147">На следующем скриншоте показана та же таблица, что и на предыдущем скриншоте, с данными в диапазоне **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="7748b-147">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="7748b-148">В этой таблице выбирается **диапазон D8:E12.**</span><span class="sxs-lookup"><span data-stu-id="7748b-148">Inside this table, the range **D8:E12** is selected.</span></span> <span data-ttu-id="7748b-149">Этот выбор *после* состояния после запуска метода для выбора всех ячеек от текущего диапазона до края используемого диапазона `Range.getExtendedRange` в направлении вниз.</span><span class="sxs-lookup"><span data-stu-id="7748b-149">This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the down direction.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a><span data-ttu-id="7748b-152">См. также</span><span class="sxs-lookup"><span data-stu-id="7748b-152">See also</span></span>

- [<span data-ttu-id="7748b-153">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="7748b-153">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="7748b-154">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="7748b-154">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="7748b-155">Установите и получите значения диапазона, текст или формулы с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="7748b-155">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="7748b-156">Настройка формата диапазона с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="7748b-156">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)

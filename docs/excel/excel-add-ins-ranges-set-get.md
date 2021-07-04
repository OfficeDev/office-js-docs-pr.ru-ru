---
title: Установите и получите выбранный диапазон с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для набора и получения выбранного диапазона с Excel API JavaScript.
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 623ba5c1b9e76151d4a2c4b169e655236b37e8c8
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290784"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a><span data-ttu-id="9deb7-103">Установите и получите выбранный диапазон с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="9deb7-103">Set and get the selected range using the Excel JavaScript API</span></span>

<span data-ttu-id="9deb7-104">В этой статье данная статья содержит примеры кода, которые устанавливают и получают выбранный диапазон с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9deb7-104">This article provides code samples that set and get the selected range with the Excel JavaScript API.</span></span> <span data-ttu-id="9deb7-105">Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="9deb7-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="9deb7-106">Задание выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-106">Set the selected range</span></span>

<span data-ttu-id="9deb7-107">В примере кода ниже показано, как выделить диапазон **B2:E6** на активном листе.</span><span class="sxs-lookup"><span data-stu-id="9deb7-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="9deb7-108">Выделенный диапазон B2:E6</span><span class="sxs-lookup"><span data-stu-id="9deb7-108">Selected range B2:E6</span></span>

![Выбранный диапазон в Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="9deb7-110">Получение выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-110">Get the selected range</span></span>

<span data-ttu-id="9deb7-111">Следующий пример кода получает выбранный диапазон, загружает его `address` свойство и пишет сообщение на консоль.</span><span class="sxs-lookup"><span data-stu-id="9deb7-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="select-the-edge-of-a-used-range"></a><span data-ttu-id="9deb7-112">Выберите край используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-112">Select the edge of a used range</span></span>

<span data-ttu-id="9deb7-113">Методы [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) и [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) позволяют надстройке реплицировать поведение ярлыков выбора клавиатуры, выбрав край используемого диапазона на основе выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="9deb7-113">The [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) and [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) methods let your add-in replicate the behavior of the keyboard selection shortcuts, selecting the edge of the used range based on the currently selected range.</span></span> <span data-ttu-id="9deb7-114">Дополнительные дополнительные новости об используемых диапазонах см. в [руб. Get used range.](excel-add-ins-ranges-get.md#get-used-range)</span><span class="sxs-lookup"><span data-stu-id="9deb7-114">To learn more about used ranges, see [Get used range](excel-add-ins-ranges-get.md#get-used-range).</span></span>

<span data-ttu-id="9deb7-115">На следующем скриншоте используется диапазон таблицы со значениями в каждой ячейке **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="9deb7-115">In the following screenshot, the used range is the table with values in each cell, **C5:F12**.</span></span> <span data-ttu-id="9deb7-116">Пустые ячейки за пределами этой таблицы находятся за пределами используемого диапазона.</span><span class="sxs-lookup"><span data-stu-id="9deb7-116">The empty cells outside this table are outside the used range.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a><span data-ttu-id="9deb7-118">Выберите ячейку на краю текущего используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-118">Select the cell at the edge of the current used range</span></span>

<span data-ttu-id="9deb7-119">В следующем примере кода показано, как использовать метод для выбора ячейки на самом дальнем краю используемого диапазона тока `Range.getRangeEdge` в направлении вверх.</span><span class="sxs-lookup"><span data-stu-id="9deb7-119">The following code sample shows how use the `Range.getRangeEdge` method to select the cell at the furthest edge of the current used range, in the direction up.</span></span> <span data-ttu-id="9deb7-120">Это действие соответствует результату использования клавиши Ctrl+Up при выборе диапазона.</span><span class="sxs-lookup"><span data-stu-id="9deb7-120">This action matches the result of using the Ctrl+Up arrow key keyboard shortcut while a range is selected.</span></span>

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

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="9deb7-121">Перед выбором ячейки на краю используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-121">Before selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="9deb7-122">На следующем скриншоте показан используемый диапазон и выбранный диапазон в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="9deb7-122">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="9deb7-123">Используемый диапазон — это таблица с данными **на C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="9deb7-123">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="9deb7-124">В этой таблице выбирается **диапазон D8:E9.**</span><span class="sxs-lookup"><span data-stu-id="9deb7-124">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="9deb7-125">Этот выбор является *состоянием до* запуска `Range.getRangeEdge` метода.</span><span class="sxs-lookup"><span data-stu-id="9deb7-125">This selection is the *before* state, prior to running the `Range.getRangeEdge` method.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="9deb7-128">После выбора ячейки на краю используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-128">After selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="9deb7-129">На следующем скриншоте показана та же таблица, что и на предыдущем скриншоте, с данными в диапазоне **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="9deb7-129">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="9deb7-130">В этой таблице выбирается **диапазон D5.**</span><span class="sxs-lookup"><span data-stu-id="9deb7-130">Inside this table, the range **D5** is selected.</span></span> <span data-ttu-id="9deb7-131">Этот выбор после *состояния,* после запуска метода, чтобы выбрать ячейку на краю используемого диапазона `Range.getRangeEdge` в направлении вверх.</span><span class="sxs-lookup"><span data-stu-id="9deb7-131">This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a><span data-ttu-id="9deb7-134">Выберите все ячейки от текущего диапазона до дальнего края используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-134">Select all cells from current range to furthest edge of used range</span></span>

<span data-ttu-id="9deb7-135">В следующем примере кода показано, как использовать метод для выбора всех ячеек из выбранного диапазона до самого дальнего края используемого диапазона в направлении `Range.getExtendedRange` вниз.</span><span class="sxs-lookup"><span data-stu-id="9deb7-135">The following code sample shows how use the `Range.getExtendedRange` method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down.</span></span> <span data-ttu-id="9deb7-136">Это действие соответствует результату использования клавиши Ctrl+Shift+Down при выборе диапазона.</span><span class="sxs-lookup"><span data-stu-id="9deb7-136">This action matches the result of using the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.</span></span>

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

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="9deb7-137">Перед выбором всех ячеек от текущего диапазона до края используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-137">Before selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="9deb7-138">На следующем скриншоте показан используемый диапазон и выбранный диапазон в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="9deb7-138">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="9deb7-139">Используемый диапазон — это таблица с данными **на C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="9deb7-139">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="9deb7-140">В этой таблице выбирается **диапазон D8:E9.**</span><span class="sxs-lookup"><span data-stu-id="9deb7-140">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="9deb7-141">Этот выбор является *состоянием до* запуска `Range.getExtendedRange` метода.</span><span class="sxs-lookup"><span data-stu-id="9deb7-141">This selection is the *before* state, prior to running the `Range.getExtendedRange` method.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="9deb7-144">После выбора всех ячеек от текущего диапазона до края используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="9deb7-144">After selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="9deb7-145">На следующем скриншоте показана та же таблица, что и на предыдущем скриншоте, с данными в диапазоне **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="9deb7-145">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="9deb7-146">В этой таблице выбирается **диапазон D8:E12.**</span><span class="sxs-lookup"><span data-stu-id="9deb7-146">Inside this table, the range **D8:E12** is selected.</span></span> <span data-ttu-id="9deb7-147">Этот выбор *после* состояния после запуска метода для выбора всех ячеек от текущего диапазона до края используемого диапазона `Range.getExtendedRange` в направлении вниз.</span><span class="sxs-lookup"><span data-stu-id="9deb7-147">This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the down direction.</span></span>

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a><span data-ttu-id="9deb7-150">См. также</span><span class="sxs-lookup"><span data-stu-id="9deb7-150">See also</span></span>

- [<span data-ttu-id="9deb7-151">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="9deb7-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9deb7-152">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="9deb7-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="9deb7-153">Установите и получите значения диапазона, текст или формулы с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="9deb7-153">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="9deb7-154">Настройка формата диапазона с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="9deb7-154">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)

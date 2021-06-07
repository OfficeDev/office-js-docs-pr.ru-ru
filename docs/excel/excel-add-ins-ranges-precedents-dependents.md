---
title: Работа с прецедентами формул и зависимыми с помощью Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для получения прецедентов формул и зависимых.
ms.date: 06/03/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 78fa4fb070ede85d139425a9d59ba1224785a605
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783531"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a><span data-ttu-id="7fa72-103">Получите прецеденты формул и иждивенцев с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="7fa72-103">Get formula precedents and dependents using the Excel JavaScript API</span></span>

<span data-ttu-id="7fa72-104">Excel часто ссылаются на другие ячейки.</span><span class="sxs-lookup"><span data-stu-id="7fa72-104">Excel formulas often refer to other cells.</span></span> <span data-ttu-id="7fa72-105">Эти межклеточные ссылки называются "прецедентами" и "зависимыми".</span><span class="sxs-lookup"><span data-stu-id="7fa72-105">These cross-cell references are known as "precedents" and "dependents".</span></span> <span data-ttu-id="7fa72-106">Прецедент — это ячейка, которая предоставляет данные формуле.</span><span class="sxs-lookup"><span data-stu-id="7fa72-106">A precedent is a cell that provides data to a formula.</span></span> <span data-ttu-id="7fa72-107">Зависимая ячейка содержит формулу, которая ссылается на другие ячейки.</span><span class="sxs-lookup"><span data-stu-id="7fa72-107">A dependent is a cell that contains a formula that refers to other cells.</span></span> <span data-ttu-id="7fa72-108">Дополнительные дополнительные Excel, связанные с отношениями между ячейками, см. в руб. Отображение взаимосвязей между [формулами и ячейками.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)</span><span class="sxs-lookup"><span data-stu-id="7fa72-108">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span>

<span data-ttu-id="7fa72-109">Ячейка может иметь ячейку прецедента, и эта ячейка прецедента может иметь свои собственные ячейки прецедента.</span><span class="sxs-lookup"><span data-stu-id="7fa72-109">A cell may have a precedent cell, and that precedent cell may have its own precedent cells.</span></span> <span data-ttu-id="7fa72-110">"Прямой прецедент" является первой предыдущей группой ячеек в этой последовательности, аналогичной концепции родителей в родительских отношениях с ребенком.</span><span class="sxs-lookup"><span data-stu-id="7fa72-110">A "direct precedent" is the first preceding group of cells in this sequence, similar to the concept of parents in a parent-child relationship.</span></span> <span data-ttu-id="7fa72-111">"Прямая зависимость" — это первая зависимая группа ячеек в последовательности, похожая на детей в отношениях между родителем и ребенком.</span><span class="sxs-lookup"><span data-stu-id="7fa72-111">A "direct dependent" is the first dependent group of cells in a sequence, similar to children in a parent-child relationship.</span></span> <span data-ttu-id="7fa72-112">Ячейки, которые относятся к другим ячейкам в книге, но отношения которых не являются отношениями между родителями и детьми, не являются прямыми иждивенцами или прямыми прецедентами.</span><span class="sxs-lookup"><span data-stu-id="7fa72-112">Cells that refer to other cells in a workbook, but whose relationship is not a parent-child relationship, are not direct dependents or direct precedents.</span></span>

<span data-ttu-id="7fa72-113">В этой статье приводится пример кода, который извлекает прямые прецеденты и напрямую зависит от формул с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7fa72-113">This article provides code samples that retrieve direct precedents and direct dependents of formulas using the Excel JavaScript API.</span></span> <span data-ttu-id="7fa72-114">Полный список свойств и методов, поддерживаемых объектом, см. в руб. `Range` [Range Object (API JavaScript для Excel).](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="7fa72-114">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="get-the-direct-precedents-of-a-formula"></a><span data-ttu-id="7fa72-115">Получите прямые прецеденты формулы</span><span class="sxs-lookup"><span data-stu-id="7fa72-115">Get the direct precedents of a formula</span></span>

<span data-ttu-id="7fa72-116">Найдите прямые ячейки прецедента формулы [с помощью Range.getDirectPrecedents.](/javascript/api/excel/excel.range#getdirectprecedents--)</span><span class="sxs-lookup"><span data-stu-id="7fa72-116">Locate a formula's direct precedent cells with [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span></span> <span data-ttu-id="7fa72-117">`Range.getDirectPrecedents` возвращает `WorkbookRangeAreas` объект.</span><span class="sxs-lookup"><span data-stu-id="7fa72-117">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="7fa72-118">Этот объект содержит адреса всех прямых прецедентов в книге.</span><span class="sxs-lookup"><span data-stu-id="7fa72-118">This object contains the addresses of all the direct precedents in the workbook.</span></span> <span data-ttu-id="7fa72-119">Для каждого таблицы имеется отдельный объект, содержащий по `RangeAreas` крайней мере один прецедент формулы.</span><span class="sxs-lookup"><span data-stu-id="7fa72-119">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="7fa72-120">Дополнительные сведения о работе с объектом см. в совместной работе с несколькими диапазонами `RangeAreas` [Excel надстройки.](excel-add-ins-multiple-ranges.md)</span><span class="sxs-lookup"><span data-stu-id="7fa72-120">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="7fa72-121">На следующем скриншоте показан результат выбора кнопки **Trace Precedents** в пользовательском Excel интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="7fa72-121">The following screenshot shows the result of selecting the **Trace Precedents** button in the Excel UI.</span></span> <span data-ttu-id="7fa72-122">Эта кнопка рисует стрелку из ячеек-прецедентов в выбранную ячейку.</span><span class="sxs-lookup"><span data-stu-id="7fa72-122">This button draws an arrow from precedent cells to the selected cell.</span></span> <span data-ttu-id="7fa72-123">Выбранная ячейка **E3** содержит формулу "=C3 \* D3", поэтому **C3** и **D3** являются прецедентными ячейками.</span><span class="sxs-lookup"><span data-stu-id="7fa72-123">The selected cell, **E3**, contains the formula "=C3 \* D3", so both **C3** and **D3** are precedent cells.</span></span> <span data-ttu-id="7fa72-124">В отличие Excel пользовательского интерфейса, `getDirectPrecedents` метод не рисует стрелки.</span><span class="sxs-lookup"><span data-stu-id="7fa72-124">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span>

![Отслеживание прецедентных ячеек стрелки Excel пользовательского интерфейса](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> <span data-ttu-id="7fa72-126">Метод `getDirectPrecedents` не может получить ячейки прецедента в книгах.</span><span class="sxs-lookup"><span data-stu-id="7fa72-126">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span>

<span data-ttu-id="7fa72-127">В следующем примере кода получаются прямые прецеденты для активного диапазона, а затем изменяется фоновый цвет этих ячеек-прецедентов на желтый.</span><span class="sxs-lookup"><span data-stu-id="7fa72-127">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span>

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula-preview"></a><span data-ttu-id="7fa72-128">Получить прямые иждивенцы формулы (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="7fa72-128">Get the direct dependents of a formula (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="7fa72-129">В `Range.getDirectDependents` настоящее время метод доступен только в общедоступных предварительных версиях.</span><span class="sxs-lookup"><span data-stu-id="7fa72-129">The `Range.getDirectDependents` method is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="7fa72-130">Найдите прямые зависимые ячейки формулы [с помощью Range.getDirectDependents.](/javascript/api/excel/excel.range#getDirectDependents__)</span><span class="sxs-lookup"><span data-stu-id="7fa72-130">Locate a formula's direct dependent cells with [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span></span> <span data-ttu-id="7fa72-131">Как `Range.getDirectPrecedents` , также возвращает `Range.getDirectDependents` `WorkbookRangeAreas` объект.</span><span class="sxs-lookup"><span data-stu-id="7fa72-131">Like `Range.getDirectPrecedents`, `Range.getDirectDependents` also returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="7fa72-132">Этот объект содержит адреса всех прямых иждивенцев в книге.</span><span class="sxs-lookup"><span data-stu-id="7fa72-132">This object contains the addresses of all the direct dependents in the workbook.</span></span> <span data-ttu-id="7fa72-133">Он имеет отдельный `RangeAreas` объект для каждого таблицы, содержащего по крайней мере одну зависимую формулу.</span><span class="sxs-lookup"><span data-stu-id="7fa72-133">It has a separate `RangeAreas` object for each worksheet containing at least one formula dependent.</span></span> <span data-ttu-id="7fa72-134">Дополнительные сведения о работе с объектом см. в совместной работе с несколькими диапазонами `RangeAreas` [Excel надстройки.](excel-add-ins-multiple-ranges.md)</span><span class="sxs-lookup"><span data-stu-id="7fa72-134">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="7fa72-135">На следующем скриншоте показан результат выбора кнопки **Trace Dependents** в пользовательском Excel интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="7fa72-135">The following screenshot shows the result of selecting the **Trace Dependents** button in the Excel UI.</span></span> <span data-ttu-id="7fa72-136">Эта кнопка рисует стрелку из зависимых ячеек в выбранную ячейку.</span><span class="sxs-lookup"><span data-stu-id="7fa72-136">This button draws an arrow from dependent cells to the selected cell.</span></span> <span data-ttu-id="7fa72-137">Выбранная ячейка **D3** имеет ячейку **E3** в качестве зависимой.</span><span class="sxs-lookup"><span data-stu-id="7fa72-137">The selected cell, **D3**, has cell **E3** as a dependent.</span></span> <span data-ttu-id="7fa72-138">**E3** содержит формулу "=C3 \* D3".</span><span class="sxs-lookup"><span data-stu-id="7fa72-138">**E3** contains the formula "=C3 \* D3".</span></span> <span data-ttu-id="7fa72-139">В отличие Excel пользовательского интерфейса, `getDirectDependents` метод не рисует стрелки.</span><span class="sxs-lookup"><span data-stu-id="7fa72-139">Unlike the Excel UI button, the `getDirectDependents` method does not draw arrows.</span></span>

![Отслеживание зависимых ячеек стрелки Excel пользовательского интерфейса](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> <span data-ttu-id="7fa72-141">Метод `getDirectDependents` не может получить зависимые ячейки в книгах.</span><span class="sxs-lookup"><span data-stu-id="7fa72-141">The `getDirectDependents` method can't retrieve dependent cells across workbooks.</span></span>

<span data-ttu-id="7fa72-142">В следующем примере кода получаются прямые иждивенцы для активного диапазона, а затем изменяется фоновый цвет этих зависимых ячеек на желтый.</span><span class="sxs-lookup"><span data-stu-id="7fa72-142">The following code sample gets the direct dependents for the active range and then changes the background color of those dependent cells to yellow.</span></span>

```js
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="7fa72-143">См. также</span><span class="sxs-lookup"><span data-stu-id="7fa72-143">See also</span></span>

- [<span data-ttu-id="7fa72-144">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="7fa72-144">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="7fa72-145">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="7fa72-145">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="7fa72-146">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="7fa72-146">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
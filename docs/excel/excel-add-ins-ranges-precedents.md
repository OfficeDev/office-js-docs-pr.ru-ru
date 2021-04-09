---
title: Работа с прецедентами формул с помощью API JavaScript Excel
description: Узнайте, как использовать API JavaScript Excel для получения прецедентов формул.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652915"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a><span data-ttu-id="fc1fa-103">Получите прецеденты формул с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="fc1fa-103">Get formula precedents using the Excel JavaScript API</span></span>

<span data-ttu-id="fc1fa-104">В этой статье приводится пример кода, который извлекает прецеденты формул с помощью API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-104">This article provides a code sample that retrieves formula precedents using the Excel JavaScript API.</span></span> <span data-ttu-id="fc1fa-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="fc1fa-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="get-formula-precedents"></a><span data-ttu-id="fc1fa-106">Получить прецеденты формул</span><span class="sxs-lookup"><span data-stu-id="fc1fa-106">Get formula precedents</span></span>

<span data-ttu-id="fc1fa-107">Формула Excel часто ссылается на другие ячейки.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-107">An Excel formula often refers to other cells.</span></span> <span data-ttu-id="fc1fa-108">Когда ячейка предоставляет данные формуле, она называется формулой "прецедент".</span><span class="sxs-lookup"><span data-stu-id="fc1fa-108">When a cell provides data to a formula, it is known as a formula "precedent".</span></span> <span data-ttu-id="fc1fa-109">Дополнительные новости о свойствах Excel, связанных с отношениями между ячейками, см. в дополнительных подробностях отображения взаимосвязей между [формулами и ячейками.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)</span><span class="sxs-lookup"><span data-stu-id="fc1fa-109">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span> 

<span data-ttu-id="fc1fa-110">С [помощью Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--)надстройка может найти прямые ячейки прецедента формулы.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-110">With [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), your add-in can locate a formula's direct precedent cells.</span></span> <span data-ttu-id="fc1fa-111">`Range.getDirectPrecedents` возвращает `WorkbookRangeAreas` объект.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-111">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="fc1fa-112">Этот объект содержит адреса всех прецедентов в книге.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-112">This object contains the addresses of all the precedents in the workbook.</span></span> <span data-ttu-id="fc1fa-113">Для каждого таблицы имеется отдельный объект, содержащий по `RangeAreas` крайней мере один прецедент формулы.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-113">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="fc1fa-114">Дополнительные сведения о работе с объектом см. в совместной работе с несколькими диапазонами в `RangeAreas` [надстройки Excel.](excel-add-ins-multiple-ranges.md)</span><span class="sxs-lookup"><span data-stu-id="fc1fa-114">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="fc1fa-115">В пользовательском интерфейсе Excel кнопка **Trace Precedents** рисует стрелку из ячеек-прецедентов в выбранную формулу.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-115">In the Excel UI, the **Trace Precedents** button draws an arrow from precedent cells to the selected formula.</span></span> <span data-ttu-id="fc1fa-116">В отличие от кнопки пользовательского интерфейса Excel, `getDirectPrecedents` метод не рисует стрелки.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-116">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="fc1fa-117">Метод `getDirectPrecedents` не может получить ячейки прецедента в книгах.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-117">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span> 

<span data-ttu-id="fc1fa-118">В следующем примере кода получаются прямые прецеденты для активного диапазона, а затем изменяется фоновый цвет этих ячеек-прецедентов на желтый.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-118">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span> 

> [!NOTE]
> <span data-ttu-id="fc1fa-119">Активный диапазон должен содержать формулу, которая ссылается на другие ячейки в той же книге, чтобы выделение работало правильно.</span><span class="sxs-lookup"><span data-stu-id="fc1fa-119">The active range must contain a formula that references other cells in the same workbook for the highlighting to work properly.</span></span> 

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
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="fc1fa-120">См. также</span><span class="sxs-lookup"><span data-stu-id="fc1fa-120">See also</span></span>

- [<span data-ttu-id="fc1fa-121">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="fc1fa-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="fc1fa-122">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="fc1fa-122">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="fc1fa-123">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="fc1fa-123">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

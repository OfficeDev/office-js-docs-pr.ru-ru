---
title: Диапазоны групп с Excel API JavaScript
description: Узнайте, как сгруппить строки или столбцы диапазона вместе, чтобы создать контур с Excel API JavaScript.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 960a394a1467ec1fe55ff8dbf7b0a3f39fd355a5
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075721"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a><span data-ttu-id="94306-103">Диапазоны групп для контура с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="94306-103">Group ranges for an outline using the Excel JavaScript API</span></span>

<span data-ttu-id="94306-104">В этой статье приводится пример кода, в который показано, как группировать диапазоны для контура с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="94306-104">This article provides a code sample that shows how to group ranges for an outline using the Excel JavaScript API.</span></span> <span data-ttu-id="94306-105">Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="94306-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a><span data-ttu-id="94306-106">Групповые строки или столбцы диапазона для контура</span><span class="sxs-lookup"><span data-stu-id="94306-106">Group rows or columns of a range for an outline</span></span>

<span data-ttu-id="94306-107">Строки или столбцы диапазона можно сгруппить для создания [контура.](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)</span><span class="sxs-lookup"><span data-stu-id="94306-107">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="94306-108">Эти группы можно свернуть и расширить, чтобы скрыть и показать соответствующие ячейки.</span><span class="sxs-lookup"><span data-stu-id="94306-108">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="94306-109">Это упрощает быстрый анализ данных верхнего верхней строки.</span><span class="sxs-lookup"><span data-stu-id="94306-109">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="94306-110">Чтобы сделать эти группы контуров, используйте [Range.group.](/javascript/api/excel/excel.range#group-groupoption-)</span><span class="sxs-lookup"><span data-stu-id="94306-110">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="94306-111">Контур может иметь иерархию, в которой небольшие группы вложены в более крупные группы.</span><span class="sxs-lookup"><span data-stu-id="94306-111">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="94306-112">Это позволяет просматривать контуры на разных уровнях.</span><span class="sxs-lookup"><span data-stu-id="94306-112">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="94306-113">Изменение уровня видимых контуров можно сделать программным путем с помощью метода [Worksheet.showOutlineLevels.](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)</span><span class="sxs-lookup"><span data-stu-id="94306-113">Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method.</span></span> <span data-ttu-id="94306-114">Обратите внимание, Excel поддерживает только восемь уровней групп контуров.</span><span class="sxs-lookup"><span data-stu-id="94306-114">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="94306-115">В следующем примере кода создается контур с двумя уровнями групп для строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="94306-115">The following code sample creates an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="94306-116">На последующем изображении показаны группировки этого контура.</span><span class="sxs-lookup"><span data-stu-id="94306-116">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="94306-117">В примере кода диапазоны, которые группуются, не включают строку или столбец управления контурами (в этом примере "Итоги").</span><span class="sxs-lookup"><span data-stu-id="94306-117">In the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="94306-118">Группа определяет, что будет свернуто, а не строка или столбец с управлением.</span><span class="sxs-lookup"><span data-stu-id="94306-118">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);
```

![Диапазон с двухуровневой двухмерной схемой.](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a><span data-ttu-id="94306-120">Удаление группировки из строк или столбцов диапазона</span><span class="sxs-lookup"><span data-stu-id="94306-120">Remove grouping from rows or columns of a range</span></span>

<span data-ttu-id="94306-121">Чтобы разгруппировать строку или группу столбцов, используйте [метод Range.ungroup.](/javascript/api/excel/excel.range#ungroup-groupoption-)</span><span class="sxs-lookup"><span data-stu-id="94306-121">To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="94306-122">Это удаляет внешний уровень из контура.</span><span class="sxs-lookup"><span data-stu-id="94306-122">This removes the outermost level from the outline.</span></span> <span data-ttu-id="94306-123">Если несколько групп одного и того же типа строки или столбца находятся на одном уровне в указанном диапазоне, все эти группы негруппировываются.</span><span class="sxs-lookup"><span data-stu-id="94306-123">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="see-also"></a><span data-ttu-id="94306-124">См. также</span><span class="sxs-lookup"><span data-stu-id="94306-124">See also</span></span>

- [<span data-ttu-id="94306-125">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="94306-125">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="94306-126">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="94306-126">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="94306-127">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="94306-127">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

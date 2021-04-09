---
title: Вставьте диапазоны с помощью API JavaScript Excel
description: Узнайте, как вставить ряд ячеек с API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 401a08dd10b3775012738ab9c80ec6ab367555ec
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652921"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a><span data-ttu-id="3e0ca-103">Вставьте диапазон ячеек с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="3e0ca-103">Insert a range of cells using the Excel JavaScript API</span></span>

<span data-ttu-id="3e0ca-104">В этой статье содержится пример кода, который вставляет ряд ячеек с API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="3e0ca-104">This article provides a code sample that inserts a range of cells with the Excel JavaScript API.</span></span> <span data-ttu-id="3e0ca-105">Полный список свойств и методов, поддерживаемых объектом, см. `Range` в класс [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="3e0ca-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="3e0ca-106">Вставка диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="3e0ca-106">Insert a range of cells</span></span>

<span data-ttu-id="3e0ca-107">В примере кода ниже показано, как вставить диапазон ячеек в расположение **B4:E4** и сдвинуть другие ячейки вниз, чтобы освободить место для новых ячеек.</span><span class="sxs-lookup"><span data-stu-id="3e0ca-107">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="3e0ca-108">Данные перед вставкой диапазона</span><span class="sxs-lookup"><span data-stu-id="3e0ca-108">Data before range is inserted</span></span>

![Данные в Excel перед вставкой диапазона](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="3e0ca-110">Данные после вставки диапазона</span><span class="sxs-lookup"><span data-stu-id="3e0ca-110">Data after range is inserted</span></span>

![Данные в Excel после вставки диапазона](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a><span data-ttu-id="3e0ca-112">См. также</span><span class="sxs-lookup"><span data-stu-id="3e0ca-112">See also</span></span>

- [<span data-ttu-id="3e0ca-113">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="3e0ca-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="3e0ca-114">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="3e0ca-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="3e0ca-115">Очистить или удалить диапазоны с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="3e0ca-115">Clear or delete a ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-clear-delete.md)

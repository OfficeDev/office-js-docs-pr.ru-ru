---
title: Вставьте диапазоны с Excel API JavaScript
description: Узнайте, как вставить ряд ячеек с Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0571e7d6140f5023008654a1e74d7abf6b3cab0a
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075784"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a><span data-ttu-id="81ade-103">Вставьте диапазон ячеек с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="81ade-103">Insert a range of cells using the Excel JavaScript API</span></span>

<span data-ttu-id="81ade-104">В этой статье содержится пример кода, который вставляет ряд ячеек с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="81ade-104">This article provides a code sample that inserts a range of cells with the Excel JavaScript API.</span></span> <span data-ttu-id="81ade-105">Полный список свойств и методов, поддерживаемых объектом, см. `Range` [в Excel. Класс Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="81ade-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="81ade-106">Вставка диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="81ade-106">Insert a range of cells</span></span>

<span data-ttu-id="81ade-107">В примере кода ниже показано, как вставить диапазон ячеек в расположение **B4:E4** и сдвинуть другие ячейки вниз, чтобы освободить место для новых ячеек.</span><span class="sxs-lookup"><span data-stu-id="81ade-107">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="81ade-108">Данные перед вставкой диапазона</span><span class="sxs-lookup"><span data-stu-id="81ade-108">Data before range is inserted</span></span>

![Данные в Excel перед вставкой диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="81ade-110">Данные после вставки диапазона</span><span class="sxs-lookup"><span data-stu-id="81ade-110">Data after range is inserted</span></span>

![Данные в Excel после вставки диапазона.](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a><span data-ttu-id="81ade-112">См. также</span><span class="sxs-lookup"><span data-stu-id="81ade-112">See also</span></span>

- [<span data-ttu-id="81ade-113">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="81ade-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="81ade-114">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="81ade-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="81ade-115">Очистить или удалить диапазоны с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="81ade-115">Clear or delete a ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-clear-delete.md)

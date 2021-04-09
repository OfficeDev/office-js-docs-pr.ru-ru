---
title: Очистить или удалить диапазоны с помощью API JavaScript Excel
description: Узнайте, как очистить или удалить диапазоны с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7e030c6b5ba7ba6e6c54e9be0524cd93c2516bcb
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652970"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="2711d-103">Очистить или удалить диапазоны с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="2711d-103">Clear or delete ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="2711d-104">В этой статье данная статья содержит примеры кода, которые очищают и удаляют диапазоны с помощью API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="2711d-104">This article provides code samples that clear and delete ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="2711d-105">Полный список свойств и методов, поддерживаемых объектом, см. в класс `Range` [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="2711d-105">For the complete list of properties and methods supported by the `Range` object, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="2711d-106">Очистка диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="2711d-106">Clear a range of cells</span></span>

<span data-ttu-id="2711d-107">В примере кода ниже показано, как удалить все содержимое и форматирование ячеек в диапазоне **E2:E5**.</span><span class="sxs-lookup"><span data-stu-id="2711d-107">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a><span data-ttu-id="2711d-108">Данные перед очисткой диапазона</span><span class="sxs-lookup"><span data-stu-id="2711d-108">Data before range is cleared</span></span>

![Данные в Excel перед очисткой диапазона](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a><span data-ttu-id="2711d-110">Данные после очистки диапазона</span><span class="sxs-lookup"><span data-stu-id="2711d-110">Data after range is cleared</span></span>

![Данные в Excel после очистки диапазона](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="2711d-112">Удаление диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="2711d-112">Delete a range of cells</span></span>

<span data-ttu-id="2711d-113">В следующем примере кода удаляются ячейки в диапазоне **B4:E4** и перемещаются другие ячейки для заполнения пространства, освобождаемого удаленными ячейками.</span><span class="sxs-lookup"><span data-stu-id="2711d-113">The following code sample deletes the cells in the range **B4:E4** and shifts other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a><span data-ttu-id="2711d-114">Данные перед удалением диапазона</span><span class="sxs-lookup"><span data-stu-id="2711d-114">Data before range is deleted</span></span>

![Данные в Excel перед удалением диапазона](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a><span data-ttu-id="2711d-116">Данные после удаления диапазона</span><span class="sxs-lookup"><span data-stu-id="2711d-116">Data after range is deleted</span></span>

![Данные в Excel после удаления диапазона](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a><span data-ttu-id="2711d-118">См. также</span><span class="sxs-lookup"><span data-stu-id="2711d-118">See also</span></span>

- [<span data-ttu-id="2711d-119">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="2711d-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="2711d-120">Настройка и получения диапазонов с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="2711d-120">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="2711d-121">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="2711d-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)

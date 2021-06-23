---
title: Очистить или удалить диапазоны с Excel API JavaScript
description: Узнайте, как очистить или удалить диапазоны с помощью Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a1bd99db3aa9af3903552d9cefc6ec6d21701136
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075833"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="f4c1b-103">Очистить или удалить диапазоны с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="f4c1b-103">Clear or delete ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="f4c1b-104">В этой статье данная статья содержит примеры кода, которые очищают и удаляют диапазоны с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f4c1b-104">This article provides code samples that clear and delete ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="f4c1b-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="f4c1b-105">For the complete list of properties and methods supported by the `Range` object, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="f4c1b-106">Очистка диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="f4c1b-106">Clear a range of cells</span></span>

<span data-ttu-id="f4c1b-107">В примере кода ниже показано, как удалить все содержимое и форматирование ячеек в диапазоне **E2:E5**.</span><span class="sxs-lookup"><span data-stu-id="f4c1b-107">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a><span data-ttu-id="f4c1b-108">Данные перед очисткой диапазона</span><span class="sxs-lookup"><span data-stu-id="f4c1b-108">Data before range is cleared</span></span>

![Данные в Excel перед очисткой диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a><span data-ttu-id="f4c1b-110">Данные после очистки диапазона</span><span class="sxs-lookup"><span data-stu-id="f4c1b-110">Data after range is cleared</span></span>

![Данные в Excel после очистки диапазона.](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="f4c1b-112">Удаление диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="f4c1b-112">Delete a range of cells</span></span>

<span data-ttu-id="f4c1b-113">В следующем примере кода удаляются ячейки в диапазоне **B4:E4** и перемещаются другие ячейки для заполнения пространства, освобождаемого удаленными ячейками.</span><span class="sxs-lookup"><span data-stu-id="f4c1b-113">The following code sample deletes the cells in the range **B4:E4** and shifts other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a><span data-ttu-id="f4c1b-114">Данные перед удалением диапазона</span><span class="sxs-lookup"><span data-stu-id="f4c1b-114">Data before range is deleted</span></span>

![Данные в Excel перед удалением диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a><span data-ttu-id="f4c1b-116">Данные после удаления диапазона</span><span class="sxs-lookup"><span data-stu-id="f4c1b-116">Data after range is deleted</span></span>

![Данные в Excel после удаления диапазона.](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a><span data-ttu-id="f4c1b-118">См. также</span><span class="sxs-lookup"><span data-stu-id="f4c1b-118">See also</span></span>

- [<span data-ttu-id="f4c1b-119">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="f4c1b-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="f4c1b-120">Настройка и получения диапазонов с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="f4c1b-120">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="f4c1b-121">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="f4c1b-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)

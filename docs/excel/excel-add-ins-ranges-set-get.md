---
title: Установите и получите выбранный диапазон с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для набора и получения диапазонов с Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0bd4a4f4bcf40e7899ee429cdc631a43ba176077
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075777"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="1a0e4-103">Настройка и получения диапазонов с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="1a0e4-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="1a0e4-104">В этой статье данная статья содержит примеры кода, которые устанавливают и получают диапазоны с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1a0e4-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="1a0e4-105">Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="1a0e4-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="1a0e4-106">Задание выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="1a0e4-106">Set the selected range</span></span>

<span data-ttu-id="1a0e4-107">В примере кода ниже показано, как выделить диапазон **B2:E6** на активном листе.</span><span class="sxs-lookup"><span data-stu-id="1a0e4-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="1a0e4-108">Выделенный диапазон B2:E6</span><span class="sxs-lookup"><span data-stu-id="1a0e4-108">Selected range B2:E6</span></span>

![Выбранный диапазон в Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="1a0e4-110">Получение выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="1a0e4-110">Get the selected range</span></span>

<span data-ttu-id="1a0e4-111">Следующий пример кода получает выбранный диапазон, загружает его `address` свойство и пишет сообщение на консоль.</span><span class="sxs-lookup"><span data-stu-id="1a0e4-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="1a0e4-112">См. также</span><span class="sxs-lookup"><span data-stu-id="1a0e4-112">See also</span></span>

- [<span data-ttu-id="1a0e4-113">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="1a0e4-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="1a0e4-114">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="1a0e4-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="1a0e4-115">Установите и получите значения диапазона, текст или формулы с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="1a0e4-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="1a0e4-116">Настройка формата диапазона с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="1a0e4-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)

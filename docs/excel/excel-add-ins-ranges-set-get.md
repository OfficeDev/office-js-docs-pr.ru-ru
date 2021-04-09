---
title: Установите и получите выбранный диапазон с помощью API JavaScript Excel
description: Узнайте, как использовать API JavaScript Excel для набора и получения диапазонов с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 06b6219924f0667ecef57d608cb417a76ef8031d
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652882"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="4be26-103">Настройка и получения диапазонов с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="4be26-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="4be26-104">В этой статье данная статья содержит примеры кода, которые устанавливают и получают диапазоны с API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="4be26-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="4be26-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="4be26-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="4be26-106">Задание выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="4be26-106">Set the selected range</span></span>

<span data-ttu-id="4be26-107">В примере кода ниже показано, как выделить диапазон **B2:E6** на активном листе.</span><span class="sxs-lookup"><span data-stu-id="4be26-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="4be26-108">Выделенный диапазон B2:E6</span><span class="sxs-lookup"><span data-stu-id="4be26-108">Selected range B2:E6</span></span>

![Выделенный диапазон в Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="4be26-110">Получение выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="4be26-110">Get the selected range</span></span>

<span data-ttu-id="4be26-111">Следующий пример кода получает выбранный диапазон, загружает его `address` свойство и пишет сообщение на консоль.</span><span class="sxs-lookup"><span data-stu-id="4be26-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4be26-112">См. также</span><span class="sxs-lookup"><span data-stu-id="4be26-112">See also</span></span>

- [<span data-ttu-id="4be26-113">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="4be26-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4be26-114">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="4be26-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="4be26-115">Установите и получите значения диапазона, текст или формулы с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="4be26-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="4be26-116">Настройка формата диапазона с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="4be26-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)

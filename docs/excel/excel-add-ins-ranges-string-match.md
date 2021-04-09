---
title: Поиск строки с помощью API JavaScript Excel
description: Узнайте, как найти строку в диапазоне с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b649bb249cd24d7578bc4f8285e5d0a23d0e4cd
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652888"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="45550-103">Поиск строки в диапазоне с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="45550-103">Find a string within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="45550-104">В этой статье приводится пример кода, который находит строку в диапазоне с помощью API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="45550-104">This article provides a code sample that finds a string within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="45550-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="45550-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a><span data-ttu-id="45550-106">Соответствие строке в диапазоне</span><span class="sxs-lookup"><span data-stu-id="45550-106">Match a string within a range</span></span>

<span data-ttu-id="45550-107">У объекта `Range` есть метод `find` для поиска указанной строки в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="45550-107">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="45550-108">Он возвращает диапазон первой ячейки с текстом, соответствующим критериям.</span><span class="sxs-lookup"><span data-stu-id="45550-108">It returns the range of the first cell with matching text.</span></span>

<span data-ttu-id="45550-109">Приведенный ниже пример кода находит первую ячейку со значением, соответствующим строке **Food** (Еда), и заносит ее адрес в консоль.</span><span class="sxs-lookup"><span data-stu-id="45550-109">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="45550-110">Обратите внимание, что метод `find` выдает ошибку `ItemNotFound`, если указанной строки не существует в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="45550-110">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="45550-111">Если ожидается, что указанная строка может отсутствовать в диапазоне, используйте вместо этого метод [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), чтобы ваш код корректно обработал этот сценарий.</span><span class="sxs-lookup"><span data-stu-id="45550-111">If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="45550-112">Если метод `find` вызывается для диапазона, представляющего одну ячейку, поиск выполняется во всем листе.</span><span class="sxs-lookup"><span data-stu-id="45550-112">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="45550-113">Поиск начинается в этой ячейке и продолжается в направлении, которое определяется параметром `SearchCriteria.searchDirection`, охватывающим концы листа при необходимости.</span><span class="sxs-lookup"><span data-stu-id="45550-113">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="45550-114">См. также</span><span class="sxs-lookup"><span data-stu-id="45550-114">See also</span></span>

- [<span data-ttu-id="45550-115">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="45550-115">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="45550-116">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="45550-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="45550-117">Поиск специальных ячеек в диапазоне с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="45550-117">Find special cells within a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-special-cells.md)

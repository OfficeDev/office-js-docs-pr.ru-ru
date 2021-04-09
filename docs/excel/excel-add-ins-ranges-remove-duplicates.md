---
title: Удаление дубликатов с помощью API JavaScript Excel
description: Узнайте, как использовать API JavaScript Excel для удаления дубликатов.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0a2a076398e15d1b3b9db963a85703782056c91e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652912"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a><span data-ttu-id="b651f-103">Удаление дубликатов с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="b651f-103">Remove duplicates using the Excel JavaScript API</span></span>

<span data-ttu-id="b651f-104">В этой статье содержится пример кода, который удаляет дублирующиеся записи в диапазоне с помощью API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="b651f-104">This article provides a code sample that removes duplicate entries in a range using the Excel JavaScript API.</span></span> <span data-ttu-id="b651f-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="b651f-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="remove-rows-with-duplicate-entries"></a><span data-ttu-id="b651f-106">Удаление строк с дублирующими записями</span><span class="sxs-lookup"><span data-stu-id="b651f-106">Remove rows with duplicate entries</span></span>

<span data-ttu-id="b651f-107">Метод [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) удаляет строки с дублирующимися записями в указанных столбцах.</span><span class="sxs-lookup"><span data-stu-id="b651f-107">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="b651f-108">Метод проходит через каждую строку в диапазоне от самого низкого значения индекса до индекса с самым высоким значением в диапазоне (сверху донизу).</span><span class="sxs-lookup"><span data-stu-id="b651f-108">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="b651f-109">Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="b651f-109">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="b651f-110">Строки в диапазоне под удаленной строкой сдвигаются вверх.</span><span class="sxs-lookup"><span data-stu-id="b651f-110">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="b651f-111">Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.</span><span class="sxs-lookup"><span data-stu-id="b651f-111">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="b651f-112">Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов.</span><span class="sxs-lookup"><span data-stu-id="b651f-112">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="b651f-113">Этот массив отсчитывается от нуля относительно диапазона, а не листа.</span><span class="sxs-lookup"><span data-stu-id="b651f-113">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="b651f-114">Метод также принимает параметр boolean, который указывает, является ли первая строка загонщиком.</span><span class="sxs-lookup"><span data-stu-id="b651f-114">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="b651f-115">При значении **true** верхняя строка игнорируется при поиске дубликатов.</span><span class="sxs-lookup"><span data-stu-id="b651f-115">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="b651f-116">Метод возвращает объект, который указывает количество удаленных строк и количество `removeDuplicates` `RemoveDuplicatesResult` оставшихся уникальных строк.</span><span class="sxs-lookup"><span data-stu-id="b651f-116">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="b651f-117">При использовании метода диапазона имейте в виду `removeDuplicates` следующее:</span><span class="sxs-lookup"><span data-stu-id="b651f-117">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="b651f-118">Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций.</span><span class="sxs-lookup"><span data-stu-id="b651f-118">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="b651f-119">Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.</span><span class="sxs-lookup"><span data-stu-id="b651f-119">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="b651f-120">Пустые ячейки не игнорируются функцией `removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="b651f-120">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="b651f-121">Значение пустой ячейки обрабатывается как любое другое значение.</span><span class="sxs-lookup"><span data-stu-id="b651f-121">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="b651f-122">Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="b651f-122">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="b651f-123">В следующем примере кода показано удаление записей с дублирующими значениями в первом столбце.</span><span class="sxs-lookup"><span data-stu-id="b651f-123">The following code sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

### <a name="data-before-duplicate-entries-are-removed"></a><span data-ttu-id="b651f-124">Данные перед удалением дублирующих записей</span><span class="sxs-lookup"><span data-stu-id="b651f-124">Data before duplicate entries are removed</span></span>

![Данные в Excel перед запуском метода удаления дубликатов диапазона](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a><span data-ttu-id="b651f-126">Данные после удаления дублирующих записей</span><span class="sxs-lookup"><span data-stu-id="b651f-126">Data after duplicate entries are removed</span></span>

![Данные в Excel после запуска метода удаления дубликатов диапазона](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="b651f-128">См. также</span><span class="sxs-lookup"><span data-stu-id="b651f-128">See also</span></span>

- [<span data-ttu-id="b651f-129">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="b651f-129">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b651f-130">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="b651f-130">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="b651f-131">Диапазоны вырезать, скопировать и вклеить с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="b651f-131">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-cut-copy-paste.md)
- [<span data-ttu-id="b651f-132">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="b651f-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

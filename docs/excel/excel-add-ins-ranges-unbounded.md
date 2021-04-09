---
title: Чтение или написание в неограниченый диапазон с помощью API JavaScript Excel
description: Узнайте, как использовать API JavaScript Excel для чтения или записи в неограниченый диапазон.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652885"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a><span data-ttu-id="db4d3-103">Чтение или написание в неограниченый диапазон с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="db4d3-103">Read or write to an unbounded range using the Excel JavaScript API</span></span>

<span data-ttu-id="db4d3-104">В этой статье описывается, как читать и писать в неограниченый диапазон с API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="db4d3-104">This article describes how to read and write to an unbounded range with the Excel JavaScript API.</span></span> <span data-ttu-id="db4d3-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="db4d3-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

<span data-ttu-id="db4d3-106">Адрес неограниченого диапазона — это адрес диапазона, который указывает целые столбцы или целые строки.</span><span class="sxs-lookup"><span data-stu-id="db4d3-106">An unbounded range address is a range address that specifies either entire columns or entire rows.</span></span> <span data-ttu-id="db4d3-107">Например:</span><span class="sxs-lookup"><span data-stu-id="db4d3-107">For example:</span></span>

- <span data-ttu-id="db4d3-108">Адреса диапазона, состоящие из целых столбцов:</span><span class="sxs-lookup"><span data-stu-id="db4d3-108">Range addresses comprised of entire columns:</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="db4d3-109">Адреса диапазона, состоящие из целых строк:</span><span class="sxs-lookup"><span data-stu-id="db4d3-109">Range addresses comprised of entire rows:</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a><span data-ttu-id="db4d3-110">Чтение из неограниченного диапазона</span><span class="sxs-lookup"><span data-stu-id="db4d3-110">Read an unbounded range</span></span>

<span data-ttu-id="db4d3-p103">Когда API отправляет запрос на получение неограниченного диапазона (например, `getRange('C:C')`), ответ будет содержать значения `null` для свойств уровня ячейки, например свойств `values`, `text`, `numberFormat` и `formula`. Другие свойства диапазона, например `address` и `cellCount`, будут содержать допустимые значения для неограниченного диапазона.</span><span class="sxs-lookup"><span data-stu-id="db4d3-p103">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

## <a name="write-to-an-unbounded-range"></a><span data-ttu-id="db4d3-113">Запись в неограниченный диапазон</span><span class="sxs-lookup"><span data-stu-id="db4d3-113">Write to an unbounded range</span></span>

<span data-ttu-id="db4d3-114">Вы не можете установить свойства уровня ячейки, такие как , и на неограниченый диапазон, так как запрос ввода `values` `numberFormat` слишком `formula` велик.</span><span class="sxs-lookup"><span data-stu-id="db4d3-114">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on an unbounded range because the input request is too large.</span></span> <span data-ttu-id="db4d3-115">Например, следующий пример кода недостоверный, так как он пытается указать для `values` неограниченого диапазона.</span><span class="sxs-lookup"><span data-stu-id="db4d3-115">For example, the following code example is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="db4d3-116">API возвращает ошибку, если вы попытаетесь установить свойства уровня ячейки для неограниченого диапазона.</span><span class="sxs-lookup"><span data-stu-id="db4d3-116">The API returns an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a><span data-ttu-id="db4d3-117">См. также</span><span class="sxs-lookup"><span data-stu-id="db4d3-117">See also</span></span>

- [<span data-ttu-id="db4d3-118">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="db4d3-118">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="db4d3-119">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="db4d3-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="db4d3-120">Чтение или написание в большом диапазоне с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="db4d3-120">Read or write to a large range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-large.md)
- [<span data-ttu-id="db4d3-121">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="db4d3-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

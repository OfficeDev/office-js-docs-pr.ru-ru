---
title: Чтение или написание в больших диапазонах с помощью API JavaScript Excel
description: Узнайте, как читать или писать в больших диапазонах с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b7a1e54d6b516889884f777bd256df8fb663c794
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652918"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a><span data-ttu-id="0c7d3-103">Чтение или написание в большом диапазоне с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="0c7d3-103">Read or write to a large range using the Excel JavaScript API</span></span>

<span data-ttu-id="0c7d3-104">В этой статье описывается, как обрабатывать чтение и запись в больших диапазонах с помощью API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="0c7d3-104">This article describes how to handle reading and writing to large ranges with the Excel JavaScript API.</span></span>

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a><span data-ttu-id="0c7d3-105">Запуск отдельных операций чтения или записи для больших диапазонов</span><span class="sxs-lookup"><span data-stu-id="0c7d3-105">Run separate read or write operations for large ranges</span></span>

<span data-ttu-id="0c7d3-106">Если диапазон содержит большое количество ячеек, значений, форматов номеров или формул, возможно, невозможно выполнить операции API на этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="0c7d3-106">If a range contains a large number of cells, values, number formats, or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="0c7d3-107">API всегда делает все возможное, чтобы выполнить запрошенную операцию над диапазоном (то есть получить или записать указанные данные), но попытка выполнить операцию чтения или записи для большого диапазона может привести к ошибке API из-за чрезмерного потребления ресурсов.</span><span class="sxs-lookup"><span data-stu-id="0c7d3-107">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="0c7d3-108">Чтобы избежать таких ошибок, мы рекомендуем выполнять отдельные операции чтения или записи для небольших подмножеств большого диапазона, а не пытаться выполнить одну операцию чтения или записи для большого диапазона.</span><span class="sxs-lookup"><span data-stu-id="0c7d3-108">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="0c7d3-109">Дополнительные сведения об ограничениях системы см. в разделе "Надстройки Excel" ограничения ресурсов и оптимизация производительности для [надстройок Office.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)</span><span class="sxs-lookup"><span data-stu-id="0c7d3-109">For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span></span>

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="0c7d3-110">Условное форматирование диапазонов</span><span class="sxs-lookup"><span data-stu-id="0c7d3-110">Conditional formatting of ranges</span></span>

<span data-ttu-id="0c7d3-111">В диапазонах может применяться форматирование к отдельным ячейкам на основе условий.</span><span class="sxs-lookup"><span data-stu-id="0c7d3-111">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="0c7d3-112">Дополнительные сведения об этом см. в статье [Применение условного форматирования к диапазонам Excel](excel-add-ins-conditional-formatting.md).</span><span class="sxs-lookup"><span data-stu-id="0c7d3-112">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0c7d3-113">См. также</span><span class="sxs-lookup"><span data-stu-id="0c7d3-113">See also</span></span>

- [<span data-ttu-id="0c7d3-114">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="0c7d3-114">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="0c7d3-115">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="0c7d3-115">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="0c7d3-116">Чтение или написание в неограниченый диапазон с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="0c7d3-116">Read or write to an unbounded range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-unbounded.md)
- [<span data-ttu-id="0c7d3-117">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="0c7d3-117">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

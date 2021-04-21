---
title: Работа с ячейками с помощью API JavaScript Excel.
description: Узнайте определение API JavaScript Excel для ячейки и узнайте, как работать с ячейками.
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917102"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="74253-103">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="74253-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="74253-104">В API JavaScript для Excel нет объекта или класса Cell.</span><span class="sxs-lookup"><span data-stu-id="74253-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="74253-105">Вместо этого все ячейки Excel являются `Range` объектами.</span><span class="sxs-lookup"><span data-stu-id="74253-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="74253-106">Отдельные ячейки в пользовательском интерфейсе Excel преобразуются в объект `Range` с одной ячейкой в API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="74253-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="74253-107">Объект `Range` также может содержать несколько соразмерных ячеек.</span><span class="sxs-lookup"><span data-stu-id="74253-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="74253-108">Дополнительные ячейки образуют неоконченный прямоугольник (включая отдельные строки или столбцы).</span><span class="sxs-lookup"><span data-stu-id="74253-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="74253-109">Чтобы узнать о работе с ячейками, которые не являются соразмерными, см. в этой ссылке Работа с дисконтными ячейками с помощью объекта [RangeAreas.](#work-with-discontiguous-cells-using-the-rangeareas-object)</span><span class="sxs-lookup"><span data-stu-id="74253-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="74253-110">Полный список свойств и методов, поддерживаемых объектом, см. в списке `Range` [Range Object (API JavaScript для Excel).](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="74253-110">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="74253-111">Работа с дисконтными ячейками с помощью объекта RangeAreas</span><span class="sxs-lookup"><span data-stu-id="74253-111">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="74253-112">Объект [RangeAreas](/javascript/api/excel/excel.rangeareas) позволяет надстройки выполнять операции сразу на нескольких диапазонах.</span><span class="sxs-lookup"><span data-stu-id="74253-112">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="74253-113">Эти диапазоны могут быть состоятельными, но они не должны быть.</span><span class="sxs-lookup"><span data-stu-id="74253-113">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="74253-114">Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="74253-114">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="74253-115">См. также</span><span class="sxs-lookup"><span data-stu-id="74253-115">See also</span></span>

- [<span data-ttu-id="74253-116">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="74253-116">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="74253-117">Получите диапазон с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="74253-117">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="74253-118">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="74253-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

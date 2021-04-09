---
title: Работа с ячейками с помощью API JavaScript Excel.
description: Узнайте определение API JavaScript Excel для ячейки и узнайте, как работать с ячейками.
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652978"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="689b6-103">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="689b6-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="689b6-104">API JavaScript Excel не имеет объекта или класса Cell.</span><span class="sxs-lookup"><span data-stu-id="689b6-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="689b6-105">Вместо этого все ячейки Excel являются `Range` объектами.</span><span class="sxs-lookup"><span data-stu-id="689b6-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="689b6-106">Индивидуальная ячейка в пользовательском интерфейсе Excel преобразуется в объект с одной ячейкой в `Range` API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="689b6-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="689b6-107">Объект `Range` также может содержать несколько соразмерных ячеек.</span><span class="sxs-lookup"><span data-stu-id="689b6-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="689b6-108">Дополнительные ячейки образуют неоконченный прямоугольник (включая отдельные строки или столбцы).</span><span class="sxs-lookup"><span data-stu-id="689b6-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="689b6-109">Чтобы узнать о работе с ячейками, которые не являются соразмерными, см. в этой ссылке Работа с дисконтными ячейками с помощью объекта [RangeAreas.](#work-with-discontiguous-cells-using-the-rangeareas-object)</span><span class="sxs-lookup"><span data-stu-id="689b6-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="689b6-110">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="689b6-110">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="excel-javascript-apis-that-mention-cells"></a><span data-ttu-id="689b6-111">API Excel JavaScript, в которых упоминаются ячейки</span><span class="sxs-lookup"><span data-stu-id="689b6-111">Excel JavaScript APIs that mention cells</span></span>

<span data-ttu-id="689b6-112">Несмотря на то, что API JavaScript Excel не имеет объекта или класса "Cell", в ряде имен API упоминаются ячейки.</span><span class="sxs-lookup"><span data-stu-id="689b6-112">Even though the Excel JavaScript API doesn't have a "Cell" object or class, a number of API names mention cells.</span></span> <span data-ttu-id="689b6-113">Эти API контролируют свойства ячейки, такие как цвет, форматирование текста и шрифт.</span><span class="sxs-lookup"><span data-stu-id="689b6-113">These APIs control cell properties like color, text formatting, and font.</span></span>

<span data-ttu-id="689b6-114">Следующий список API JavaScript Excel относится к ячейкам.</span><span class="sxs-lookup"><span data-stu-id="689b6-114">The following list of Excel JavaScript APIs refer to cells.</span></span>

- [<span data-ttu-id="689b6-115">CellBorder</span><span class="sxs-lookup"><span data-stu-id="689b6-115">CellBorder</span></span>](/javascript/api/excel/excel.cellborder)
- [<span data-ttu-id="689b6-116">CellBorderCollection</span><span class="sxs-lookup"><span data-stu-id="689b6-116">CellBorderCollection</span></span>](/javascript/api/excel/excel.cellbordercollection)
- [<span data-ttu-id="689b6-117">CellProperties</span><span class="sxs-lookup"><span data-stu-id="689b6-117">CellProperties</span></span>](/javascript/api/excel/excel.cellproperties)
- [<span data-ttu-id="689b6-118">CellPropertiesFill</span><span class="sxs-lookup"><span data-stu-id="689b6-118">CellPropertiesFill</span></span>](/javascript/api/excel/excel.cellpropertiesfill)
- [<span data-ttu-id="689b6-119">CellPropertiesFont</span><span class="sxs-lookup"><span data-stu-id="689b6-119">CellPropertiesFont</span></span>](/javascript/api/excel/excel.cellpropertiesfont)
- [<span data-ttu-id="689b6-120">CellPropertiesFormat</span><span class="sxs-lookup"><span data-stu-id="689b6-120">CellPropertiesFormat</span></span>](/javascript/api/excel/excel.cellpropertiesformat)
- [<span data-ttu-id="689b6-121">CellPropertiesProtection</span><span class="sxs-lookup"><span data-stu-id="689b6-121">CellPropertiesProtection</span></span>](/javascript/api/excel/excel.cellpropertiesprotection)
- [<span data-ttu-id="689b6-122">CellValueConditionalFormat</span><span class="sxs-lookup"><span data-stu-id="689b6-122">CellValueConditionalFormat</span></span>](/javascript/api/excel/excel.cellvalueconditionalformat)
- [<span data-ttu-id="689b6-123">ConditionalCellValueRule</span><span class="sxs-lookup"><span data-stu-id="689b6-123">ConditionalCellValueRule</span></span>](/javascript/api/excel/excel.conditionalcellvaluerule)
- [<span data-ttu-id="689b6-124">SettableCellProperties</span><span class="sxs-lookup"><span data-stu-id="689b6-124">SettableCellProperties</span></span>](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="689b6-125">Работа с дисконтными ячейками с помощью объекта RangeAreas</span><span class="sxs-lookup"><span data-stu-id="689b6-125">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="689b6-126">Объект [RangeAreas](/javascript/api/excel/excel.rangeareas) позволяет надстройки выполнять операции сразу на нескольких диапазонах.</span><span class="sxs-lookup"><span data-stu-id="689b6-126">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="689b6-127">Эти диапазоны могут быть состоятельными, но они не должны быть.</span><span class="sxs-lookup"><span data-stu-id="689b6-127">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="689b6-128">Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="689b6-128">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="689b6-129">См. также</span><span class="sxs-lookup"><span data-stu-id="689b6-129">See also</span></span>

- [<span data-ttu-id="689b6-130">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="689b6-130">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="689b6-131">Получите диапазон с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="689b6-131">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="689b6-132">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="689b6-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

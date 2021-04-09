---
title: Обработка динамических массивов и разлива диапазона с помощью API JavaScript Excel
description: Узнайте, как обрабатывать динамические массивы и разливать диапазоны с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c224fc336791440911519a6d24aee6c208d90c9e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652939"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a><span data-ttu-id="d0f38-103">Обработка динамических массивов и разлив с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="d0f38-103">Handle dynamic arrays and spilling using the Excel JavaScript API</span></span>

<span data-ttu-id="d0f38-104">В этой статье содержится пример кода, который обрабатывает динамические массивы и разлив диапазона с помощью API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="d0f38-104">This article provides a code sample that handles dynamic arrays and range spilling using the Excel JavaScript API.</span></span> <span data-ttu-id="d0f38-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="d0f38-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="dynamic-arrays"></a><span data-ttu-id="d0f38-106">Динамические массивы</span><span class="sxs-lookup"><span data-stu-id="d0f38-106">Dynamic arrays</span></span>

<span data-ttu-id="d0f38-107">Некоторые формулы Excel возвращают [динамические массивы.](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)</span><span class="sxs-lookup"><span data-stu-id="d0f38-107">Some Excel formulas return [Dynamic arrays](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span> <span data-ttu-id="d0f38-108">Они заполняют значения нескольких ячеек за пределами исходной ячейки формулы.</span><span class="sxs-lookup"><span data-stu-id="d0f38-108">These fill the values of multiple cells outside of the formula's original cell.</span></span> <span data-ttu-id="d0f38-109">Это переполнение значения называется "разлив".</span><span class="sxs-lookup"><span data-stu-id="d0f38-109">This value overflow is referred to as a "spill".</span></span> <span data-ttu-id="d0f38-110">Надстройка может найти диапазон, используемый для разлива с помощью метода [Range.getSpillingToRange.](/javascript/api/excel/excel.range#getspillingtorange--)</span><span class="sxs-lookup"><span data-stu-id="d0f38-110">Your add-in can find the range used for a spill with the [Range.getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) method.</span></span> <span data-ttu-id="d0f38-111">Существует также [версия \*OrNullObject](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .</span><span class="sxs-lookup"><span data-stu-id="d0f38-111">There is also a [\*OrNullObject version](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject`.</span></span>

<span data-ttu-id="d0f38-112">В следующем примере показана базовая формула, которая копирует содержимое диапазона в ячейку, которая разливается в соседние ячейки.</span><span class="sxs-lookup"><span data-stu-id="d0f38-112">The following sample shows a basic formula that copies the contents of a range into a cell, which spills into neighboring cells.</span></span> <span data-ttu-id="d0f38-113">Затем надстройка регистрит диапазон, содержащий разлив.</span><span class="sxs-lookup"><span data-stu-id="d0f38-113">The add-in then logs the range that contains the spill.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a><span data-ttu-id="d0f38-114">Разлиение диапазона</span><span class="sxs-lookup"><span data-stu-id="d0f38-114">Range spilling</span></span>

<span data-ttu-id="d0f38-115">Найдите ячейку, ответственную за разлив в заданную ячейку с помощью метода [Range.getSpillParent.](/javascript/api/excel/excel.range#getspillparent--)</span><span class="sxs-lookup"><span data-stu-id="d0f38-115">Find the cell responsible for spilling into a given cell by using the [Range.getSpillParent](/javascript/api/excel/excel.range#getspillparent--) method.</span></span> <span data-ttu-id="d0f38-116">Обратите `getSpillParent` внимание, что работает только в том случае, если объект диапазона является одной ячейкой.</span><span class="sxs-lookup"><span data-stu-id="d0f38-116">Note that `getSpillParent` only works when the range object is a single cell.</span></span> <span data-ttu-id="d0f38-117">Вызов диапазона с несколькими ячейками приведет к ошибке, которая будет выброшена (или возвращается диапазон `getSpillParent` `Range.getSpillParentOrNullObject` null).</span><span class="sxs-lookup"><span data-stu-id="d0f38-117">Calling `getSpillParent` on a range with multiple cells will result in an error being thrown (or a null range being returned for `Range.getSpillParentOrNullObject`).</span></span>

## <a name="see-also"></a><span data-ttu-id="d0f38-118">См. также</span><span class="sxs-lookup"><span data-stu-id="d0f38-118">See also</span></span>

- [<span data-ttu-id="d0f38-119">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="d0f38-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d0f38-120">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="d0f38-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="d0f38-121">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="d0f38-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

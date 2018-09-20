---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016460"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="519cc-102">Работа с несколькими диапазонами одновременно в надстройках Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="519cc-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="519cc-103">Библиотека Excel JavaScript позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов.</span><span class="sxs-lookup"><span data-stu-id="519cc-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="519cc-104">Диапазоны не должны быть непрерывными.</span><span class="sxs-lookup"><span data-stu-id="519cc-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="519cc-105">В дополнение к упрощению вашего кода этот способ установки свойства выполняется намного быстрее, чем установка этого же свойства индивидуально для каждого из диапазонов.</span><span class="sxs-lookup"><span data-stu-id="519cc-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="519cc-106">Для API-интерфейсов, описанных в этой статье, требуется **версия Office 2016 Click-to-Run 1809 сборки 10820.20000** или более поздняя версия.</span><span class="sxs-lookup"><span data-stu-id="519cc-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="519cc-107">(Возможно, вам потребуется присоединиться к [программе предварительной оценки Office](https://products.office.com/office-insider) для получения соответствующей сборки.) Кроме того, необходимо загрузить бета-версию библиотеки Office JavaScript из [сети CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="519cc-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="519cc-108">К тому же, у нас еще нет страниц со ссылкой на эти API.</span><span class="sxs-lookup"><span data-stu-id="519cc-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="519cc-109">Но следующий файл типа определения содержит описания для них: [office.d.ts бета-версии](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="519cc-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="519cc-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="519cc-110">RangeAreas</span></span>

<span data-ttu-id="519cc-111">Набор диапазонов (возможно, разобщенных) представлен объектом `Excel.RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="519cc-112">Он имеет свойства и методы, аналогичные типу `Range` (многие из которых имеют одинаковые или похожие имена), но изменения были внесены в:</span><span class="sxs-lookup"><span data-stu-id="519cc-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="519cc-113">Типы данных для свойств и поведений методов задания и методов получения.</span><span class="sxs-lookup"><span data-stu-id="519cc-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="519cc-114">Типы данных параметров метода и поведений метода.</span><span class="sxs-lookup"><span data-stu-id="519cc-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="519cc-115">Типы данных возвращаемых значений метода.</span><span class="sxs-lookup"><span data-stu-id="519cc-115">The data types of method return values.</span></span>

<span data-ttu-id="519cc-116">Некоторые примеры:</span><span class="sxs-lookup"><span data-stu-id="519cc-116">Some examples:</span></span>

- <span data-ttu-id="519cc-117">`RangeAreas` имеет свойство `address`, которое возвращает строку с адресами диапазона, разделенными диапазонами, а не только один адрес, как в случае со свойством `Range.address`.</span><span class="sxs-lookup"><span data-stu-id="519cc-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="519cc-118">`RangeAreas` имеет свойство `dataValidation`, которое возвращает объект`DataValidation`, представляющий проверку данных всех диапазонов в `RangeAreas`при соответствии.</span><span class="sxs-lookup"><span data-stu-id="519cc-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="519cc-119">Свойство является`null`, если идентичные объекты `DataValidation` не применяются к каждому из всех диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="519cc-120">Это общие, но не универсальные принципы для объекта`RangeAreas`: *если свойство не имеет согласованных значений для каждого из всех диапазонов в `RangeAreas`, тогда оно является `null`.*</span><span class="sxs-lookup"><span data-stu-id="519cc-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="519cc-121">См. [свойства чтения RangeAreas](#reading-properties-of-rangeareas), чтобы ознакомиться с дополнительными сведениями и некоторыми исключениями.</span><span class="sxs-lookup"><span data-stu-id="519cc-121">See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="519cc-122">`RangeAreas.cellCount` возвращает общее число ячеек во все диапазоны в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="519cc-123">`RangeAreas.calculate` пересчитывает ячейки всех диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="519cc-124">`RangeAreas.getEntireColumn` и `RangeAreas.getEntireRow` возвращают другой объект `RangeAreas`, представляющий все столбцы (или строки) во всех диапазонах в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="519cc-125">Например, если `RangeAreas` представляет "A1:C4" и "F14:L15", то `RangeAreas.getEntireColumn` возвращает объект `RangeAreas`, представляющий "A:C" и "F:L".</span><span class="sxs-lookup"><span data-stu-id="519cc-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="519cc-126">`RangeAreas.copyFrom` можно использовать параметр `Range` или `RangeAreas`, представляющий диапазон(ы) источника операции копирования.</span><span class="sxs-lookup"><span data-stu-id="519cc-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="519cc-127">Полный список элементов диапазона Range, которые также доступны на RangeAreas</span><span class="sxs-lookup"><span data-stu-id="519cc-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="519cc-128">Свойства</span><span class="sxs-lookup"><span data-stu-id="519cc-128">Properties</span></span>

<span data-ttu-id="519cc-129">Ознакомьтесь с [Чтением свойств RangeAreas](#reading-properties-of-rangeareas) до написания кода, который считывает все свойства из списка.</span><span class="sxs-lookup"><span data-stu-id="519cc-129">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="519cc-130">Существуют подзаголовки к тому, что будет возвращено.</span><span class="sxs-lookup"><span data-stu-id="519cc-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="519cc-131">address</span><span class="sxs-lookup"><span data-stu-id="519cc-131">address</span></span>
- <span data-ttu-id="519cc-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="519cc-132">addressLocal</span></span>
- <span data-ttu-id="519cc-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="519cc-133">cellCount</span></span>
- <span data-ttu-id="519cc-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="519cc-134">conditionalFormats</span></span>
- <span data-ttu-id="519cc-135">context</span><span class="sxs-lookup"><span data-stu-id="519cc-135">context</span></span>
- <span data-ttu-id="519cc-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="519cc-136">dataValidation</span></span>
- <span data-ttu-id="519cc-137">format</span><span class="sxs-lookup"><span data-stu-id="519cc-137">format</span></span>
- <span data-ttu-id="519cc-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="519cc-138">isEntireColumn</span></span>
- <span data-ttu-id="519cc-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="519cc-139">isEntireRow</span></span>
- <span data-ttu-id="519cc-140">стиль</span><span class="sxs-lookup"><span data-stu-id="519cc-140">style</span></span>
- <span data-ttu-id="519cc-141">лист</span><span class="sxs-lookup"><span data-stu-id="519cc-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="519cc-142">Методы</span><span class="sxs-lookup"><span data-stu-id="519cc-142">Methods</span></span>

<span data-ttu-id="519cc-143">Помеченные методы диапазона в режиме предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="519cc-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="519cc-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="519cc-144">calculate()</span></span>
- <span data-ttu-id="519cc-145">clear()</span><span class="sxs-lookup"><span data-stu-id="519cc-145">clear()</span></span>
- <span data-ttu-id="519cc-146">convertDataTypeToText() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="519cc-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="519cc-147">convertToLinkedDataType() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="519cc-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="519cc-148">copyFrom() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="519cc-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="519cc-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="519cc-149">getEntireColumn()</span></span>
- <span data-ttu-id="519cc-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="519cc-150">getEntireRow()</span></span>
- <span data-ttu-id="519cc-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="519cc-151">getIntersection()</span></span>
- <span data-ttu-id="519cc-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="519cc-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="519cc-153">getOffsetRange() (с именем getOffsetRangeAreas на объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="519cc-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="519cc-154">getSpecialCells() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="519cc-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="519cc-155">getSpecialCellsOrNullObject() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="519cc-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="519cc-156">getTables() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="519cc-156">getTables() (preview)</span></span>
- <span data-ttu-id="519cc-157">getUsedRange() (с именем getUsedRangeAreas на объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="519cc-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="519cc-158">getUsedRangeOrNullObject() (с именем getUsedRangeAreasOrNullObject на объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="519cc-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="519cc-159">load()</span><span class="sxs-lookup"><span data-stu-id="519cc-159">load()</span></span>
- <span data-ttu-id="519cc-160">set()</span><span class="sxs-lookup"><span data-stu-id="519cc-160">set\*</span></span>
- <span data-ttu-id="519cc-161">setDirty() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="519cc-161">setDirty() (preview)</span></span>
- <span data-ttu-id="519cc-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="519cc-162">toJSON()</span></span>
- <span data-ttu-id="519cc-163">track()</span><span class="sxs-lookup"><span data-stu-id="519cc-163">track</span></span>
- <span data-ttu-id="519cc-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="519cc-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="519cc-165">Свойства и методы, характерные для объекта RangeArea</span><span class="sxs-lookup"><span data-stu-id="519cc-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="519cc-166">Тип `RangeAreas` имеет некоторые свойства и методы, которые не входят в объект `Range`.</span><span class="sxs-lookup"><span data-stu-id="519cc-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object:</span></span> <span data-ttu-id="519cc-167">Ниже приведено их выделение:</span><span class="sxs-lookup"><span data-stu-id="519cc-167">The following is a selection of them:</span></span>

- <span data-ttu-id="519cc-168">`areas`: Объект `RangeCollection`, содержащий все диапазоны, представленные объектом `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="519cc-169">Объект `RangeCollection` – также новый и аналогичен другим объектам коллекции Excel.</span><span class="sxs-lookup"><span data-stu-id="519cc-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="519cc-170">Он имеет свойство `items`, которое представляет собой массив из объектов `Range`, представляющих диапазоны.</span><span class="sxs-lookup"><span data-stu-id="519cc-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="519cc-171">`areaCount`: Общее число диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-171">The total number of recipients in the message.</span></span>
- <span data-ttu-id="519cc-172">`getOffsetRangeAreas`: Работает так же, как [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), за исключением того, что `RangeAreas` возвращается и содержит диапазоны, каждый из которых смещен от одного из диапазонов в исходном `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="519cc-173">Создание RangeAreas и установка свойств</span><span class="sxs-lookup"><span data-stu-id="519cc-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="519cc-174">Можно создать объект `RangeAreas` двумя основными способами:</span><span class="sxs-lookup"><span data-stu-id="519cc-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="519cc-175">Вызовите `Worksheet.getRanges()` и передайте его в строку с адресами диапазона, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="519cc-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="519cc-176">Если диапазон, который вы хотите включить, был переделан в [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), вы можете включить в строку имя вместо адреса.</span><span class="sxs-lookup"><span data-stu-id="519cc-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="519cc-177">вызова метода `Workbook.getSelectedRanges()`;</span><span class="sxs-lookup"><span data-stu-id="519cc-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="519cc-178">Этот метод возвращает `RangeAreas`, представляющий все диапазоны, выбранные на активном в данный момент листе.</span><span class="sxs-lookup"><span data-stu-id="519cc-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="519cc-179">После получения объекта `RangeAreas` можно создать другие с помощью методов, применяемых к объекту, который возвращает `RangeAreas`, такие как `getOffsetRangeAreas` и `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="519cc-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="519cc-180">Нельзя непосредственно добавить дополнительные диапазоны для объекта `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="519cc-181">Например, коллекция в `RangeAreas.areas` не имеет метода `add`.</span><span class="sxs-lookup"><span data-stu-id="519cc-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="519cc-182">Не пытайтесь непосредственно добавлять или удалять элементы из массива `RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="519cc-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="519cc-183">Это приведет к нежелательному поведению в вашем коде.</span><span class="sxs-lookup"><span data-stu-id="519cc-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="519cc-184">К примеру, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, так как свойства и методы `RangeAreas` функционируют так, как если бы новый элемент не был добавлен.</span><span class="sxs-lookup"><span data-stu-id="519cc-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="519cc-185">Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)` вызывает ошибку, если `index` больше, чем `areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="519cc-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="519cc-186">Аналогично, удаление объекта `Range` в диапазоне `RangeAreas.areas.items` путем получения ссылки на него и вызов его метода `Range.delete` вызывает ошибки: хотя объект `Range`*будет* удален, свойства и методы родительского объекта `RangeAreas` будут вести себя так, как если бы он все еще присутствовал (или будут стремиться к таком поведению).</span><span class="sxs-lookup"><span data-stu-id="519cc-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="519cc-187">Например, если код вызывает метод `RangeAreas.calculate`, Office будет пытаться рассчитать диапазон, но это завершится ошибкой, поскольку отсутствует объект range.</span><span class="sxs-lookup"><span data-stu-id="519cc-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="519cc-188">Установка свойства для `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="519cc-189">Ниже приведен пример установки свойства для нескольких диапазонов.</span><span class="sxs-lookup"><span data-stu-id="519cc-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="519cc-190">Функция выделяет диапазоны **F3:F5** и **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="519cc-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="519cc-191">Этот пример применяется к сценариям, в которых можно создать серьезный код адресов диапазона, передаваемых в `getRanges`, или легко рассчитать их во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="519cc-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="519cc-192">Ниже перечислены некоторые сценарии, в которых это возможно:</span><span class="sxs-lookup"><span data-stu-id="519cc-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="519cc-193">Код выполняется в контексте известного шаблона.</span><span class="sxs-lookup"><span data-stu-id="519cc-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="519cc-194">Код выполняется в контексте импортированных данных, в котором известна схема данных.</span><span class="sxs-lookup"><span data-stu-id="519cc-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="519cc-195">Когда во время создания кода не известно, с какими диапазонами вам придется работать, необходимо обнаружить их во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="519cc-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="519cc-196">В следующем разделе описываются эти сценарии.</span><span class="sxs-lookup"><span data-stu-id="519cc-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="519cc-197">Обнаружение областей диапазона с помощью программных средств</span><span class="sxs-lookup"><span data-stu-id="519cc-197">Discover range areas programmatically</span></span>

<span data-ttu-id="519cc-198">Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` можно использовать для поиска во время выполнения диапазонов, с которыми вы хотите работать, на основе характеристик ячеек и типа значений в ячейках.</span><span class="sxs-lookup"><span data-stu-id="519cc-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="519cc-199">Вот подписи методов из файла типов данных TypeScript:</span><span class="sxs-lookup"><span data-stu-id="519cc-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="519cc-200">Ниже приведен пример об использовании первого из них.</span><span class="sxs-lookup"><span data-stu-id="519cc-200">The following is an example of using the "Between" operator:</span></span> <span data-ttu-id="519cc-201">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="519cc-201">About this code, note:</span></span>

- <span data-ttu-id="519cc-202">Он ограничивает часть листа, которую нужно искать, вызвав сначала `Worksheet.getUsedRange`, а затем вызвав `getSpecialCells` только для этого диапазона.</span><span class="sxs-lookup"><span data-stu-id="519cc-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="519cc-203">Передает в качестве параметра для `getSpecialCells` версию строки значения из перечисления `Excel.SpecialCellType`.</span><span class="sxs-lookup"><span data-stu-id="519cc-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="519cc-204">Некоторые другие значения, которые могут быть переданы вместо этого, – это "Blanks" для пустых ячеек, "Constants" для ячейки со значениями литералов вместо формул и "SameConditionalFormat" для ячеек, у которых такое же условное форматирование, как и у первой ячейки в `usedRange`.</span><span class="sxs-lookup"><span data-stu-id="519cc-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="519cc-205">Первая ячейка является верхней крайней слева ячейкой.</span><span class="sxs-lookup"><span data-stu-id="519cc-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="519cc-206">Полный список значений перечисления см. в [office.d.ts бета-версии](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="519cc-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="519cc-207">Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами залиты розовым цветом даже в том случае, если не все они непрерывны.</span><span class="sxs-lookup"><span data-stu-id="519cc-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="519cc-208">В некоторых случаях диапазон не содержит*ни одну* из ячеек с целевой характеристикой.</span><span class="sxs-lookup"><span data-stu-id="519cc-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="519cc-209">Если `getSpecialCells` не находит требуемой ячейки, он вызывает ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="519cc-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="519cc-210">Это будет переадресовать поток управления к блоку или методу `catch`, если он существует.</span><span class="sxs-lookup"><span data-stu-id="519cc-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="519cc-211">Если нет, ошибка будет останавливать функцию.</span><span class="sxs-lookup"><span data-stu-id="519cc-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="519cc-212">Могут быть сценарии, в которых выдача ошибки – это именно то, что должно происходить при отсутствуют ячейки с целевой характеристикой.</span><span class="sxs-lookup"><span data-stu-id="519cc-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="519cc-213">Однако в сценариях, для которых это нормально, но, возможно, необычно, может не оказаться соответствующих ячеек. Ваш код должен проверить наличие такой возможности и аккуратно провести работу с сценарием без выдачи ошибки.</span><span class="sxs-lookup"><span data-stu-id="519cc-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="519cc-214">Для этих сценариев следует использовать метод `getSpecialCellsOrNullObject` и протестировать свойство `RangeAreas.isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="519cc-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="519cc-215">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="519cc-215">The following is an example.</span></span> <span data-ttu-id="519cc-216">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="519cc-216">Note about this code:</span></span>

- <span data-ttu-id="519cc-217">Метод `getSpecialCellsOrNullObject` всегда возвращает объект прокси-сервера, поэтому он не может быть `null` в обычном смысле JavaScript.</span><span class="sxs-lookup"><span data-stu-id="519cc-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="519cc-218">Но если не обнаружено соответствующих ячеек, свойству `isNullObject` объекта присваивается значение `true`.</span><span class="sxs-lookup"><span data-stu-id="519cc-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="519cc-219">Оно вызывает `context.sync` *прежде*, чем протестировать свойство `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="519cc-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="519cc-220">Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство для его чтения.</span><span class="sxs-lookup"><span data-stu-id="519cc-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="519cc-221">Тем не менее, не требуется *явным образом* нагружать свойство `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="519cc-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="519cc-222">Он автоматически загружается с `context.sync` даже в том случае, если `load` не вызывается в объекте.</span><span class="sxs-lookup"><span data-stu-id="519cc-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="519cc-223">Для получения дополнительных сведений см. [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="519cc-223">For more information see [\*](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)</span></span>
- <span data-ttu-id="519cc-224">Этот код можно проверить, выбрав сначала диапазон, у которого нет ячеек формулы, и запустив его.</span><span class="sxs-lookup"><span data-stu-id="519cc-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="519cc-225">Выберите диапазон, который содержит по крайней мере одну ячейку с формулой, и снова запустите его.</span><span class="sxs-lookup"><span data-stu-id="519cc-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

<span data-ttu-id="519cc-226">Для простоты во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="519cc-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="519cc-227">Сужение целевых ячеек с типом значений ячеек</span><span class="sxs-lookup"><span data-stu-id="519cc-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="519cc-228">Есть также необязательный второй параметр типа перечисления `Excel.SpecialCellValueType`, который в дальнейшем сужает ячейки до целевого объекта.</span><span class="sxs-lookup"><span data-stu-id="519cc-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="519cc-229">Можно использовать его только в том случае, если передается значение "Formulas" или "Constants" для `getSpecialCells` или `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="519cc-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="519cc-230">Этот параметр указывает, что требуются только ячейки с определенными типами значений.</span><span class="sxs-lookup"><span data-stu-id="519cc-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="519cc-231">Существует четыре основных типа: "Error", "Logical" (то же самое, что и boolean- логический), "Numbers" и "Text".</span><span class="sxs-lookup"><span data-stu-id="519cc-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="519cc-232">(Перечисление имеет другие значения помимо этих четырех, которые рассматриваются ниже.) Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="519cc-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="519cc-233">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="519cc-233">About this code, note:</span></span>

- <span data-ttu-id="519cc-234">Он будет выделять только ячейки, имеющие числовое значение литерала.</span><span class="sxs-lookup"><span data-stu-id="519cc-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="519cc-235">Не выделяет ячейки, в которых содержится формула (даже в том случае, если результат является числом), логическое, текстовое значение, или ячейки с состоянием ошибки.</span><span class="sxs-lookup"><span data-stu-id="519cc-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="519cc-236">Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.</span><span class="sxs-lookup"><span data-stu-id="519cc-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="519cc-237">В некоторых случаях вам требуется работать с более чем одним типом значения ячейки, например, все ячейки с текстовыми значениями и все ячейки с логическими значениями ("Logical").</span><span class="sxs-lookup"><span data-stu-id="519cc-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="519cc-238">Перечисление `Excel.SpecialCellValueType` содержит значения, которые позволяют объединять типы.</span><span class="sxs-lookup"><span data-stu-id="519cc-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="519cc-239">Например, "LogicalText" будет обрабатывать все логические и все текстовые ячейки.</span><span class="sxs-lookup"><span data-stu-id="519cc-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="519cc-240">Можно использовать любые два или три из четырех основных типов.</span><span class="sxs-lookup"><span data-stu-id="519cc-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="519cc-241">Имена этих значений перечисления, которые объединяют основные типы, всегда находятся в алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="519cc-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="519cc-242">Таким образом, для объединения ячеек со значениями ошибок, текстовыми и логическими значениями используйте "ErrorLogicalText", а не "LogicalErrorText" или "TextErrorLogical".</span><span class="sxs-lookup"><span data-stu-id="519cc-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="519cc-243">Параметр по умолчанию "All" объединяет все четыре типа.</span><span class="sxs-lookup"><span data-stu-id="519cc-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="519cc-244">В следующем примере выделены все ячейки с формулами, которые производят числовые или логические значения:</span><span class="sxs-lookup"><span data-stu-id="519cc-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> <span data-ttu-id="519cc-245">Параметр `Excel.SpecialCellValueType` можно использовать, только если параметр `Excel.SpecialCellType` – это "Formulas" или "Constants".</span><span class="sxs-lookup"><span data-stu-id="519cc-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="519cc-246">Получение объектов RangeAreas в RangeAreas</span><span class="sxs-lookup"><span data-stu-id="519cc-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="519cc-247">Тип `RangeAreas` также имеет методы `getSpecialCells` и `getSpecialCellsOrNullObject`, которые принимают те же два параметра.</span><span class="sxs-lookup"><span data-stu-id="519cc-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="519cc-248">Эти методы возвращают все целевые ячейки из всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="519cc-249">Существует одно небольшое отличие в поведении методов при вызове объекта `RangeAreas` вместо объекта `Range`: когда вы передаете "SameConditionalFormat" в качестве первого параметра, метод возвращает все ячейки, имеющие одинаковое условное форматирование, в качестве верхней крайней слева ячейки *в первом диапазоне в коллекции `RangeAreas.areas`*.</span><span class="sxs-lookup"><span data-stu-id="519cc-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="519cc-250">То же касается и "SameDataValidation": при передаче к`Range.getSpecialCells` он возвращает все ячейки, которые имеют такое же правило проверки данных, в качестве верхней крайней слева ячейки *в диапазоне*.</span><span class="sxs-lookup"><span data-stu-id="519cc-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="519cc-251">Но при передаче к `RangeAreas.getSpecialCells` она возвращает все ячейки, которые имеют такое же правило проверки данных, в качестве верхней крайней слева ячейки *в первый диапазон в коллекции `RangeAreas.areas`*.</span><span class="sxs-lookup"><span data-stu-id="519cc-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="519cc-252">Чтение свойств RangeAreas</span><span class="sxs-lookup"><span data-stu-id="519cc-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="519cc-253">Чтение значения свойств `RangeAreas` требует осторожности, так как данное свойство может иметь разные значения для разных диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="519cc-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="519cc-254">Общее правило заключается в том, что если соответствующее значение *может* быть возвращено, оно будет возвращено.</span><span class="sxs-lookup"><span data-stu-id="519cc-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="519cc-255">Например, в следующем коде RGB-код для розовой заливки (`#FFC0CB`) и `true` будет выполнять вход в консоль, так как оба диапазона в объекте `RangeAreas` имеют розовую заливку и оба являются целыми столбцами.</span><span class="sxs-lookup"><span data-stu-id="519cc-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

<span data-ttu-id="519cc-256">Все усложняется, когда согласованность невозможна.</span><span class="sxs-lookup"><span data-stu-id="519cc-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="519cc-257">Поведение свойств `RangeAreas` следует следующим тремя принципам:</span><span class="sxs-lookup"><span data-stu-id="519cc-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="519cc-258">Логическое свойство объекта `RangeAreas` возвращает `false`, кроме случаев, когда свойство имеет значение true для всех диапазонов элементов.</span><span class="sxs-lookup"><span data-stu-id="519cc-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="519cc-259">Свойства, не являющиеся логическими, за исключением свойства `address`, возвращают `null`, кроме тех случаев, когда соответствующее свойство для всех элементов диапазона обладает тем же значением.</span><span class="sxs-lookup"><span data-stu-id="519cc-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="519cc-260">Свойство `address` возвращает строку с адресами диапазонов элементов, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="519cc-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="519cc-261">Например, следующий код создает `RangeAreas`, в котором только один диапазон — это целый столбец, и только один заполнен розовым цветом.</span><span class="sxs-lookup"><span data-stu-id="519cc-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="519cc-262">Консоль покажет `null` для цвета заливки, `false` для свойства `isEntireRow` и "Sheet1! F3:F5 Sheet1! H:H" (при условии, что имя листа – это "Sheet1") для свойства`address`.</span><span class="sxs-lookup"><span data-stu-id="519cc-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a><span data-ttu-id="519cc-263">См. также</span><span class="sxs-lookup"><span data-stu-id="519cc-263">See also</span></span>

- [<span data-ttu-id="519cc-264">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="519cc-264">Excel JavaScript API core concepts</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="519cc-265">Объект Range (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="519cc-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="519cc-266">Объект[RangeAreas (JavaScript API для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Эта ссылка может не работать, пока API находится в режиме предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="519cc-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="519cc-267">В качестве альтернативы см. [office.d.ts бета-версии](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span><span class="sxs-lookup"><span data-stu-id="519cc-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>
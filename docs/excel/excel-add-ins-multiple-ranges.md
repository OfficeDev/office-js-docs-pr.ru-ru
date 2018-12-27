---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: ''
ms.date: 12/26/2018
ms.openlocfilehash: ab7cd9757adaedf2b6cc43fdcc604b98a60b6ecd
ms.sourcegitcommit: 8d248cd890dae1e9e8ef1bd47e09db4c1cf69593
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/27/2018
ms.locfileid: "27447234"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="c4323-102">Работа с несколькими диапазонами одновременно в надстройках Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c4323-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="c4323-103">Библиотека JavaScript для Excel позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов.</span><span class="sxs-lookup"><span data-stu-id="c4323-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="c4323-104">Диапазоны необязательно должны быть смежными.</span><span class="sxs-lookup"><span data-stu-id="c4323-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="c4323-105">Этот способ установки свойства не только упрощает код, но и выполняется намного быстрее, чем установка одинакового свойства отдельно для каждого диапазона.</span><span class="sxs-lookup"><span data-stu-id="c4323-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="c4323-106">Для работы с API-интерфейсами, описанными в этой статье, требуется **Office 2016 "нажми и работай" версии 1809 сборки 10820.20000** или более поздней версии</span><span class="sxs-lookup"><span data-stu-id="c4323-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="c4323-107">(возможно, вам придется принять участие в [программе предварительной оценки Office](https://products.office.com/office-insider) для получения нужной сборки). Кроме того, необходимо загрузить бета-версию библиотеки JavaScript для Office из сети [CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="c4323-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="c4323-108">В настоящее время нет справочных страниц для этих API.</span><span class="sxs-lookup"><span data-stu-id="c4323-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="c4323-109">Но следующий файл типа определения содержит их описания: [бета-версия office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="c4323-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="c4323-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c4323-110">RangeAreas</span></span>

<span data-ttu-id="c4323-111">Набор диапазонов (возможно, несмежных) представлен объектом `Excel.RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="c4323-112">Его свойства и методы аналогичны типу `Range` (многие с одинаковыми или похожими именами), но с изменением указанных ниже параметров:</span><span class="sxs-lookup"><span data-stu-id="c4323-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="c4323-113">Типы данных для свойств и поведений методов задания и методов получения.</span><span class="sxs-lookup"><span data-stu-id="c4323-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="c4323-114">Типы данных параметров метода и поведений метода.</span><span class="sxs-lookup"><span data-stu-id="c4323-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="c4323-115">Типы данных возвращаемых значений метода.</span><span class="sxs-lookup"><span data-stu-id="c4323-115">The data types of method return values.</span></span>

<span data-ttu-id="c4323-116">Примеры:</span><span class="sxs-lookup"><span data-stu-id="c4323-116">Some examples:</span></span>

- <span data-ttu-id="c4323-117">У `RangeAreas` есть свойство `address`, возвращающее строку с адресами диапазона, разделенными запятой, а не только один адрес, как в случае со свойством `Range.address`.</span><span class="sxs-lookup"><span data-stu-id="c4323-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="c4323-118">У `RangeAreas` есть свойство `dataValidation`, которое возвращает объект `DataValidation`, представляющий проверку данных всех диапазонов в `RangeAreas` при соответствии.</span><span class="sxs-lookup"><span data-stu-id="c4323-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="c4323-119">Значение этого свойства будет равно `null`, если ко всем диапазонам в `RangeAreas` не применяются одинаковые объекты `DataValidation`.</span><span class="sxs-lookup"><span data-stu-id="c4323-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="c4323-120">Это общий, но не универсальный принцип для объекта `RangeAreas`: *если у свойства нет согласованных значений во всех диапазонах в `RangeAreas`, его значением будет `null`.*</span><span class="sxs-lookup"><span data-stu-id="c4323-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="c4323-121">Дополнительные сведения и некоторые исключения см. в разделе [Чтение свойств RangeAreas](#read-properties-of-rangeareas).</span><span class="sxs-lookup"><span data-stu-id="c4323-121">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="c4323-122">`RangeAreas.cellCount` получает общее количество ячеек во всех диапазонах в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="c4323-123">`RangeAreas.calculate` пересчитывает ячейки всех диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="c4323-124">`RangeAreas.getEntireColumn` и `RangeAreas.getEntireRow` возвращают другой объект `RangeAreas`, представляющий все столбцы (или строки) во всех диапазонах в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="c4323-125">Например, если `RangeAreas` представляет "A1:C4" и "F14:L15", то `RangeAreas.getEntireColumn` возвращает объект `RangeAreas`, представляющий "A:C" и "F:L".</span><span class="sxs-lookup"><span data-stu-id="c4323-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="c4323-126">`RangeAreas.copyFrom` может использовать параметр `Range` или `RangeAreas`, представляющий исходный диапазон (или диапазоны) операции копирования.</span><span class="sxs-lookup"><span data-stu-id="c4323-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="c4323-127">Полный список элементов Range, также доступных в RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c4323-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="c4323-128">Свойства</span><span class="sxs-lookup"><span data-stu-id="c4323-128">Properties</span></span>

<span data-ttu-id="c4323-129">Ознакомьтесь со статьей [Чтение свойств RangeAreas](#read-properties-of-rangeareas) перед написанием кода, считывающего любое из перечисленных свойств.</span><span class="sxs-lookup"><span data-stu-id="c4323-129">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="c4323-130">Возвращаемое значение зависит от ряда факторов.</span><span class="sxs-lookup"><span data-stu-id="c4323-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="c4323-131">address</span><span class="sxs-lookup"><span data-stu-id="c4323-131">address</span></span>
- <span data-ttu-id="c4323-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="c4323-132">addressLocal</span></span>
- <span data-ttu-id="c4323-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="c4323-133">cellCount</span></span>
- <span data-ttu-id="c4323-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="c4323-134">conditionalFormats</span></span>
- <span data-ttu-id="c4323-135">context</span><span class="sxs-lookup"><span data-stu-id="c4323-135">context</span></span>
- <span data-ttu-id="c4323-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="c4323-136">dataValidation</span></span>
- <span data-ttu-id="c4323-137">format</span><span class="sxs-lookup"><span data-stu-id="c4323-137">format</span></span>
- <span data-ttu-id="c4323-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="c4323-138">isEntireColumn</span></span>
- <span data-ttu-id="c4323-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="c4323-139">isEntireRow</span></span>
- <span data-ttu-id="c4323-140">style</span><span class="sxs-lookup"><span data-stu-id="c4323-140">style</span></span>
- <span data-ttu-id="c4323-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="c4323-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="c4323-142">Методы</span><span class="sxs-lookup"><span data-stu-id="c4323-142">Methods</span></span>

<span data-ttu-id="c4323-143">Методы Range в предварительной версии помечены.</span><span class="sxs-lookup"><span data-stu-id="c4323-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="c4323-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="c4323-144">calculate()</span></span>
- <span data-ttu-id="c4323-145">clear()</span><span class="sxs-lookup"><span data-stu-id="c4323-145">clear()</span></span>
- <span data-ttu-id="c4323-146">convertDataTypeToText() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c4323-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="c4323-147">convertToLinkedDataType() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c4323-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="c4323-148">copyFrom() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c4323-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="c4323-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="c4323-149">getEntireColumn()</span></span>
- <span data-ttu-id="c4323-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="c4323-150">getEntireRow()</span></span>
- <span data-ttu-id="c4323-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="c4323-151">getIntersection()</span></span>
- <span data-ttu-id="c4323-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="c4323-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="c4323-153">getOffsetRange() (называется getOffsetRangeAreas в объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="c4323-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="c4323-154">getSpecialCells() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c4323-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="c4323-155">getSpecialCellsOrNullObject() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c4323-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="c4323-156">getTables() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c4323-156">getTables() (preview)</span></span>
- <span data-ttu-id="c4323-157">getUsedRange() (называется getUsedRangeAreas в объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="c4323-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="c4323-158">getUsedRangeOrNullObject() (называется getUsedRangeAreasOrNullObject в объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="c4323-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="c4323-159">load()</span><span class="sxs-lookup"><span data-stu-id="c4323-159">load()</span></span>
- <span data-ttu-id="c4323-160">set()</span><span class="sxs-lookup"><span data-stu-id="c4323-160">set()</span></span>
- <span data-ttu-id="c4323-161">setDirty() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c4323-161">setDirty() (preview)</span></span>
- <span data-ttu-id="c4323-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="c4323-162">toJSON()</span></span>
- <span data-ttu-id="c4323-163">track()</span><span class="sxs-lookup"><span data-stu-id="c4323-163">track()</span></span>
- <span data-ttu-id="c4323-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="c4323-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="c4323-165">Свойства и методы, характерные для объекта RangeArea</span><span class="sxs-lookup"><span data-stu-id="c4323-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="c4323-166">Для типа `RangeAreas` существуют несколько свойств и методов, отсутствующих в объекте `Range`.</span><span class="sxs-lookup"><span data-stu-id="c4323-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="c4323-167">Ниже приведены некоторые из них.</span><span class="sxs-lookup"><span data-stu-id="c4323-167">The following is a selection of them:</span></span>

- <span data-ttu-id="c4323-168">`areas`. Объект `RangeCollection`, содержащий все диапазоны, которые представлены объектом `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="c4323-169">Объект `RangeCollection` — еще один новый объект, аналогичный другим объектам коллекции Excel.</span><span class="sxs-lookup"><span data-stu-id="c4323-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="c4323-170">У него есть свойство `items`, являющееся массивом объектов `Range`, которые представляют диапазоны.</span><span class="sxs-lookup"><span data-stu-id="c4323-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="c4323-171">`areaCount`. Общее количество диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="c4323-172">`getOffsetRangeAreas`. Действует аналогично методу [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), за исключением того, что возвращается объект `RangeAreas`, содержащий диапазоны, каждый из которых смещен относительно одного из диапазонов в исходном объекте `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="c4323-173">Создание RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c4323-173">Create RangeAreas</span></span>

<span data-ttu-id="c4323-174">Объект `RangeAreas` можно создать двумя основными способами:</span><span class="sxs-lookup"><span data-stu-id="c4323-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="c4323-175">Вызвать метод `Worksheet.getRanges()` и передать ему строку с адресами диапазона, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="c4323-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="c4323-176">Если диапазон, который нужно включить, был преобразован [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), в строку можно включить имя вместо адреса.</span><span class="sxs-lookup"><span data-stu-id="c4323-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="c4323-177">Вызвать метод `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="c4323-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="c4323-178">Этот метод возвращает объект `RangeAreas`, представляющий все диапазоны, выбранные в активном на данный момент листе.</span><span class="sxs-lookup"><span data-stu-id="c4323-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="c4323-179">После получения объекта `RangeAreas` можно создать другие с помощью методов объекта, возвращающих `RangeAreas`, например `getOffsetRangeAreas` и `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="c4323-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="c4323-180">Нельзя напрямую добавить дополнительные диапазоны к объекту `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="c4323-181">Например, у коллекции в `RangeAreas.areas` нет метода `add`.</span><span class="sxs-lookup"><span data-stu-id="c4323-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="c4323-182">Не пытайтесь напрямую добавлять или удалять элементы из массива `RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="c4323-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="c4323-183">Это приведет к нежелательному поведению кода. </span><span class="sxs-lookup"><span data-stu-id="c4323-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="c4323-184">Например, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, поскольку свойства и методы `RangeAreas` действуют, как будто новый элемент не был добавлен.</span><span class="sxs-lookup"><span data-stu-id="c4323-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="c4323-185">Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)` вызывает ошибку, если `index` больше, чем `areasCount-1`. </span><span class="sxs-lookup"><span data-stu-id="c4323-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="c4323-186">Аналогичным образом, удаление объекта `Range` в массиве `RangeAreas.areas.items` путем получения ссылки на него и вызова его метода `Range.delete` приводит к ошибкам: хотя объект `Range` *удален*, свойства и методы родительского объекта `RangeAreas` будут действовать (или пытаться действовать), как будто он еще существует.</span><span class="sxs-lookup"><span data-stu-id="c4323-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="c4323-187">Например, если код вызывает метод `RangeAreas.calculate`, Office попытается рассчитать диапазон, но это завершится ошибкой, поскольку объект range отсутствует.</span><span class="sxs-lookup"><span data-stu-id="c4323-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="c4323-188">Задание свойств для нескольких диапазонов</span><span class="sxs-lookup"><span data-stu-id="c4323-188">Set properties on multiple ranges</span></span>

<span data-ttu-id="c4323-189">Установка свойства для объекта `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-189">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="c4323-190">Ниже приведен пример установки свойства в нескольких диапазонах.</span><span class="sxs-lookup"><span data-stu-id="c4323-190">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="c4323-191">Функция выделяет диапазоны **F3:F5** и **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="c4323-191">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="c4323-192">Этот пример применяется к сценариям, в которых можно жестко задать адреса диапазонов, передаваемых в `getRanges`, или легко рассчитать их во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="c4323-192">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="c4323-193">Ниже перечислены некоторые сценарии, в которых это возможно:</span><span class="sxs-lookup"><span data-stu-id="c4323-193">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="c4323-194">Код выполняется в контексте известного шаблона.</span><span class="sxs-lookup"><span data-stu-id="c4323-194">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="c4323-195">Код выполняется в контексте импортированных данных, в котором известна схема данных.</span><span class="sxs-lookup"><span data-stu-id="c4323-195">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="c4323-196">Получение специальных ячеек из нескольких диапазонов</span><span class="sxs-lookup"><span data-stu-id="c4323-196">Get special cells from multiple ranges</span></span>

<span data-ttu-id="c4323-197">Методы `getSpecialCells` и `getSpecialCellsOrNullObject` для объекта `RangeAreas` действуют аналогично методам с теми же названиями для объекта `Range`.</span><span class="sxs-lookup"><span data-stu-id="c4323-197">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="c4323-198">Эти методы возвращают ячейки с указанными характеристиками из всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-198">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="c4323-199">Дополнительные сведения о специальных ячейках см. в разделе [Поиск специальных ячеек в диапазоне](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview).</span><span class="sxs-lookup"><span data-stu-id="c4323-199">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) section for more details on special cells.</span></span>

<span data-ttu-id="c4323-200">При вызове метода `getSpecialCells` или `getSpecialCellsOrNullObject` для объекта `RangeAreas`:</span><span class="sxs-lookup"><span data-stu-id="c4323-200">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="c4323-201">Если в качестве первого параметра передается `Excel.SpecialCellType.sameConditionalFormat`, метод возвращает все ячейки с таким же условным форматированием, как у крайней левой верхней ячейки первого диапазона в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-201">There is one small difference in the behavior of the methods when called on a  object instead of a  object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell in the first range in the  collection.</span></span>
- <span data-ttu-id="c4323-202">Если в качестве первого параметра передается `Excel.SpecialCellType.sameDataValidation`, метод возвращает все ячейки с таким же правилом проверки данных, как у крайней левой верхней ячейки первого диапазона в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-202">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="c4323-203">Чтение свойств RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c4323-203">Read properties of RangeAreas</span></span>

<span data-ttu-id="c4323-204">Чтение значений свойств `RangeAreas` требует внимания, так как определенное свойство может иметь разные значения для разных диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="c4323-204">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="c4323-205">Общее правило заключается в том, что если соответствующее значение *может* быть возвращено, оно будет возвращено.</span><span class="sxs-lookup"><span data-stu-id="c4323-205">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="c4323-206">Например, в приведенном ниже коде RGB-код для розового цвета (`#FFC0CB`) и `true` будут записаны в консоль, так как оба диапазона в объекте `RangeAreas` имеют розовую заливку и оба являются целыми столбцами.</span><span class="sxs-lookup"><span data-stu-id="c4323-206">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
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

<span data-ttu-id="c4323-207">Все усложняется, если согласование невозможно.</span><span class="sxs-lookup"><span data-stu-id="c4323-207">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="c4323-208">Свойства `RangeAreas` действуют в соответствии с приведенными ниже тремя принципами:</span><span class="sxs-lookup"><span data-stu-id="c4323-208">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="c4323-209">Логическое свойство объекта `RangeAreas` возвращает значение `false`, кроме случаев, когда свойство имеет значение true для всех диапазонов элементов.</span><span class="sxs-lookup"><span data-stu-id="c4323-209">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="c4323-210">Свойства, не являющиеся логическими, за исключением свойства `address`, возвращают значение `null`, кроме тех случаев, когда соответствующее свойство для всех диапазонов элементов обладает тем же значением.</span><span class="sxs-lookup"><span data-stu-id="c4323-210">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="c4323-211">Свойство `address` возвращает строку с адресами диапазонов элементов, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="c4323-211">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="c4323-212">Например, в приведенном ниже коде создается объект `RangeAreas`, в котором только один диапазон является целым столбцом и только один залит розовым цветом.</span><span class="sxs-lookup"><span data-stu-id="c4323-212">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="c4323-213">Консоль отобразит значение `null` для цвета заливки, `false` для свойства `isEntireRow` и "Sheet1!F3:F5, Sheet1!H:H" (при условии, что имя листа — "Sheet1") для свойства `address`.</span><span class="sxs-lookup"><span data-stu-id="c4323-213">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
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

## <a name="see-also"></a><span data-ttu-id="c4323-214">См. также</span><span class="sxs-lookup"><span data-stu-id="c4323-214">See also</span></span>

- [<span data-ttu-id="c4323-215">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="c4323-215">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="c4323-216">Работа с диапазонами с использованием API JavaScript для Excel (основные задачи)</span><span class="sxs-lookup"><span data-stu-id="c4323-216">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="c4323-217">Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)</span><span class="sxs-lookup"><span data-stu-id="c4323-217">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
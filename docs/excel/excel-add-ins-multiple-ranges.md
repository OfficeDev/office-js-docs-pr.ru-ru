---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: f1217fc76d14269882a73ec5eb7758e519563456
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383227"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="f7f7e-102">Работа с несколькими диапазонами одновременно в надстройках Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="f7f7e-103">Библиотека JavaScript для Excel позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="f7f7e-104">Диапазоны необязательно должны быть смежными.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="f7f7e-105">Этот способ установки свойства не только упрощает код, но и выполняется намного быстрее, чем установка одинакового свойства отдельно для каждого диапазона.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="f7f7e-106">Для работы с API-интерфейсами, описанными в этой статье, требуется **Office 2016 "нажми и работай" версии 1809 сборки 10820.20000** или более поздней версии</span><span class="sxs-lookup"><span data-stu-id="f7f7e-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="f7f7e-107">(возможно, вам придется принять участие в [программе предварительной оценки Office](https://products.office.com/office-insider) для получения нужной сборки). Кроме того, необходимо загрузить бета-версию библиотеки JavaScript для Office из сети [CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="f7f7e-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="f7f7e-108">В настоящее время нет справочных страниц для этих API.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="f7f7e-109">Но следующий файл типа определения содержит их описания: [бета-версия office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="f7f7e-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="f7f7e-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="f7f7e-110">RangeAreas</span></span>

<span data-ttu-id="f7f7e-111">Набор диапазонов (возможно, несмежных) представлен объектом `Excel.RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="f7f7e-112">Его свойства и методы аналогичны типу `Range` (многие с одинаковыми или похожими именами), но с изменением указанных ниже параметров:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="f7f7e-113">Типы данных для свойств и поведений методов задания и методов получения.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="f7f7e-114">Типы данных параметров метода и поведений метода.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="f7f7e-115">Типы данных возвращаемых значений метода.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-115">The data types of method return values.</span></span>

<span data-ttu-id="f7f7e-116">Примеры:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-116">Some examples:</span></span>

- <span data-ttu-id="f7f7e-117">У `RangeAreas` есть свойство `address`, возвращающее строку с адресами диапазона, разделенными запятой, а не только один адрес, как в случае со свойством `Range.address`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="f7f7e-118">У `RangeAreas` есть свойство `dataValidation`, которое возвращает объект `DataValidation`, представляющий проверку данных всех диапазонов в `RangeAreas` при соответствии.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="f7f7e-119">Значение этого свойства будет равно `null`, если ко всем диапазонам в `RangeAreas` не применяются одинаковые объекты `DataValidation`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="f7f7e-120">Это общий, но не универсальный принцип для объекта `RangeAreas`: *если у свойства нет согласованных значений во всех диапазонах в `RangeAreas`, его значением будет `null`.*</span><span class="sxs-lookup"><span data-stu-id="f7f7e-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="f7f7e-121">Дополнительные сведения и некоторые исключения см. в разделе [Чтение свойств RangeAreas](#read-properties-of-rangeareas).</span><span class="sxs-lookup"><span data-stu-id="f7f7e-121">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="f7f7e-122">`RangeAreas.cellCount` получает общее количество ячеек во всех диапазонах в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="f7f7e-123">`RangeAreas.calculate` пересчитывает ячейки всех диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="f7f7e-124">`RangeAreas.getEntireColumn` и `RangeAreas.getEntireRow` возвращают другой объект `RangeAreas`, представляющий все столбцы (или строки) во всех диапазонах в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="f7f7e-125">Например, если `RangeAreas` представляет "A1:C4" и "F14:L15", то `RangeAreas.getEntireColumn` возвращает объект `RangeAreas`, представляющий "A:C" и "F:L".</span><span class="sxs-lookup"><span data-stu-id="f7f7e-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="f7f7e-126">`RangeAreas.copyFrom` может использовать параметр `Range` или `RangeAreas`, представляющий исходный диапазон (или диапазоны) операции копирования.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="f7f7e-127">Полный список элементов Range, также доступных в RangeAreas</span><span class="sxs-lookup"><span data-stu-id="f7f7e-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="f7f7e-128">Свойства</span><span class="sxs-lookup"><span data-stu-id="f7f7e-128">Properties</span></span>

<span data-ttu-id="f7f7e-129">Ознакомьтесь со статьей [Чтение свойств RangeAreas](#read-properties-of-rangeareas) перед написанием кода, считывающего любое из перечисленных свойств.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-129">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="f7f7e-130">Возвращаемое значение зависит от ряда факторов.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="f7f7e-131">address</span><span class="sxs-lookup"><span data-stu-id="f7f7e-131">address</span></span>
- <span data-ttu-id="f7f7e-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="f7f7e-132">addressLocal</span></span>
- <span data-ttu-id="f7f7e-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="f7f7e-133">cellCount</span></span>
- <span data-ttu-id="f7f7e-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="f7f7e-134">conditionalFormats</span></span>
- <span data-ttu-id="f7f7e-135">context</span><span class="sxs-lookup"><span data-stu-id="f7f7e-135">context</span></span>
- <span data-ttu-id="f7f7e-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="f7f7e-136">dataValidation</span></span>
- <span data-ttu-id="f7f7e-137">format</span><span class="sxs-lookup"><span data-stu-id="f7f7e-137">format</span></span>
- <span data-ttu-id="f7f7e-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="f7f7e-138">isEntireColumn</span></span>
- <span data-ttu-id="f7f7e-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="f7f7e-139">isEntireRow</span></span>
- <span data-ttu-id="f7f7e-140">style</span><span class="sxs-lookup"><span data-stu-id="f7f7e-140">style</span></span>
- <span data-ttu-id="f7f7e-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="f7f7e-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="f7f7e-142">Методы</span><span class="sxs-lookup"><span data-stu-id="f7f7e-142">Methods</span></span>

<span data-ttu-id="f7f7e-143">Методы Range в предварительной версии помечены.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="f7f7e-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-144">calculate()</span></span>
- <span data-ttu-id="f7f7e-145">clear()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-145">clear()</span></span>
- <span data-ttu-id="f7f7e-146">convertDataTypeToText() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="f7f7e-147">convertToLinkedDataType() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="f7f7e-148">copyFrom() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="f7f7e-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-149">getEntireColumn()</span></span>
- <span data-ttu-id="f7f7e-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-150">getEntireRow()</span></span>
- <span data-ttu-id="f7f7e-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-151">getIntersection()</span></span>
- <span data-ttu-id="f7f7e-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="f7f7e-153">getOffsetRange() (называется getOffsetRangeAreas в объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="f7f7e-154">getSpecialCells() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="f7f7e-155">getSpecialCellsOrNullObject() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="f7f7e-156">getTables() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-156">getTables() (preview)</span></span>
- <span data-ttu-id="f7f7e-157">getUsedRange() (называется getUsedRangeAreas в объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="f7f7e-158">getUsedRangeOrNullObject() (называется getUsedRangeAreasOrNullObject в объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="f7f7e-159">load()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-159">load()</span></span>
- <span data-ttu-id="f7f7e-160">set()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-160">set()</span></span>
- <span data-ttu-id="f7f7e-161">setDirty() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-161">setDirty() (preview)</span></span>
- <span data-ttu-id="f7f7e-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-162">toJSON()</span></span>
- <span data-ttu-id="f7f7e-163">track()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-163">track()</span></span>
- <span data-ttu-id="f7f7e-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="f7f7e-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="f7f7e-165">Свойства и методы, характерные для объекта RangeArea</span><span class="sxs-lookup"><span data-stu-id="f7f7e-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="f7f7e-166">Для типа `RangeAreas` существуют несколько свойств и методов, отсутствующих в объекте `Range`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="f7f7e-167">Ниже приведены некоторые из них.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-167">The following is a selection of them:</span></span>

- <span data-ttu-id="f7f7e-168">`areas`. Объект `RangeCollection`, содержащий все диапазоны, которые представлены объектом `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="f7f7e-169">Объект `RangeCollection` — еще один новый объект, аналогичный другим объектам коллекции Excel.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="f7f7e-170">У него есть свойство `items`, являющееся массивом объектов `Range`, которые представляют диапазоны.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="f7f7e-171">`areaCount`. Общее количество диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="f7f7e-172">`getOffsetRangeAreas`. Действует аналогично методу [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), за исключением того, что возвращается объект `RangeAreas`, содержащий диапазоны, каждый из которых смещен относительно одного из диапазонов в исходном объекте `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="f7f7e-173">Создание RangeAreas и установка свойств</span><span class="sxs-lookup"><span data-stu-id="f7f7e-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="f7f7e-174">Объект `RangeAreas` можно создать двумя основными способами:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="f7f7e-175">Вызвать метод `Worksheet.getRanges()` и передать ему строку с адресами диапазона, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="f7f7e-176">Если диапазон, который нужно включить, был преобразован [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), в строку можно включить имя вместо адреса.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="f7f7e-177">Вызвать метод `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="f7f7e-178">Этот метод возвращает объект `RangeAreas`, представляющий все диапазоны, выбранные в активном на данный момент листе.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="f7f7e-179">После получения объекта `RangeAreas` можно создать другие с помощью методов объекта, возвращающих `RangeAreas`, например `getOffsetRangeAreas` и `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="f7f7e-180">Нельзя напрямую добавить дополнительные диапазоны к объекту `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="f7f7e-181">Например, у коллекции в `RangeAreas.areas` нет метода `add`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="f7f7e-182">Не пытайтесь напрямую добавлять или удалять элементы из массива `RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="f7f7e-183">Это приведет к нежелательному поведению кода. </span><span class="sxs-lookup"><span data-stu-id="f7f7e-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="f7f7e-184">Например, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, поскольку свойства и методы `RangeAreas` действуют, как будто новый элемент не был добавлен.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="f7f7e-185">Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)` вызывает ошибку, если `index` больше, чем `areasCount-1`. </span><span class="sxs-lookup"><span data-stu-id="f7f7e-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="f7f7e-186">Аналогичным образом, удаление объекта `Range` в массиве `RangeAreas.areas.items` путем получения ссылки на него и вызова его метода `Range.delete` приводит к ошибкам: хотя объект `Range` *удален*, свойства и методы родительского объекта `RangeAreas` будут действовать (или пытаться действовать), как будто он еще существует.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="f7f7e-187">Например, если код вызывает метод `RangeAreas.calculate`, Office попытается рассчитать диапазон, но это завершится ошибкой, поскольку объект range отсутствует.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="f7f7e-188">Установка свойства для `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="f7f7e-189">Ниже приведен пример установки свойства в нескольких диапазонах.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="f7f7e-190">Функция выделяет диапазоны **F3:F5** и **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="f7f7e-191">Этот пример применяется к сценариям, в которых можно жестко задать адреса диапазонов, передаваемых в `getRanges`, или легко рассчитать их во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="f7f7e-192">Ниже перечислены некоторые сценарии, в которых это возможно:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="f7f7e-193">Код выполняется в контексте известного шаблона.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="f7f7e-194">Код выполняется в контексте импортированных данных, в котором известна схема данных.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="f7f7e-195">Если при создании кода неизвестно, с какими диапазонами придется работать, необходимо обнаружить их во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="f7f7e-196">В следующем разделе рассматриваются эти сценарии.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="f7f7e-197">Обнаружение областей диапазона программным способом</span><span class="sxs-lookup"><span data-stu-id="f7f7e-197">Discover range areas programmatically</span></span>

<span data-ttu-id="f7f7e-198">Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` позволяют во время выполнения обнаруживать диапазоны, с которыми нужно работать, на основе характеристик ячеек и типа значений в ячейках.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="f7f7e-199">Подписи методов из файла типов данных TypeScript:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="f7f7e-200">Ниже приведен пример использования первого из них.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-200">The following is an example of using the first one.</span></span> <span data-ttu-id="f7f7e-201">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-201">About this code, note:</span></span>

- <span data-ttu-id="f7f7e-202">Он ограничивает часть листа, в которой требуется выполнять поиск, путем вызова сначала метода `Worksheet.getUsedRange`, а затем метода `getSpecialCells` только для этого диапазона.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="f7f7e-203">В качестве параметра для `getSpecialCells` он передает строковое представление значения из перечисления `Excel.SpecialCellType`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="f7f7e-204">Некоторые другие значения, которые могут быть переданы вместо этого: "Blanks" для пустых ячеек, "Constants" для ячеек со значениями литералов вместо формул и "SameConditionalFormat" для ячеек, у которых такое же условное форматирование, как и у первой ячейки в `usedRange`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="f7f7e-205">Первая ячейка — это верхняя крайняя ячейка слева.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="f7f7e-206">Полный список значений перечисления см. в [бета-версии office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="f7f7e-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="f7f7e-207">Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами окрашены розовым цветом даже в том случае, если они не являются смежными.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="f7f7e-208">В некоторых случаях диапазон не содержит *ни одной* ячейки с целевой характеристикой.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="f7f7e-209">Если метод `getSpecialCells` не находит ни одной такой ячейки, он выдает ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="f7f7e-210">Это приведет к переадресации потока управления к блоку/методу `catch`, если таковой существует.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="f7f7e-211">Если нет, ошибка останавливает исполнение функции.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="f7f7e-212">Могут существовать сценарии, в которых выдача ошибки – это именно то, что должно происходить при отсутствии ячейки с целевой характеристикой.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="f7f7e-213">Если в сценариях отсутствие соответствующих ячеек является нормальной, но редкой ситуацией, ваш код должен проверить наличие такой возможности и корректно выполнить действие без выдачи ошибки.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="f7f7e-214">Для этих сценариев следует использовать метод `getSpecialCellsOrNullObject` и протестировать свойство `RangeAreas.isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="f7f7e-215">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-215">The following is an example.</span></span> <span data-ttu-id="f7f7e-216">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-216">Note about this code:</span></span>

- <span data-ttu-id="f7f7e-217">Метод `getSpecialCellsOrNullObject` всегда возвращает прокси-объект, поэтому он не может иметь значение `null` в обычном смысле JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="f7f7e-218">Но если соответствующие ячейки не обнаружены, свойству `isNullObject` объекта присваивается значение `true`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="f7f7e-219">Он вызывает `context.sync` *перед* тестированием свойства `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="f7f7e-220">Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="f7f7e-221">Однако необязательно *явно* загружать свойство `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="f7f7e-222">Оно автоматически загружается с помощью `context.sync`, даже если `load` не вызывается для объекта.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="f7f7e-223">Дополнительные сведения см. в разделе [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="f7f7e-223">For more information, see [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="f7f7e-224">Этот код можно проверить, выбрав сначала диапазон без ячеек с формулами и запустив его.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="f7f7e-225">Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
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

<span data-ttu-id="f7f7e-226">Для удобства во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="f7f7e-227">Ограничение целевых ячеек с помощью типа значений ячеек</span><span class="sxs-lookup"><span data-stu-id="f7f7e-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="f7f7e-228">Существует необязательный второй параметр типа перечисления `Excel.SpecialCellValueType`, который дополнительно ограничивает целевые ячейки.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="f7f7e-229">Его можно использовать только в том случае, если передается значение "Formulas" или "Constants" для `getSpecialCells` или `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="f7f7e-230">Этот параметр указывает, что требуются только ячейки с определенными типами значений.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="f7f7e-231">Существует четыре основных типа: "Error", "Logical" (логический), "Numbers" и "Text"</span><span class="sxs-lookup"><span data-stu-id="f7f7e-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="f7f7e-232">(у перечисления есть другие значения помимо этих четырех, которые рассматриваются ниже). Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="f7f7e-233">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-233">About this code, note:</span></span>

- <span data-ttu-id="f7f7e-234">Он выделяет только ячейки с числовым значением литерала.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="f7f7e-235">Он не выделяет ячейки с формулой (даже если результат является числом), логическим значением, текстовым значением или ячейки с состоянием ошибки.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="f7f7e-236">Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="f7f7e-237">Иногда требуется работать с ячейками, имеющими несколько типов значений, например со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями ("Logical").</span><span class="sxs-lookup"><span data-stu-id="f7f7e-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="f7f7e-238">Перечисление `Excel.SpecialCellValueType` содержит значения, позволяющие объединять типы.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="f7f7e-239">Например, "LogicalText" обрабатывает все ячейки с логическими и текстовыми значениями.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="f7f7e-240">Можно объединить любые два или три из четырех основных типов.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="f7f7e-241">Имена этих значений перечисления, объединяющих основные типы, всегда располагаются в алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="f7f7e-242">Поэтому для объединения ячеек со значениями ошибок, текстовыми и логическими значениями используйте параметр "ErrorLogicalText", а не "LogicalErrorText" или "TextErrorLogical".</span><span class="sxs-lookup"><span data-stu-id="f7f7e-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="f7f7e-243">Параметр по умолчанию "All" объединяет все четыре типа.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="f7f7e-244">В приведенном ниже примере выделены все ячейки с формулами, которые производят числовые или логические значения:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> <span data-ttu-id="f7f7e-245">Параметр `Excel.SpecialCellValueType` можно использовать, только если параметру `Excel.SpecialCellType` присвоено значение "Formulas" или "Constants".</span><span class="sxs-lookup"><span data-stu-id="f7f7e-245">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` parameter is "Formulas" or "Constants".</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="f7f7e-246">Получение объектов RangeAreas в RangeAreas</span><span class="sxs-lookup"><span data-stu-id="f7f7e-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="f7f7e-247">У типа `RangeAreas` также есть методы `getSpecialCells` и `getSpecialCellsOrNullObject`, которые используют те же два параметра.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="f7f7e-248">Эти методы возвращают все целевые ячейки из всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="f7f7e-249">Существует одно небольшое отличие в поведении методов при вызове объекта `RangeAreas` вместо объекта `Range`: если вы передаете "SameConditionalFormat" в качестве первого параметра, метод возвращает все ячейки с таким же условным форматированием, как у крайней левой верхней ячейки *первого диапазона в коллекции `RangeAreas.areas`*.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="f7f7e-250">То же касается и "SameDataValidation": при передаче к `Range.getSpecialCells` возвращаются все ячейки с таким же правилом проверки данных, как у крайней левой верхней ячейки *диапазона*. </span><span class="sxs-lookup"><span data-stu-id="f7f7e-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="f7f7e-251">Но при передаче к `RangeAreas.getSpecialCells` возвращаются все ячейки с таким же правилом проверки данных, как у крайней левой верхней ячейки *первого диапазона в коллекции `RangeAreas.areas`*.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="f7f7e-252">Чтение свойств RangeAreas</span><span class="sxs-lookup"><span data-stu-id="f7f7e-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="f7f7e-253">Чтение значений свойств `RangeAreas` требует внимания, так как определенное свойство может иметь разные значения для разных диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="f7f7e-254">Общее правило заключается в том, что если соответствующее значение *может* быть возвращено, оно будет возвращено.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="f7f7e-255">Например, в приведенном ниже коде RGB-код для розового цвета (`#FFC0CB`) и `true` будут записаны в консоль, так как оба диапазона в объекте `RangeAreas` имеют розовую заливку и оба являются целыми столбцами.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="f7f7e-256">Все усложняется, если согласование невозможно.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="f7f7e-257">Свойства `RangeAreas` действуют в соответствии с приведенными ниже тремя принципами:</span><span class="sxs-lookup"><span data-stu-id="f7f7e-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="f7f7e-258">Логическое свойство объекта `RangeAreas` возвращает значение `false`, кроме случаев, когда свойство имеет значение true для всех диапазонов элементов.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="f7f7e-259">Свойства, не являющиеся логическими, за исключением свойства `address`, возвращают значение `null`, кроме тех случаев, когда соответствующее свойство для всех диапазонов элементов обладает тем же значением.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="f7f7e-260">Свойство `address` возвращает строку с адресами диапазонов элементов, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="f7f7e-261">Например, в приведенном ниже коде создается объект `RangeAreas`, в котором только один диапазон является целым столбцом и только один залит розовым цветом.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="f7f7e-262">Консоль отобразит значение `null` для цвета заливки, `false` для свойства `isEntireRow` и "Sheet1!F3:F5, Sheet1!H:H" (при условии, что имя листа — "Sheet1") для свойства `address`.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="f7f7e-263">См. также</span><span class="sxs-lookup"><span data-stu-id="f7f7e-263">See also</span></span>

- [<span data-ttu-id="f7f7e-264">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="f7f7e-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="f7f7e-265">Объект Range (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="f7f7e-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="f7f7e-266">[Объект RangeAreas (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (эта ссылка может не работать, пока API находится в предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="f7f7e-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="f7f7e-267">В качестве альтернативы см. [бета-версию office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)).</span><span class="sxs-lookup"><span data-stu-id="f7f7e-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>
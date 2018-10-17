---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: a00bbf15b53649147fb2c2b1dfa590f15c5739be
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506296"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="a3744-102">Работа с несколькими диапазонами одновременно в надстройках Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a3744-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="a3744-p101">Библиотека JavaScript для Excel позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов. Диапазоны необязательно должны быть непрерывными. Этот способ установки свойства не только упрощает код, но и выполняется намного быстрее, чем установка каждого отдельного свойства для каждого диапазона.</span><span class="sxs-lookup"><span data-stu-id="a3744-p101">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously. The ranges do not have to be contiguous. In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="a3744-p102">Для работы с API-интерфейсами, описанными в этой статье, требуется **версия Office 2016 Click-to-Run 1809 сборки 10820.20000** или более поздняя версия (возможно, вам придется принять участие в [программе предварительной оценки Office](https://products.office.com/office-insider) для получения нужной сборки). Кроме того, необходимо загрузить бета-версию библиотеки Office JavaScript из [сети CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Наконец, в настоящее время у нас еще нет страниц ссылки для этих API. Но следующий файл типа определения содержит их описания: [бета-версию office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="a3744-p102">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later. (You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Finally, we don't have reference pages for these APIs yet. But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="a3744-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="a3744-110">RangeAreas</span></span>

<span data-ttu-id="a3744-p103">Набор диапазонов (возможно, разобщенных) представлен `Excel.RangeAreas`объектом . Он имеет свойства и методы, аналогичные `Range`типу  (многие из которых имеют одинаковые или похожие имена), но следующие параметры были изменены:</span><span class="sxs-lookup"><span data-stu-id="a3744-p103">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object. It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="a3744-113">Типы данных для свойств и поведений методов задания и методов получения.</span><span class="sxs-lookup"><span data-stu-id="a3744-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="a3744-114">Типы данных параметров метода и поведений метода.</span><span class="sxs-lookup"><span data-stu-id="a3744-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="a3744-115">Типы данных возвращаемых значений метода.</span><span class="sxs-lookup"><span data-stu-id="a3744-115">The data types of method return values.</span></span>

<span data-ttu-id="a3744-116">Некоторые примеры:</span><span class="sxs-lookup"><span data-stu-id="a3744-116">Some examples:</span></span>

- <span data-ttu-id="a3744-117">`RangeAreas` имеет свойство `address`, которое возвращает строку с адресами диапазона, разделенными запятой, а не только один адрес, как в случае со свойством `Range.address`.</span><span class="sxs-lookup"><span data-stu-id="a3744-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="a3744-p104">`RangeAreas` имеет < свойство `dataValidation`, которое возвращает `DataValidation`объект  , представляющий собой проверку данных всех диапазонов в `RangeAreas` при соответствии. Этим свойством будет `null`, если идентичные `DataValidation`объекты  не применяются ко всем диапазонам в `RangeAreas`. Это общие, но не универсальные принципы для`RangeAreas` объекта : *если свойство не имеет согласованных значений для каждого из всех диапазонов в `RangeAreas`, то это свойство будет `null`.*  Дополнительные сведения и некоторые исключения см. в статье [Чтение свойств RangeAreas](#reading-properties-of-rangeareas).</span><span class="sxs-lookup"><span data-stu-id="a3744-p104">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent. The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`. This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.* See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="a3744-122">`RangeAreas.cellCount` возвращает общее число ячеек во всех диапазонах в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a3744-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="a3744-123">`RangeAreas.calculate` пересчитывает ячейки всех диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a3744-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="a3744-p105">`RangeAreas.getEntireColumn` и `RangeAreas.getEntireRow` возвращает другой объект `RangeAreas`, представляющий все столбцы (или строки) во всех диапазонах в `RangeAreas`. Например, если `RangeAreas` представляет «A1:C4» и «F14:L15», то `RangeAreas.getEntireColumn` возвращает объект `RangeAreas`, представляющий «A:C» и «F:L».</span><span class="sxs-lookup"><span data-stu-id="a3744-p105">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`. For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="a3744-126">`RangeAreas.copyFrom` может использовать параметр `Range` или `RangeAreas`, представляющий диапазон(ы) источника операции копирования.</span><span class="sxs-lookup"><span data-stu-id="a3744-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="a3744-127">Полный список элементов диапазона Range, которые также доступны на RangeAreas</span><span class="sxs-lookup"><span data-stu-id="a3744-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="a3744-128">Свойства</span><span class="sxs-lookup"><span data-stu-id="a3744-128">Properties</span></span>

<span data-ttu-id="a3744-p106">Ознакомьтесь со статьей [Чтение свойств RangeAreas](#reading-properties-of-rangeareas) до написания кода, который считывает все свойства из списка. Возвращаемое значение зависит от ряда факторов.</span><span class="sxs-lookup"><span data-stu-id="a3744-p106">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed. There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="a3744-131">address</span><span class="sxs-lookup"><span data-stu-id="a3744-131">address</span></span>
- <span data-ttu-id="a3744-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="a3744-132">addressLocal</span></span>
- <span data-ttu-id="a3744-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="a3744-133">cellCount</span></span>
- <span data-ttu-id="a3744-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="a3744-134">conditionalFormats</span></span>
- <span data-ttu-id="a3744-135">context</span><span class="sxs-lookup"><span data-stu-id="a3744-135">context</span></span>
- <span data-ttu-id="a3744-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="a3744-136">dataValidation</span></span>
- <span data-ttu-id="a3744-137">format</span><span class="sxs-lookup"><span data-stu-id="a3744-137">format</span></span>
- <span data-ttu-id="a3744-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="a3744-138">isEntireColumn</span></span>
- <span data-ttu-id="a3744-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="a3744-139">isEntireRow</span></span>
- <span data-ttu-id="a3744-140">style</span><span class="sxs-lookup"><span data-stu-id="a3744-140">style</span></span>
- <span data-ttu-id="a3744-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="a3744-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="a3744-142">Методы</span><span class="sxs-lookup"><span data-stu-id="a3744-142">Methods</span></span>

<span data-ttu-id="a3744-143">Помеченные методы диапазона в режиме предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="a3744-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="a3744-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="a3744-144">calculate()</span></span>
- <span data-ttu-id="a3744-145">clear()</span><span class="sxs-lookup"><span data-stu-id="a3744-145">clear()</span></span>
- <span data-ttu-id="a3744-146">convertDataTypeToText() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a3744-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="a3744-147">convertToLinkedDataType() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a3744-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="a3744-148">copyFrom() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a3744-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="a3744-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="a3744-149">getEntireColumn()</span></span>
- <span data-ttu-id="a3744-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="a3744-150">getEntireRow()</span></span>
- <span data-ttu-id="a3744-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="a3744-151">getIntersection()</span></span>
- <span data-ttu-id="a3744-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="a3744-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="a3744-153">getOffsetRange() (с именем getOffsetRangeAreas на объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="a3744-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="a3744-154">getSpecialCells() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a3744-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="a3744-155">getSpecialCellsOrNullObject() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a3744-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="a3744-156">getTables() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a3744-156">getTables() (preview)</span></span>
- <span data-ttu-id="a3744-157">getUsedRange() (с именем getUsedRangeAreas на объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="a3744-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="a3744-158">getUsedRangeOrNullObject() (с именем getUsedRangeAreasOrNullObject на объекте RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="a3744-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="a3744-159">load()</span><span class="sxs-lookup"><span data-stu-id="a3744-159">load()</span></span>
- <span data-ttu-id="a3744-160">set()</span><span class="sxs-lookup"><span data-stu-id="a3744-160">set\*</span></span>
- <span data-ttu-id="a3744-161">setDirty() (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a3744-161">setDirty() (preview)</span></span>
- <span data-ttu-id="a3744-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="a3744-162">toJSON()</span></span>
- <span data-ttu-id="a3744-163">track()</span><span class="sxs-lookup"><span data-stu-id="a3744-163">track</span></span>
- <span data-ttu-id="a3744-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="a3744-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="a3744-165">Свойства и методы, характерные для объекта RangeArea</span><span class="sxs-lookup"><span data-stu-id="a3744-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="a3744-p107">Тип `RangeAreas` имеет некоторые свойства и методы, которые не входят в объект `Range`. Ниже приведены некоторые из них:</span><span class="sxs-lookup"><span data-stu-id="a3744-p107">The `RangeAreas` type has some properties and methods that are not on the `Range` object. The following is a selection of them:</span></span>

- <span data-ttu-id="a3744-p108">`areas`: объект `RangeCollection`, содержащий все диапазоны, представленные объектом `RangeAreas`. Объект `RangeCollection`  — также новый и аналогичен другим объектам коллекции Excel. Он имеет свойство `items`, которое представляет собой массив объектов `Range`, представляющих диапазоны.</span><span class="sxs-lookup"><span data-stu-id="a3744-p108">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object. The `RangeCollection` object is also new and is similar to other Excel collection objects. It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="a3744-171">`areaCount`: общее число диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a3744-171">The total number of recipients in the message.</span></span>
- <span data-ttu-id="a3744-172">`getOffsetRangeAreas`: работает так же, как [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), за исключением того, что `RangeAreas` возвращается и содержит диапазоны, каждый из которых смещен от одного из диапазонов в исходном `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a3744-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="a3744-173">Создание RangeAreas и установка свойств</span><span class="sxs-lookup"><span data-stu-id="a3744-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="a3744-174">Вы можете создать объект `RangeAreas` двумя основными способами:</span><span class="sxs-lookup"><span data-stu-id="a3744-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="a3744-p109">Вызвать `Worksheet.getRanges()`  и передать его в строку с адресами диапазона, разделенными запятыми. Если диапазон, который вы хотите включить, был переделан в [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), вы можете включить в строку имя вместо адреса.</span><span class="sxs-lookup"><span data-stu-id="a3744-p109">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses. If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="a3744-p110">Вызвать `Workbook.getSelectedRanges()`. Этот метод возвращает `RangeAreas`, представляющий все диапазоны, выбранные в активном на данный момент листе.</span><span class="sxs-lookup"><span data-stu-id="a3744-p110">Call `Workbook.getSelectedRanges()`. This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="a3744-179">После получения объекта `RangeAreas` можно создать другие с помощью методов, применяемых к объекту, который возвращает `RangeAreas`, такие как `getOffsetRangeAreas` и `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="a3744-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="a3744-p111">Невозможно напрямую добавить дополнительные диапазоны к объекту `RangeAreas`. Например, коллекция в `RangeAreas.areas` не имеет метода `add`.</span><span class="sxs-lookup"><span data-stu-id="a3744-p111">You cannot directly add additional ranges to a `RangeAreas` object. For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="a3744-p112">Не пытайтесь напрямую добавлять или удалять элементы из массива `RangeAreas.areas.items`. Это приведет к нежелательному функционированию кода. Например, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, так как свойства `RangeAreas` и методы функционируют так, как если бы новый элемент не был добавлен. Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)`  вызывает ошибку, если `index` больше, чем `areasCount-1`. Аналогичным образом, удаление объекта `Range`  в массиве `RangeAreas.areas.items` путем получения ссылки на него и вызов его метода `Range.delete`  вызывает ошибки: хотя объект `Range` \* будет\* удален, свойства и методы родительского объекта `RangeAreas`  будут функционировать (или пытаться функционировать) так, как если бы он все еще присутствовал. Например, если код вызывает метод `RangeAreas.calculate`, Office попытается рассчитать диапазон, но это завершится ошибкой, поскольку объект range отсутствует.</span><span class="sxs-lookup"><span data-stu-id="a3744-p112">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array. This will lead to undesirable behavior in your code. For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there. For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`. Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence. For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="a3744-188">Установка свойства для `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="a3744-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="a3744-p113">Ниже приведен пример установки свойства в нескольких диапазонах. Функция выделяет диапазоны **F3:F5** и **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="a3744-p113">The following is an example of setting a property on multiple ranges. The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="a3744-p114">Этот пример применяется к сценариям, в которых можно создать серьезный код адресов диапазона, передаваемых в `getRanges`, или легко рассчитать их во время выполнения. Ниже перечислены некоторые сценарии, в которых это возможно:</span><span class="sxs-lookup"><span data-stu-id="a3744-p114">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime. Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="a3744-193">Код выполняется в контексте известного шаблона.</span><span class="sxs-lookup"><span data-stu-id="a3744-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="a3744-194">Код выполняется в контексте импортированных данных, в котором известна схема данных.</span><span class="sxs-lookup"><span data-stu-id="a3744-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="a3744-p115">Когда во время создания кода не известно, с какими диапазонами вам придется работать, необходимо обнаружить их во время выполнения. В следующем разделе описываются эти сценарии.</span><span class="sxs-lookup"><span data-stu-id="a3744-p115">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime. The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="a3744-197">Обнаружение областей диапазона с помощью программных средств</span><span class="sxs-lookup"><span data-stu-id="a3744-197">Discover range areas programmatically</span></span>

<span data-ttu-id="a3744-p116">Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` можно использовать для поиска во время выполнения диапазонов, с которыми вы хотите работать, на основе характеристик ячеек и типа значений в ячейках. Вот подписи методов из файла типов данных TypeScript:</span><span class="sxs-lookup"><span data-stu-id="a3744-p116">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells. Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="a3744-p117">Ниже приведен пример использования первого из них. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a3744-p117">The following is an example of using the first one. About this code, note:</span></span>

- <span data-ttu-id="a3744-202">Он ограничивает часть листа, которую нужно искать, вызвав сначала `Worksheet.getUsedRange`, а затем вызвав `getSpecialCells` только для этого диапазона.</span><span class="sxs-lookup"><span data-stu-id="a3744-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="a3744-p118">Передает в качестве параметра `getSpecialCells` строчную версию значения из перечисления `Excel.SpecialCellType`. Некоторые другие значения, которые могут быть переданы вместо этого, — это "Blanks" для пустых ячеек, "Constants" для ячейки со значениями литералов вместо формул и "SameConditionalFormat" для ячеек, у которых такое же условное форматирование, как и у первой ячейки в `usedRange`. Первая ячейка — это самая верхняя ячейка слева. Полный список значений перечисления см. в [бета-версии office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="a3744-p118">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum. Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`. The first cell is the upper leftmost cell. For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="a3744-207">Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами залиты розовым цветом даже в том случае, если они не последовательны.</span><span class="sxs-lookup"><span data-stu-id="a3744-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="a3744-p119">В некоторых случаях диапазон не содержит *ни одной* ячейки с целевой характеристикой. Если `getSpecialCells` не находит ни одной такой ячеки, он выдает ошибку **ItemNotFound** . Это приведет к переадресации потока управления к блоку/методу `catch`, если таковой существует. Если нет, ошибка приведет к прекращению исполнения функции. Могут существовать сценарии, в которых выдача ошибки – это именно то, что должно происходить при отсутствии ячейки с целевой характеристикой.</span><span class="sxs-lookup"><span data-stu-id="a3744-p119">Sometimes the range doesn't have *any* cells with the targeted characteristic. If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error. This would divert the flow of control to a `catch` block/method, if there is one. If there isn't, the error halts the function. There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="a3744-p120">Но в некоторых сценариях отсутствие соответствующих ячеек нормально, хотя и, возможно, необычно; ваш код должен проверить наличие такой возможности и аккуратно провести работу со сценарием без выдачи ошибки. Для этих сценариев следует использовать метод `getSpecialCellsOrNullObject`  и протестировать свойство `RangeAreas.isNullObject`. Пример см. ниже. Что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a3744-p120">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error. For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property. The following is an example. Note about this code:</span></span>

- <span data-ttu-id="a3744-p121">`getSpecialCellsOrNullObject`Метод  всегда возвращает объект прокси-сервера, поэтому он не может быть `null` в обычном смысле JavaScript. Но если соответствующие ячейки не обнаружены, `isNullObject`свойству  объекта присваивается значение `true`.</span><span class="sxs-lookup"><span data-stu-id="a3744-p121">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense. But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="a3744-p122">Он вызывает `context.sync` *перед* тестированием `isNullObject`свойства . Это требование для всех `*OrNullObject`методов и свойств, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть. Тем не менее, необязательно *прямо* загружать`isNullObject` свойство. Оно автоматически загружается `context.sync` даже в том случае, если `load` не вызывается для объекта. Дополнительные сведения см. в разделе [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="a3744-p122">It calls `context.sync` *before* it tests the `isNullObject` property. This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it. However, it is not necessary to *explicitly* load the `isNullObject` property. It is automatically loaded by the `context.sync` even if `load` is not called on the object. For more information, see [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="a3744-p123">Этот код можно проверить, выбрав сначала диапазон, у которого нет ячеек формулы, и запустив его. Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.</span><span class="sxs-lookup"><span data-stu-id="a3744-p123">You can test this code by first selecting a range that has no formula cells and running it. Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="a3744-226">Для простоты во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a3744-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="a3744-227">Сужение целевых ячеек с типом значений ячеек</span><span class="sxs-lookup"><span data-stu-id="a3744-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="a3744-p124">Также существует необязательный второй параметр типа перечисления `Excel.SpecialCellValueType`, который в дальнейшем сужает ячейки до целевого объекта.  Его можно использовать только в том случае, если передается значение "Formulas" или "Constants" для `getSpecialCells` или `getSpecialCellsOrNullObject`. Этот параметр указывает, что требуются только ячейки с определенными типами значений. Существует четыре основных типа: "Error", "Logical" (то же самое, что и boolean — логический), "Numbers" и "Text" (перечисление имеет другие значения помимо этих четырех, которые рассматриваются ниже). См. пример ниже. Что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a3744-p124">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target. You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`. The parameter specifies that you only want cells with certain types of values. There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text". (The enum has other values besides these four which are discussed below.) The following is an example. About this code, note:</span></span>

- <span data-ttu-id="a3744-p125">Он выделяет только ячейки, имеющие числовое значение литерала и не выделяет ячейки, в которых содержится формула (даже в том случае, если результат является числом), логическое, текстовое значение, или ячейки с состоянием ошибки.</span><span class="sxs-lookup"><span data-stu-id="a3744-p125">It will only highlight cells that have a literal number value. It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="a3744-236">Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.</span><span class="sxs-lookup"><span data-stu-id="a3744-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="a3744-p126">В некоторых случаях вам нужно работать с ячейками, имеющими более одного типа значения, например, со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями ("Logical"). Перечисление `Excel.SpecialCellValueType` содержит значения, которые позволяют объединять различные типы. Например, "LogicalText" обрабатывает все логические и все текстовые ячейки. Вы можете использовать любые два или три из четырех основных типов. Имена этих значений перечисления, которые объединяют основные типы, всегда располагаются в алфавитном порядке. Поэтому для объединения ячеек со значениями ошибок, текстовыми и логическими значениями используйте "ErrorLogicalText", а не "LogicalErrorText" или "TextErrorLogical". Параметр по умолчанию "All" объединяет все четыре типа. В следующем примере выделены все ячейки с формулами, которые производят числовые или логические значения:</span><span class="sxs-lookup"><span data-stu-id="a3744-p126">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells. The `Excel.SpecialCellValueType` enum has values that let you combine types. For example, "LogicalText" will target all boolean and all text-valued cells. You can combine any two or any three of the four basic types. The names of these enum values that combine basic types are always in alphabetical order. So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical". The default parameter of "All" combines all four types. The following example highlights all cells with formulas that produce number or boolean values:</span></span>

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
> <span data-ttu-id="a3744-245">Параметр `Excel.SpecialCellValueType` можно использовать, только если параметр `Excel.SpecialCellType` — это "Formulas" или "Constants".</span><span class="sxs-lookup"><span data-stu-id="a3744-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="a3744-246">Получение объектов RangeAreas в RangeAreas</span><span class="sxs-lookup"><span data-stu-id="a3744-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="a3744-p127">Тип  `RangeAreas` также имеет методы  `getSpecialCells` и `getSpecialCellsOrNullObject`, которые принимают те же два параметра. Эти методы возвращают все целевые ячейки из всех диапазонов в коллекции `RangeAreas.areas`. Существует одно небольшое отличие в поведении методов при вызове объекта `RangeAreas`вместо объекта `Range`: когда вы передаете "SameConditionalFormat" в качестве первого параметра, метод возвращает все ячейки, имеющие одинаковое условное форматирование, как верхнюю крайнюю слева ячейку *в первом диапазоне в коллекции `getSpecialCellsOrNullObject`*. То же касается и "SameDataValidation": при передаче к `Range.getSpecialCells`он возвращает все ячейки, которые имеют такое же правило проверки данных, как верхнюю крайнюю слева ячейку *в диапазоне*. Но при передаче к `RangeAreas.getSpecialCells` он возвращает все ячейки, которые имеют такое же правило проверки данных, как верхнюю крайнюю слева ячейку \* в первом диапазоне в коллекции`RangeAreas.areas`\*.</span><span class="sxs-lookup"><span data-stu-id="a3744-p127">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters. These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection. There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*. The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*. But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="a3744-252">Чтение свойств RangeAreas</span><span class="sxs-lookup"><span data-stu-id="a3744-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="a3744-p128">Чтение значения свойств `RangeAreas` требует внимания, так как то или иное свойство может иметь разные значения для разных диапазонов в `RangeAreas`. Общее правило заключается в том, что если соответствующее значение *может*  быть возвращено, оно будет возвращено. Например, в следующем коде RGB-код для розовой заливки (`#FFC0CB`) и `true`  будет выполнять вход в консоль, так как оба диапазона в объекте `RangeAreas`  имеют розовую заливку и оба являются целыми столбцами.</span><span class="sxs-lookup"><span data-stu-id="a3744-p128">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`. The general rule is that if a consistent value *can* be returned it will be returned. For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="a3744-p129">Все усложняется, когда согласование невозможно. Свойство `RangeAreas` работает в соответствии со следующими тремя принципами:</span><span class="sxs-lookup"><span data-stu-id="a3744-p129">Things get more complicated when consistency isn't possible. The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="a3744-258">Логическое свойство объекта `RangeAreas` возвращает `false`, кроме случаев, когда свойство имеет значение true для всех диапазонов элементов.</span><span class="sxs-lookup"><span data-stu-id="a3744-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="a3744-259">Свойства, не являющиеся логическими, за исключением свойства `address`, возвращают `null`, кроме тех случаев, когда соответствующее свойство для всех элементов диапазона обладает тем же значением.</span><span class="sxs-lookup"><span data-stu-id="a3744-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="a3744-260">Свойство `address` возвращает строку с адресами диапазонов элементов, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="a3744-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="a3744-p130">Например, следующий код создает `RangeAreas`, в котором только один диапазон — это целый столбец, и только один заполнен розовым цветом. Консоль покажет `null`для цвета заливки, `false`для `isEntireRow`свойства  и "Sheet1! F3:F5 Sheet1! H:H" (при условии, что имя листа – это "Sheet1") для `address`свойства .</span><span class="sxs-lookup"><span data-stu-id="a3744-p130">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink. The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="a3744-263">См. также</span><span class="sxs-lookup"><span data-stu-id="a3744-263">See also</span></span>

- [<span data-ttu-id="a3744-264">Основные принципы программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a3744-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="a3744-265">Объект Range (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="a3744-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="a3744-p131">Объект[RangeAreas (JavaScript API для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (эта ссылка может не работать, пока API находится в режиме предварительной версии. В качестве альтернативы см. [бета-версию office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span><span class="sxs-lookup"><span data-stu-id="a3744-p131">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview. As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>
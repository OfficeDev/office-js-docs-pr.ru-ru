---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: Узнайте, как библиотека JavaScript для Excel позволяет надстройке выполнять операции, а также задавать свойства для нескольких диапазонов одновременно.
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: 97481b4b8ab76f7bbc5bd10378d4cc6512bc7b6a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717070"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a><span data-ttu-id="7d770-103">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="7d770-103">Work with multiple ranges simultaneously in Excel add-ins</span></span>

<span data-ttu-id="7d770-104">Библиотека JavaScript для Excel позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов.</span><span class="sxs-lookup"><span data-stu-id="7d770-104">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="7d770-105">Диапазоны необязательно должны быть смежными.</span><span class="sxs-lookup"><span data-stu-id="7d770-105">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="7d770-106">Этот способ установки свойства не только упрощает код, но и выполняется намного быстрее, чем установка одинакового свойства отдельно для каждого диапазона.</span><span class="sxs-lookup"><span data-stu-id="7d770-106">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

## <a name="rangeareas"></a><span data-ttu-id="7d770-107">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="7d770-107">RangeAreas</span></span>

<span data-ttu-id="7d770-108">Набор диапазонов (возможно, несмежных) представлен объектом [RangeAreas](/javascript/api/excel/excel.rangeareas) .</span><span class="sxs-lookup"><span data-stu-id="7d770-108">A set of (possibly discontiguous) ranges is represented by a [RangeAreas](/javascript/api/excel/excel.rangeareas) object.</span></span> <span data-ttu-id="7d770-109">Его свойства и методы аналогичны типу `Range` (многие с одинаковыми или похожими именами), но с изменением указанных ниже параметров:</span><span class="sxs-lookup"><span data-stu-id="7d770-109">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="7d770-110">Типы данных для свойств и поведений методов задания и методов получения.</span><span class="sxs-lookup"><span data-stu-id="7d770-110">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="7d770-111">Типы данных параметров метода и поведений метода.</span><span class="sxs-lookup"><span data-stu-id="7d770-111">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="7d770-112">Типы данных возвращаемых значений метода.</span><span class="sxs-lookup"><span data-stu-id="7d770-112">The data types of method return values.</span></span>

<span data-ttu-id="7d770-113">Примеры:</span><span class="sxs-lookup"><span data-stu-id="7d770-113">Some examples:</span></span>

- <span data-ttu-id="7d770-114">У `RangeAreas` есть свойство `address`, возвращающее строку с адресами диапазона, разделенными запятой, а не только один адрес, как в случае со свойством `Range.address`.</span><span class="sxs-lookup"><span data-stu-id="7d770-114">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="7d770-115">У `RangeAreas` есть свойство `dataValidation`, которое возвращает объект `DataValidation`, представляющий проверку данных всех диапазонов в `RangeAreas` при соответствии.</span><span class="sxs-lookup"><span data-stu-id="7d770-115">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="7d770-116">Значение этого свойства будет равно `null`, если ко всем диапазонам в `RangeAreas` не применяются одинаковые объекты `DataValidation`.</span><span class="sxs-lookup"><span data-stu-id="7d770-116">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="7d770-117">Это общий, но не универсальный принцип для объекта `RangeAreas`: *если у свойства нет согласованных значений во всех диапазонах в `RangeAreas`, его значением будет `null`.*</span><span class="sxs-lookup"><span data-stu-id="7d770-117">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="7d770-118">Дополнительные сведения и некоторые исключения см. в разделе [Чтение свойств RangeAreas](#read-properties-of-rangeareas).</span><span class="sxs-lookup"><span data-stu-id="7d770-118">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="7d770-119">`RangeAreas.cellCount` получает общее количество ячеек во всех диапазонах в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-119">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="7d770-120">`RangeAreas.calculate` пересчитывает ячейки всех диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-120">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="7d770-121">`RangeAreas.getEntireColumn` и `RangeAreas.getEntireRow` возвращают другой объект `RangeAreas`, представляющий все столбцы (или строки) во всех диапазонах в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-121">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="7d770-122">Например, если `RangeAreas` представляет "A1:C4" и "F14:L15", то `RangeAreas.getEntireColumn` возвращает объект `RangeAreas`, представляющий "A:C" и "F:L".</span><span class="sxs-lookup"><span data-stu-id="7d770-122">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="7d770-123">`RangeAreas.copyFrom` может использовать параметр `Range` или `RangeAreas`, представляющий исходный диапазон (или диапазоны) операции копирования.</span><span class="sxs-lookup"><span data-stu-id="7d770-123">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="7d770-124">Полный список элементов Range, также доступных в RangeAreas</span><span class="sxs-lookup"><span data-stu-id="7d770-124">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="7d770-125">Свойства</span><span class="sxs-lookup"><span data-stu-id="7d770-125">Properties</span></span>

<span data-ttu-id="7d770-126">Ознакомьтесь со статьей [Чтение свойств RangeAreas](#read-properties-of-rangeareas) перед написанием кода, считывающего любое из перечисленных свойств.</span><span class="sxs-lookup"><span data-stu-id="7d770-126">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="7d770-127">Возвращаемое значение зависит от ряда факторов.</span><span class="sxs-lookup"><span data-stu-id="7d770-127">There are subtleties to what gets returned.</span></span>

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

##### <a name="methods"></a><span data-ttu-id="7d770-128">Методы</span><span class="sxs-lookup"><span data-stu-id="7d770-128">Methods</span></span>

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- <span data-ttu-id="7d770-129">`getOffsetRange()`(с `getOffsetRangeAreas` именем для `RangeAreas` объекта)</span><span class="sxs-lookup"><span data-stu-id="7d770-129">`getOffsetRange()` (named `getOffsetRangeAreas` on the `RangeAreas` object)</span></span>
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- <span data-ttu-id="7d770-130">`getUsedRange()`(с `getUsedRangeAreas` именем для `RangeAreas` объекта)</span><span class="sxs-lookup"><span data-stu-id="7d770-130">`getUsedRange()` (named `getUsedRangeAreas` on the `RangeAreas` object)</span></span>
- <span data-ttu-id="7d770-131">`getUsedRangeOrNullObject()`(с `getUsedRangeAreasOrNullObject` именем для `RangeAreas` объекта)</span><span class="sxs-lookup"><span data-stu-id="7d770-131">`getUsedRangeOrNullObject()` (named `getUsedRangeAreasOrNullObject` on the `RangeAreas` object)</span></span>
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="7d770-132">Свойства и методы, характерные для объекта RangeArea</span><span class="sxs-lookup"><span data-stu-id="7d770-132">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="7d770-133">Для типа `RangeAreas` существуют несколько свойств и методов, отсутствующих в объекте `Range`.</span><span class="sxs-lookup"><span data-stu-id="7d770-133">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="7d770-134">Ниже приведены некоторые из них.</span><span class="sxs-lookup"><span data-stu-id="7d770-134">The following is a selection of them:</span></span>

- <span data-ttu-id="7d770-135">`areas`. Объект `RangeCollection`, содержащий все диапазоны, которые представлены объектом `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-135">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="7d770-136">Объект `RangeCollection` — еще один новый объект, аналогичный другим объектам коллекции Excel.</span><span class="sxs-lookup"><span data-stu-id="7d770-136">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="7d770-137">У него есть свойство `items`, являющееся массивом объектов `Range`, которые представляют диапазоны.</span><span class="sxs-lookup"><span data-stu-id="7d770-137">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="7d770-138">`areaCount`. Общее количество диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-138">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="7d770-139">`getOffsetRangeAreas`. Действует аналогично методу [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), за исключением того, что возвращается объект `RangeAreas`, содержащий диапазоны, каждый из которых смещен относительно одного из диапазонов в исходном объекте `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-139">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="7d770-140">Создание RangeAreas</span><span class="sxs-lookup"><span data-stu-id="7d770-140">Create RangeAreas</span></span>

<span data-ttu-id="7d770-141">Объект `RangeAreas` можно создать двумя основными способами:</span><span class="sxs-lookup"><span data-stu-id="7d770-141">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="7d770-142">Вызвать метод `Worksheet.getRanges()` и передать ему строку с адресами диапазона, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="7d770-142">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="7d770-143">Если диапазон, который нужно включить, был преобразован [NamedItem](/javascript/api/excel/excel.nameditem), в строку можно включить имя вместо адреса.</span><span class="sxs-lookup"><span data-stu-id="7d770-143">If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="7d770-144">Вызвать метод `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="7d770-144">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="7d770-145">Этот метод возвращает объект `RangeAreas`, представляющий все диапазоны, выбранные в активном на данный момент листе.</span><span class="sxs-lookup"><span data-stu-id="7d770-145">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="7d770-146">После получения объекта `RangeAreas` можно создать другие с помощью методов объекта, возвращающих `RangeAreas`, например `getOffsetRangeAreas` и `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="7d770-146">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="7d770-147">Нельзя напрямую добавить дополнительные диапазоны к объекту `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-147">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="7d770-148">Например, у коллекции в `RangeAreas.areas` нет метода `add`.</span><span class="sxs-lookup"><span data-stu-id="7d770-148">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="7d770-149">Не пытайтесь напрямую добавлять или удалять элементы из массива `RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="7d770-149">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="7d770-150">Это приведет к нежелательному поведению кода. </span><span class="sxs-lookup"><span data-stu-id="7d770-150">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="7d770-151">Например, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, поскольку свойства и методы `RangeAreas` действуют, как будто новый элемент не был добавлен.</span><span class="sxs-lookup"><span data-stu-id="7d770-151">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="7d770-152">Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)` вызывает ошибку, если `index` больше, чем `areasCount-1`. </span><span class="sxs-lookup"><span data-stu-id="7d770-152">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="7d770-153">Аналогичным образом, удаление объекта `Range` в массиве `RangeAreas.areas.items` путем получения ссылки на него и вызова его метода `Range.delete` приводит к ошибкам: хотя объект `Range` *удален*, свойства и методы родительского объекта `RangeAreas` будут действовать (или пытаться действовать), как будто он еще существует.</span><span class="sxs-lookup"><span data-stu-id="7d770-153">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="7d770-154">Например, если код вызывает метод `RangeAreas.calculate`, Office попытается рассчитать диапазон, но это завершится ошибкой, поскольку объект range отсутствует.</span><span class="sxs-lookup"><span data-stu-id="7d770-154">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="7d770-155">Задание свойств для нескольких диапазонов</span><span class="sxs-lookup"><span data-stu-id="7d770-155">Set properties on multiple ranges</span></span>

<span data-ttu-id="7d770-156">Установка свойства для объекта `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-156">Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="7d770-157">Ниже приведен пример установки свойства в нескольких диапазонах.</span><span class="sxs-lookup"><span data-stu-id="7d770-157">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="7d770-158">Функция выделяет диапазоны **F3:F5** и **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="7d770-158">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="7d770-159">Этот пример применяется к сценариям, в которых можно жестко задать адреса диапазонов, передаваемых в `getRanges`, или легко рассчитать их во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="7d770-159">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="7d770-160">Ниже перечислены некоторые сценарии, в которых это возможно:</span><span class="sxs-lookup"><span data-stu-id="7d770-160">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="7d770-161">Код выполняется в контексте известного шаблона.</span><span class="sxs-lookup"><span data-stu-id="7d770-161">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="7d770-162">Код выполняется в контексте импортированных данных, в котором известна схема данных.</span><span class="sxs-lookup"><span data-stu-id="7d770-162">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="7d770-163">Получение специальных ячеек из нескольких диапазонов</span><span class="sxs-lookup"><span data-stu-id="7d770-163">Get special cells from multiple ranges</span></span>

<span data-ttu-id="7d770-164">Методы `getSpecialCells` и `getSpecialCellsOrNullObject` для объекта `RangeAreas` действуют аналогично методам с теми же названиями для объекта `Range`.</span><span class="sxs-lookup"><span data-stu-id="7d770-164">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="7d770-165">Эти методы возвращают ячейки с указанными характеристиками из всех диапазонов в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-165">These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="7d770-166">Дополнительные сведения о специальных ячейках см. в разделе [Поиск специальных ячеек в диапазоне](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range).</span><span class="sxs-lookup"><span data-stu-id="7d770-166">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) section for more details on special cells.</span></span>

<span data-ttu-id="7d770-167">При вызове метода `getSpecialCells` или `getSpecialCellsOrNullObject` для объекта `RangeAreas`:</span><span class="sxs-lookup"><span data-stu-id="7d770-167">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="7d770-168">Если в качестве первого параметра передается `Excel.SpecialCellType.sameConditionalFormat`, метод возвращает все ячейки с таким же условным форматированием, как у крайней левой верхней ячейки первого диапазона в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-168">If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>
- <span data-ttu-id="7d770-169">Если в качестве первого параметра передается `Excel.SpecialCellType.sameDataValidation`, метод возвращает все ячейки с таким же правилом проверки данных, как у крайней левой верхней ячейки первого диапазона в коллекции `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-169">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="7d770-170">Чтение свойств RangeAreas</span><span class="sxs-lookup"><span data-stu-id="7d770-170">Read properties of RangeAreas</span></span>

<span data-ttu-id="7d770-171">Чтение значений свойств `RangeAreas` требует внимания, так как определенное свойство может иметь разные значения для разных диапазонов в `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="7d770-171">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="7d770-172">Общее правило заключается в том, что если соответствующее значение *может* быть возвращено, оно будет возвращено.</span><span class="sxs-lookup"><span data-stu-id="7d770-172">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="7d770-173">Например, в приведенном ниже коде RGB-код для розового цвета (`#FFC0CB`) и `true` будут записаны в консоль, так как оба диапазона в объекте `RangeAreas` имеют розовую заливку и оба являются целыми столбцами.</span><span class="sxs-lookup"><span data-stu-id="7d770-173">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="7d770-174">Все усложняется, если согласование невозможно.</span><span class="sxs-lookup"><span data-stu-id="7d770-174">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="7d770-175">Свойства `RangeAreas` действуют в соответствии с приведенными ниже тремя принципами:</span><span class="sxs-lookup"><span data-stu-id="7d770-175">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="7d770-176">Логическое свойство объекта `RangeAreas` возвращает значение `false`, кроме случаев, когда свойство имеет значение true для всех диапазонов элементов.</span><span class="sxs-lookup"><span data-stu-id="7d770-176">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="7d770-177">Свойства, не являющиеся логическими, за исключением свойства `address`, возвращают значение `null`, кроме тех случаев, когда соответствующее свойство для всех диапазонов элементов обладает тем же значением.</span><span class="sxs-lookup"><span data-stu-id="7d770-177">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="7d770-178">Свойство `address` возвращает строку с адресами диапазонов элементов, разделенными запятыми.</span><span class="sxs-lookup"><span data-stu-id="7d770-178">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="7d770-179">Например, в приведенном ниже коде создается объект `RangeAreas`, в котором только один диапазон является целым столбцом и только один залит розовым цветом.</span><span class="sxs-lookup"><span data-stu-id="7d770-179">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="7d770-180">Консоль отобразит значение `null` для цвета заливки, `false` для свойства `isEntireRow` и "Sheet1!F3:F5, Sheet1!H:H" (при условии, что имя листа — "Sheet1") для свойства `address`.</span><span class="sxs-lookup"><span data-stu-id="7d770-180">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="7d770-181">См. также</span><span class="sxs-lookup"><span data-stu-id="7d770-181">See also</span></span>

- [<span data-ttu-id="7d770-182">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="7d770-182">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="7d770-183">Работа с диапазонами с использованием API JavaScript для Excel (основные задачи)</span><span class="sxs-lookup"><span data-stu-id="7d770-183">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="7d770-184">Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)</span><span class="sxs-lookup"><span data-stu-id="7d770-184">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)

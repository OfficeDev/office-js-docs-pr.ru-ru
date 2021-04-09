---
title: Поиск специальных ячеек в диапазоне с помощью API JavaScript Excel
description: Узнайте, как использовать API JavaScript Excel для поиска специальных ячеек, таких как ячейки с формулами, ошибками или числами.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6504873bcd8ab50bd4c03fe4f54b71d0bd920c5b
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652891"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="31e57-103">Поиск специальных ячеек в диапазоне с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="31e57-103">Find special cells within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="31e57-104">В этой статье данная статья содержит примеры кода, которые находят специальные ячейки в диапазоне с помощью API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="31e57-104">This article provides code samples that find special cells within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="31e57-105">Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="31e57-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="find-ranges-with-special-cells"></a><span data-ttu-id="31e57-106">Поиск диапазонов с помощью специальных ячеек</span><span class="sxs-lookup"><span data-stu-id="31e57-106">Find ranges with special cells</span></span>

<span data-ttu-id="31e57-107">Методы [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) и [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) находят диапазоны, основанные на характеристиках их клеток и типах значений их клеток.</span><span class="sxs-lookup"><span data-stu-id="31e57-107">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="31e57-108">Оба этих метода возвращают объекты `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="31e57-108">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="31e57-109">Подписи методов из файла типов данных TypeScript:</span><span class="sxs-lookup"><span data-stu-id="31e57-109">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="31e57-110">В следующем примере кода используется `getSpecialCells` метод для поиска всех ячеек с формулами.</span><span class="sxs-lookup"><span data-stu-id="31e57-110">The following code sample uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="31e57-111">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="31e57-111">About this code, note:</span></span>

- <span data-ttu-id="31e57-112">Он ограничивает часть листа, в которой требуется выполнять поиск, путем вызова сначала метода `Worksheet.getUsedRange`, а затем метода `getSpecialCells` только для этого диапазона.</span><span class="sxs-lookup"><span data-stu-id="31e57-112">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="31e57-113">Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами окрашены розовым цветом даже в том случае, если они не являются смежными.</span><span class="sxs-lookup"><span data-stu-id="31e57-113">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="31e57-114">Если в диапазоне нет ячеек с целевыми характеристиками, метод `getSpecialCells` выдает ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="31e57-114">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="31e57-115">Это приведет к переадресации потока управления к блоку `catch`, если таковой существует.</span><span class="sxs-lookup"><span data-stu-id="31e57-115">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="31e57-116">Если блокировки `catch` нет, ошибка останавливает метод.</span><span class="sxs-lookup"><span data-stu-id="31e57-116">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="31e57-117">Если ожидается, что всегда должны существовать ячейки с целевыми характеристиками, скорее всего вы захотите, чтобы код выдавал ошибку при их отсутствии.</span><span class="sxs-lookup"><span data-stu-id="31e57-117">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="31e57-118">Если отсутствие соответствующих ячеек является допустимым сценарием, ваш код должен проверить наличие такой возможности и корректно выполнить действие без выдачи ошибки.</span><span class="sxs-lookup"><span data-stu-id="31e57-118">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="31e57-119">Добиться такого поведения можно с помощью метода `getSpecialCellsOrNullObject` и возвращаемого им свойства `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="31e57-119">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="31e57-120">В следующем примере кода используется этот шаблон.</span><span class="sxs-lookup"><span data-stu-id="31e57-120">The following code sample uses this pattern.</span></span> <span data-ttu-id="31e57-121">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="31e57-121">About this code, note:</span></span>

- <span data-ttu-id="31e57-122">Метод всегда возвращает прокси-объект, поэтому он никогда не находится в `getSpecialCellsOrNullObject` `null` обычном смысле JavaScript.</span><span class="sxs-lookup"><span data-stu-id="31e57-122">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it's never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="31e57-123">Но если соответствующие ячейки не обнаружены, свойству `isNullObject` объекта присваивается значение `true`.</span><span class="sxs-lookup"><span data-stu-id="31e57-123">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="31e57-124">Он вызывает `context.sync` *перед* тестированием свойства `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="31e57-124">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="31e57-125">Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть.</span><span class="sxs-lookup"><span data-stu-id="31e57-125">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="31e57-126">Однако не нужно явно *загружать* `isNullObject` свойство.</span><span class="sxs-lookup"><span data-stu-id="31e57-126">However, it's not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="31e57-127">Он автоматически загружается объектом, даже если он не `context.sync` `load` вызван.</span><span class="sxs-lookup"><span data-stu-id="31e57-127">It's automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="31e57-128">Дополнительные сведения см. в дополнительных сведениях о методах [ \* и свойствах OrNullObject.](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)</span><span class="sxs-lookup"><span data-stu-id="31e57-128">For more information, see [\*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).</span></span>
- <span data-ttu-id="31e57-129">Этот код можно проверить, выбрав сначала диапазон без ячеек с формулами и запустив его.</span><span class="sxs-lookup"><span data-stu-id="31e57-129">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="31e57-130">Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.</span><span class="sxs-lookup"><span data-stu-id="31e57-130">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
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

<span data-ttu-id="31e57-131">Для простоты все остальные образцы кода в этой статье используют `getSpecialCells` метод вместо  `getSpecialCellsOrNullObject` .</span><span class="sxs-lookup"><span data-stu-id="31e57-131">For simplicity, all other code samples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

## <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="31e57-132">Ограничение целевых ячеек с помощью типа значений ячеек</span><span class="sxs-lookup"><span data-stu-id="31e57-132">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="31e57-133">Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` принимают необязательный второй параметр, используемый для дополнительного ограничения целевых ячеек.</span><span class="sxs-lookup"><span data-stu-id="31e57-133">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="31e57-134">Этот второй параметр `Excel.SpecialCellValueType` используется для указания того, что требуются только ячейки, содержащие определенные типы значений.</span><span class="sxs-lookup"><span data-stu-id="31e57-134">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="31e57-135">Параметр `Excel.SpecialCellValueType` можно использовать, только если для параметра `Excel.SpecialCellType` задано значение `Excel.SpecialCellType.formulas` или `Excel.SpecialCellType.constants`.</span><span class="sxs-lookup"><span data-stu-id="31e57-135">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="31e57-136">Тестирование для ячеек с одним типом значений</span><span class="sxs-lookup"><span data-stu-id="31e57-136">Test for a single cell value type</span></span>

<span data-ttu-id="31e57-137">Для перечисления `Excel.SpecialCellValueType` существует четыре основных типа (в дополнение к другим объединенным значениям, описанным ниже в этом разделе):</span><span class="sxs-lookup"><span data-stu-id="31e57-137">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="31e57-138">`Excel.SpecialCellValueType.logical` (означает логическое значение)</span><span class="sxs-lookup"><span data-stu-id="31e57-138">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="31e57-139">В следующем примере кода находятся специальные ячейки, которые являются числовыми константами, и цвета этих клеток розовыми.</span><span class="sxs-lookup"><span data-stu-id="31e57-139">The following code sample finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="31e57-140">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="31e57-140">About this code, note:</span></span>

- <span data-ttu-id="31e57-141">Он выделяет только ячейки, которые имеют буквальное значение числа.</span><span class="sxs-lookup"><span data-stu-id="31e57-141">It only highlights cells that have a literal number value.</span></span> <span data-ttu-id="31e57-142">В нем не будут выделены ячейки, у них есть формула (даже если результат — число) или клеток состояния boolean, text или error.</span><span class="sxs-lookup"><span data-stu-id="31e57-142">It won't highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="31e57-143">Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.</span><span class="sxs-lookup"><span data-stu-id="31e57-143">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="31e57-144">Тестирование для ячеек с несколькими типами значений</span><span class="sxs-lookup"><span data-stu-id="31e57-144">Test for multiple cell value types</span></span>

<span data-ttu-id="31e57-145">Иногда требуется работать с ячейками, имеющими несколько типов значений, например со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями (`Excel.SpecialCellValueType.logical`).</span><span class="sxs-lookup"><span data-stu-id="31e57-145">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="31e57-146">Для перечисления `Excel.SpecialCellValueType` существуют значения с объединенными типами.</span><span class="sxs-lookup"><span data-stu-id="31e57-146">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="31e57-147">Например, `Excel.SpecialCellValueType.logicalText` обрабатывает все ячейки с логическими и текстовыми значениями.</span><span class="sxs-lookup"><span data-stu-id="31e57-147">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="31e57-148">`Excel.SpecialCellValueType.all` является значением по умолчанию, которое не ограничивает возвращаемые типы значений ячеек.</span><span class="sxs-lookup"><span data-stu-id="31e57-148">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="31e57-149">В следующем примере кода цвета всех ячеек с формулами, которые производят количество или значение boolean.</span><span class="sxs-lookup"><span data-stu-id="31e57-149">The following code sample colors all cells with formulas that produce number or boolean value.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="see-also"></a><span data-ttu-id="31e57-150">См. также</span><span class="sxs-lookup"><span data-stu-id="31e57-150">See also</span></span>

- [<span data-ttu-id="31e57-151">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="31e57-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="31e57-152">Работа с ячейками с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="31e57-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="31e57-153">Поиск строки с помощью API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="31e57-153">Find a string using the Excel JavaScript API</span></span>](excel-add-ins-ranges-string-match.md)
- [<span data-ttu-id="31e57-154">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="31e57-154">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)

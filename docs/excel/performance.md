---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизация производительности с помощью API Excel JavaScript
ms.date: 03/28/2018
ms.openlocfilehash: dabbb69f8dee0df782a265edcfdfb1c89894e915
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437411"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="b62b4-103">Оптимизация производительности с использованием API Excel JavaScript</span><span class="sxs-lookup"><span data-stu-id="b62b4-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="b62b4-104">Существует несколько способов выполнения общих задач с помощью API Excel JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b62b4-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="b62b4-105">Вы найдете значительные различия в производительности между различными подходами.</span><span class="sxs-lookup"><span data-stu-id="b62b4-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="b62b4-106">В этой статье приведены примеры руководств и кода, чтобы показать вам, как эффективно выполнять общие задачи с помощью API Excel JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b62b4-106">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="b62b4-107">Минимизировать количество вызовов sync ()</span><span class="sxs-lookup"><span data-stu-id="b62b4-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="b62b4-108">В Excel JavaScript API, ```sync()``` является единственной асинхронной операцией, и в некоторых случаях она может быть медленной, особенно для Excel Online.</span><span class="sxs-lookup"><span data-stu-id="b62b4-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="b62b4-109">Чтобы оптимизировать производительность, минимизируйте количество вызовов ```sync()```, поставив в очередь столько изменений, сколько возможно, прежде чем вызвать его.</span><span class="sxs-lookup"><span data-stu-id="b62b4-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="b62b4-110">См.статью [Основные понятия - синхронизация ()](excel-add-ins-core-concepts.md#sync) для образцов кода, которые следуют этой практике.</span><span class="sxs-lookup"><span data-stu-id="b62b4-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="b62b4-111">Минимизировать количество созданных прокси-объектов</span><span class="sxs-lookup"><span data-stu-id="b62b4-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="b62b4-112">Избегайте повторного создания одного и того же прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="b62b4-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="b62b4-113">Вместо этого, если вам нужен один и тот же прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.</span><span class="sxs-lookup"><span data-stu-id="b62b4-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```javascript
// BAD: repeated calls to .getRange() to create the same proxy object
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: create the range proxy object once and assign to a variable
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

## <a name="load-necessary-properties-only"></a><span data-ttu-id="b62b4-114">Загрузка только необходимых свойств</span><span class="sxs-lookup"><span data-stu-id="b62b4-114">Load necessary properties only</span></span>

<span data-ttu-id="b62b4-115">В Excel JavaScript API вам необходимо явно загрузить свойства прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="b62b4-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="b62b4-116">Хотя вы можете сразу загрузить все свойства с пустым ```load()``` вызовом, этот подход может иметь значительные накладные расходы.</span><span class="sxs-lookup"><span data-stu-id="b62b4-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="b62b4-117">Вместо этого мы предлагаем вам загружать только необходимые свойства, особенно для тех объектов, которые имеют большое количество свойств.</span><span class="sxs-lookup"><span data-stu-id="b62b4-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="b62b4-118">Например, если вы собираетесь считать свойство **address** объекта range, при вызове метода **load()** укажите только это свойство:</span><span class="sxs-lookup"><span data-stu-id="b62b4-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="b62b4-119">Вы можете вызвать метод **load()** любым из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="b62b4-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="b62b4-120">_Синтаксис:_</span><span class="sxs-lookup"><span data-stu-id="b62b4-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="b62b4-121">_Где:_</span><span class="sxs-lookup"><span data-stu-id="b62b4-121">_Where:_</span></span>
 
* <span data-ttu-id="b62b4-122">`properties` Это список свойств для загрузки, указанных как строки с разделителями-запятыми или как массив имен.</span><span class="sxs-lookup"><span data-stu-id="b62b4-122">`properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names.</span></span> <span data-ttu-id="b62b4-123">Дополнительные сведения см. в описаниях методов **load()**, определенных для объектов в [справочнике по API JavaScript для Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="b62b4-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="b62b4-p106">`loadOption` указывает объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](https://dev.office.com/reference/add-ins/excel/loadoption).</span><span class="sxs-lookup"><span data-stu-id="b62b4-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://dev.office.com/reference/add-ins/excel/loadoption) for details.</span></span>

<span data-ttu-id="b62b4-126">Помните, что некоторые «свойства» под объектом могут иметь то же имя, что и другой объект.</span><span class="sxs-lookup"><span data-stu-id="b62b4-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="b62b4-127">Например, `format` является свойством объекта диапазона, но `format` сам по себе является объектом.</span><span class="sxs-lookup"><span data-stu-id="b62b4-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="b62b4-128">Итак, если вы вызываете например, `range.load("format")`, это эквивалентно `range.format.load()`, который представляет собой вызов пустой нагрузки (), который может вызвать проблемы с производительностью, как описано ранее.</span><span class="sxs-lookup"><span data-stu-id="b62b4-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="b62b4-129">Чтобы избежать  этого, ваш код должен загружать только «листовые узлы» в представлении объектов.</span><span class="sxs-lookup"><span data-stu-id="b62b4-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="b62b4-130">Временно приостанавливать расчет</span><span class="sxs-lookup"><span data-stu-id="b62b4-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="b62b4-131">Если вы пытаетесь выполнить операцию на большом количестве ячеек (например, установив значение огромного объекта диапазона), и вы не возражаете временно приостановить вычисление в Excel во время завершения операции, мы рекомендуем приостановить расчет до следующего ```context.sync()``` вызова.</span><span class="sxs-lookup"><span data-stu-id="b62b4-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="b62b4-132">См.статью [Объект приложения](https://dev.office.com/reference/add-ins/excel/application), справочную документацию для получения информации о том, как использовать ```suspendApiCalculationUntilNextSync()``` API для приостановки и повторного включения вычислений очень удобным способом.</span><span class="sxs-lookup"><span data-stu-id="b62b4-132">See [Application Object](https://dev.office.com/reference/add-ins/excel/application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="b62b4-133">Следующий код демонстрирует, как временно приостановить расчет:</span><span class="sxs-lookup"><span data-stu-id="b62b4-133">The following code demonstrates how to suspend calculation temporarily:</span></span>

```js
Excel.run(async function(ctx) {
    var app = ctx.workbook.application;
    var sheet = ctx.workbook.worksheets.getItem("sheet1");
    var rangeToSet: Excel.Range;
    var rangeToGet: Excel.Range;
    app.load("calculationMode");
    await ctx.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);
    
    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalc
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with supend recalc
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="b62b4-134">Изменение всех ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="b62b4-134">Update all cells in a range</span></span> 

<span data-ttu-id="b62b4-135">Когда вам нужно обновить все ячейки в диапазоне с одним и тем же значением или свойством, может уйти много времени на выполнение этого с помощью двумерного массива, который многократно задает одно и то же значение, поскольку для этого подхода Excel требует итерации по всем ячейкам в диапазон для установки каждой отдельно.</span><span class="sxs-lookup"><span data-stu-id="b62b4-135">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="b62b4-136">Excel имеет более эффективный способ обновления всех ячеек в диапазоне с тем же значением или свойством.</span><span class="sxs-lookup"><span data-stu-id="b62b4-136">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="b62b4-137">Если вам нужно применить одно и то же значение, тот же формат номера или ту же формулу для диапазона ячеек, более эффективно указывать одно значение вместо массива значений.</span><span class="sxs-lookup"><span data-stu-id="b62b4-137">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="b62b4-138">Это значительно улучшит производительность.</span><span class="sxs-lookup"><span data-stu-id="b62b4-138">Doing so will significantly improve performance.</span></span> <span data-ttu-id="b62b4-139">Для примера кода, который показывает этот подход в действии, см.статью [Основные понятия -Обновление всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="b62b4-139">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="b62b4-140">Обычный сценарий, в котором вы можете применить этот подход, - это установка разных форматов чисел в разных столбцах на листе.</span><span class="sxs-lookup"><span data-stu-id="b62b4-140">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="b62b4-141">В этом случае вы можете просто выполнить итерацию столбцов и задавать формат чисел для каждого столбца с одним значением.</span><span class="sxs-lookup"><span data-stu-id="b62b4-141">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="b62b4-142">Обрабатывайте каждый столбец как диапазон, как показано в [Обновление всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) образца кода.</span><span class="sxs-lookup"><span data-stu-id="b62b4-142">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="b62b4-143">Если вы используете TypeScript, вы заметите ошибку компиляции, заявив, что одно значение не может быть установлено в 2D-массив.</span><span class="sxs-lookup"><span data-stu-id="b62b4-143">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="b62b4-144">Это неизбежно, поскольку значения *находятся* 2D-массив при извлечении свойств, а TypeScript не допускает использование разных типов setter vs getter.</span><span class="sxs-lookup"><span data-stu-id="b62b4-144">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="b62b4-145">Однако простым обходным путем является установление значений с помощью суффикса`as any`, например, `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="b62b4-145">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="b62b4-146">Импорт данных в таблицы</span><span class="sxs-lookup"><span data-stu-id="b62b4-146">Importing data into tables</span></span>

<span data-ttu-id="b62b4-147">При попытке импортировать огромное количество данных непосредственно в [Таблицу](https://dev.office.com/reference/add-ins/excel/table) объекта (например, используя `TableRowCollection.add()`), вы можете столкнуться с низкой производительностью.</span><span class="sxs-lookup"><span data-stu-id="b62b4-147">When trying to import a huge amount of data directly into a [Table](https://dev.office.com/reference/add-ins/excel/table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="b62b4-148">Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону.</span><span class="sxs-lookup"><span data-stu-id="b62b4-148">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="b62b4-149">Если вы пытаетесь записать данные в существующую таблицу, напишите данные в объект диапазона через `table.getDataBodyRange()`, и таблица будет автоматически расшириться.</span><span class="sxs-lookup"><span data-stu-id="b62b4-149">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="b62b4-150">Вот пример такого подхода:</span><span class="sxs-lookup"><span data-stu-id="b62b4-150">Here is an example in JavaScript of this operation.</span></span>

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first 
    var range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    var table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await ctx.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await ctx.sync();
})
```

> [!NOTE]
> <span data-ttu-id="b62b4-151">Вы можете удобно преобразовать объект Table в объект Range, используя [метод Table.convertToRange ()](https://dev.office.com/reference/add-ins/excel/table#converttorange) .</span><span class="sxs-lookup"><span data-stu-id="b62b4-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="b62b4-152">См. также</span><span class="sxs-lookup"><span data-stu-id="b62b4-152">See also</span></span>

- [<span data-ttu-id="b62b4-153">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="b62b4-153">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b62b4-154">Сложные понятия, связанные с API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="b62b4-154">Excel JavaScript API advanced concepts</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="b62b4-155">Открытая спецификация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="b62b4-155">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="b62b4-156">Объект Worksheet Functions (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="b62b4-156">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/functions)

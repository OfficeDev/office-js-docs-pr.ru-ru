---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизируйте производительность с использованием API JavaScript для Excel
ms.date: 11/29/2018
ms.openlocfilehash: fb0f81b79d2eac847a91a7b2a4fab92362330a10
ms.sourcegitcommit: e2ba9d7210c921d068f40d9f689314c73ad5ab4a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/05/2018
ms.locfileid: "27156581"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="25419-103">Оптимизация производительности с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="25419-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="25419-104">Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="25419-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="25419-105">Вы обнаружите существенные различия в производительности между разными подходами.</span><span class="sxs-lookup"><span data-stu-id="25419-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="25419-106">В этой статье приведены инструкции и примеры кода, показывающие, как эффективно выполнять стандартные задачи, используя API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="25419-106">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="25419-107">Минимизация количества вызовов sync()</span><span class="sxs-lookup"><span data-stu-id="25419-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="25419-108">В API JavaScript для Excel ```sync()``` является единственной асинхронной операцией и в некоторых обстоятельствах может выполняться медленно, особенно в случае с Excel Online.</span><span class="sxs-lookup"><span data-stu-id="25419-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="25419-109">Для оптимизации производительности минимизируйте количество вызовов ```sync()```, поставив в очередь максимально возможное количество изменений до ее вызова.</span><span class="sxs-lookup"><span data-stu-id="25419-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="25419-110">Примеры кода, использующие этот подход, см. в статье [Основные концепции — sync()](excel-add-ins-core-concepts.md#sync).</span><span class="sxs-lookup"><span data-stu-id="25419-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="25419-111">Минимизация количества созданных прокси-объектов</span><span class="sxs-lookup"><span data-stu-id="25419-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="25419-112">Избегайте повторного создания одного и того же прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="25419-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="25419-113">Вместо этого, если вам нужен одинаковый прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.</span><span class="sxs-lookup"><span data-stu-id="25419-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="25419-114">Загрузка только необходимых свойств</span><span class="sxs-lookup"><span data-stu-id="25419-114">Load necessary properties only</span></span>

<span data-ttu-id="25419-115">В API JavaScript для Excel необходимо явно загрузить свойства прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="25419-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="25419-116">Несмотря на то, что вы можете загрузить все свойства одновременно, сделав пустой вызов ```load()```, этот подход может значительно замедлить производительность.</span><span class="sxs-lookup"><span data-stu-id="25419-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="25419-117">Вместо этого предлагается загружать только необходимые свойства, особенно для объектов с большим количеством свойств.</span><span class="sxs-lookup"><span data-stu-id="25419-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="25419-118">Например, если вы собираетесь считать свойство **address** объекта range, при вызове метода **load()** укажите только это свойство:</span><span class="sxs-lookup"><span data-stu-id="25419-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="25419-119">Вы можете вызвать метод **load()** любым из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="25419-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="25419-120">_Синтаксис:_</span><span class="sxs-lookup"><span data-stu-id="25419-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="25419-121">_Где:_</span><span class="sxs-lookup"><span data-stu-id="25419-121">_Where:_</span></span>
 
* <span data-ttu-id="25419-122">`properties` — это список свойств для загрузки, указанных как строки с разделителями-запятыми или как массив имен.</span><span class="sxs-lookup"><span data-stu-id="25419-122">`properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names.</span></span> <span data-ttu-id="25419-123">Дополнительные сведения см. в описаниях методов **load()**, определенных для объектов, в [справочнике по API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="25419-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="25419-p106">`loadOption` указывает объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в статье, посвященной [параметрам](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) загрузки объектов.</span><span class="sxs-lookup"><span data-stu-id="25419-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="25419-126">Имейте в виду, что некоторые "свойства" объекта могут совпадать с именем другого объекта.</span><span class="sxs-lookup"><span data-stu-id="25419-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="25419-127">Например, `format` — это свойство объекта range, но также имеется и объект `format`.</span><span class="sxs-lookup"><span data-stu-id="25419-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="25419-128">Поэтому если вы, например, вызываете `range.load("format")`, это эквивалентно `range.format.load()`, являющемуся пустым вызовом load(), который может стать причиной проблем с производительностью, как описано ранее.</span><span class="sxs-lookup"><span data-stu-id="25419-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="25419-129">Чтобы избежать этого, ваш код должен загружать только "конечные узлы" в дереве объектов.</span><span class="sxs-lookup"><span data-stu-id="25419-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="25419-130">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="25419-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="25419-131">Если вы пытаетесь выполнить операцию с большим количеством ячеек (например, установка значения огромного объекта range) и не возражаете временно приостановить расчеты в Excel до завершения операции, рекомендуется приостановить вычисление до следующего вызова ```context.sync()```.</span><span class="sxs-lookup"><span data-stu-id="25419-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="25419-132">Дополнительные сведения об использовании API ```suspendApiCalculationUntilNextSync()``` для приостановки и повторного включения вычислений удобным способом см. в справочном документе [Объект Application](https://docs.microsoft.com/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="25419-132">See [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="25419-133">В приведенном ниже коде показано, как временно приостановить вычисление:</span><span class="sxs-lookup"><span data-stu-id="25419-133">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="25419-134">Изменение всех ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="25419-134">Update all cells in a range</span></span> 

<span data-ttu-id="25419-135">Если нужно изменить все ячейки в диапазоне с использованием одинакового значения или свойства, это может занять много времени при применении двумерного массива, многократно задающего одно и то же значение, поскольку в этом способе Excel требуется выполнять итерации по всем ячейкам в диапазоне для установки каждой отдельно.</span><span class="sxs-lookup"><span data-stu-id="25419-135">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="25419-136">В Excel есть более эффективный способ изменения всех ячеек в диапазоне с использованием одинакового значения или свойства.</span><span class="sxs-lookup"><span data-stu-id="25419-136">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="25419-137">Если нужно применить одинаковое значение, одинаковый числовой формат или одинаковую формулу для диапазона ячеек, эффективнее указывать одно значение вместо массива значений.</span><span class="sxs-lookup"><span data-stu-id="25419-137">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="25419-138">Это значительно повысит производительность.</span><span class="sxs-lookup"><span data-stu-id="25419-138">Doing so will significantly improve performance.</span></span> <span data-ttu-id="25419-139">Пример кода, демонстрирующий этот способ в действии, см. в статье [Основные концепции — Изменение всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="25419-139">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="25419-140">Распространенным сценарием применения этого способа является установка разных числовых форматов в разных столбцах на листе.</span><span class="sxs-lookup"><span data-stu-id="25419-140">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="25419-141">В этом случае можно просто выполнить итерацию столбцов и установить числовой формат для каждого столбца с помощью одного значения.</span><span class="sxs-lookup"><span data-stu-id="25419-141">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="25419-142">Обработайте каждый столбец в качестве диапазона, как показано в примере кода [Изменение всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="25419-142">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="25419-143">При использовании TypeScript вы заметите ошибку компиляции с сообщением, что одно значение не может быть установлено в двумерный массив.</span><span class="sxs-lookup"><span data-stu-id="25419-143">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="25419-144">Это неизбежно, поскольку значения *являются* двумерным массивом при извлечении свойств, а TypeScript не допускает использования разных типов методов задания и получения.</span><span class="sxs-lookup"><span data-stu-id="25419-144">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="25419-145">Однако есть простой обходной путь — установка значений с суффиксом `as any`, например `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="25419-145">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="25419-146">Импорт данных в таблицы</span><span class="sxs-lookup"><span data-stu-id="25419-146">Importing data into tables</span></span>

<span data-ttu-id="25419-147">При попытке импортировать огромное количество данных непосредственно в объект [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) (например, с помощью `TableRowCollection.add()`) можно столкнуться с низкой производительностью.</span><span class="sxs-lookup"><span data-stu-id="25419-147">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="25419-148">Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону.</span><span class="sxs-lookup"><span data-stu-id="25419-148">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="25419-149">Если вы пытаетесь записать данные в существующую таблицу, запишите данные в объект range с помощью `table.getDataBodyRange()`, и таблица расширится автоматически.</span><span class="sxs-lookup"><span data-stu-id="25419-149">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="25419-150">Ниже приведен пример такого способа.</span><span class="sxs-lookup"><span data-stu-id="25419-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="25419-151">Можно легко преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="25419-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="25419-152">Прекращение отслеживания ненужных диапазонов</span><span class="sxs-lookup"><span data-stu-id="25419-152">Untrack unneeded ranges</span></span>

<span data-ttu-id="25419-153">Слой JavaScript создает прокси-объекты для вашей надстройки для взаимодействия с книгой Excel и базовыми диапазонами.</span><span class="sxs-lookup"><span data-stu-id="25419-153">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="25419-154">Эти объекты хранятся в памяти до вызова `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="25419-154">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="25419-155">Операции с большими пакетами могут создавать много прокси-объектов, необходимых надстройке лишь один раз, которые можно удалить из памяти до выполнения пакетных действий.</span><span class="sxs-lookup"><span data-stu-id="25419-155">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="25419-156">Метод [Range.untrack()](/javascript/api/excel/excel.range#untrack--) удаляет объект Excel Range из памяти.</span><span class="sxs-lookup"><span data-stu-id="25419-156">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="25419-157">Вызов этого метода после завершения действий надстройки с диапазоном должен приводить к заметному повышению производительности при использовании большого количества объектов Range.</span><span class="sxs-lookup"><span data-stu-id="25419-157">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span> 

> [!NOTE]
> <span data-ttu-id="25419-158">`Range.untrack()` — это ярлык для [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span><span class="sxs-lookup"><span data-stu-id="25419-158">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="25419-159">Отслеживание любого прокси-объекта можно прекратить, удалив его из списка отслеживаемых объектов в контексте.</span><span class="sxs-lookup"><span data-stu-id="25419-159">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="25419-160">Обычно объекты Range являются единственными объектами Excel, используемыми в достаточных количествах для применения прекращения отслеживания.</span><span class="sxs-lookup"><span data-stu-id="25419-160">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="25419-161">В приведенном ниже примере кода выбранный диапазон заполняется данными по одной ячейке.</span><span class="sxs-lookup"><span data-stu-id="25419-161">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="25419-162">После добавления значения в ячейку, диапазон отображает, что отслеживание ячейки прекращено.</span><span class="sxs-lookup"><span data-stu-id="25419-162">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="25419-163">Выполните этот код с выбранным диапазоном от 10 000 до 20 000 ячеек сначала со строкой `cell.untrack()`, а затем без нее.</span><span class="sxs-lookup"><span data-stu-id="25419-163">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="25419-164">Вы должны заметить, что код выполняется с использованием строки `cell.untrack()` быстрее, чем без нее.</span><span class="sxs-lookup"><span data-stu-id="25419-164">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="25419-165">Вы также можете заметить уменьшение времени отклика в конце, так как этап очистки занимает меньше времени.</span><span class="sxs-lookup"><span data-stu-id="25419-165">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();
    
    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // call untrack() to release the range from memory
            cell.untrack();
        }
    }

    await context.sync();
});
```

## <a name="enable-and-disable-events"></a><span data-ttu-id="25419-166">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="25419-166">Enable and disable events</span></span>

<span data-ttu-id="25419-167">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="25419-167">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="25419-168">Пример кода, в котором показано, как включить и отключить события, см. в статье [Работа с событиями](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="25419-168">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="25419-169">См. также</span><span class="sxs-lookup"><span data-stu-id="25419-169">See also</span></span>

- [<span data-ttu-id="25419-170">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="25419-170">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="25419-171">Дополнительные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="25419-171">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="25419-172">Открытая спецификация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="25419-172">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="25419-173">Объект Worksheet Functions (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="25419-173">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)

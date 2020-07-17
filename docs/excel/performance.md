---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизируйте производительность с использованием API JavaScript для Excel
ms.date: 07/14/2020
localization_priority: Normal
ms.openlocfilehash: 193cbe8c8cd1a432c6567401ed645990cb93e5e9
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159096"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="b7937-103">Оптимизация производительности с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="b7937-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="b7937-104">Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="b7937-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="b7937-105">Вы обнаружите существенные различия в производительности между разными подходами.</span><span class="sxs-lookup"><span data-stu-id="b7937-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="b7937-106">В этой статье приведены инструкции и примеры кода, показывающие, как эффективно выполнять стандартные задачи, используя API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="b7937-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="b7937-107">Минимизация количества вызовов sync()</span><span class="sxs-lookup"><span data-stu-id="b7937-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="b7937-108">В API JavaScript для Excel `sync()` является единственной асинхронной операцией и в некоторых обстоятельствах может выполняться медленно, особенно в случае с Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="b7937-108">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="b7937-109">Для оптимизации производительности минимизируйте количество вызовов `sync()`, поставив в очередь максимально возможное количество изменений до ее вызова.</span><span class="sxs-lookup"><span data-stu-id="b7937-109">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="b7937-110">Примеры кода, использующие этот подход, см. в статье [Основные концепции — sync()](excel-add-ins-core-concepts.md#sync).</span><span class="sxs-lookup"><span data-stu-id="b7937-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="b7937-111">Минимизация количества созданных прокси-объектов</span><span class="sxs-lookup"><span data-stu-id="b7937-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="b7937-112">Избегайте повторного создания одного и того же прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="b7937-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="b7937-113">Вместо этого, если вам нужен одинаковый прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.</span><span class="sxs-lookup"><span data-stu-id="b7937-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```js
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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="b7937-114">Загрузка только необходимых свойств</span><span class="sxs-lookup"><span data-stu-id="b7937-114">Load necessary properties only</span></span>

<span data-ttu-id="b7937-115">В API JavaScript для Excel необходимо явно загрузить свойства прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="b7937-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="b7937-116">Несмотря на то, что вы можете загрузить все свойства одновременно, сделав пустой вызов `load()`, этот подход может значительно замедлить производительность.</span><span class="sxs-lookup"><span data-stu-id="b7937-116">Although you're able to load all the properties at once with an empty `load()` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="b7937-117">Вместо этого предлагается загружать только необходимые свойства, особенно для объектов с большим количеством свойств.</span><span class="sxs-lookup"><span data-stu-id="b7937-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="b7937-118">Например, если требуется только прочитать `address` свойство объекта Range, при вызове метода укажите только это свойство `load()` :</span><span class="sxs-lookup"><span data-stu-id="b7937-118">For example, if you only intend to read the `address` property of a range object, specify only that property when you call the `load()` method:</span></span>

```js
range.load('address');
```

<span data-ttu-id="b7937-119">Метод можно вызвать `load()` одним из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="b7937-119">You can call `load()` method in any of the following ways:</span></span>

<span data-ttu-id="b7937-120">_Синтаксис:_</span><span class="sxs-lookup"><span data-stu-id="b7937-120">_Syntax:_</span></span>

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

<span data-ttu-id="b7937-121">_Где:_</span><span class="sxs-lookup"><span data-stu-id="b7937-121">_Where:_</span></span>

* <span data-ttu-id="b7937-122">`properties` — это список свойств для загрузки, указанных как строки с разделителями-запятыми или как массив имен.</span><span class="sxs-lookup"><span data-stu-id="b7937-122">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names.</span></span> <span data-ttu-id="b7937-123">Дополнительные сведения приведены в статье `load()` методы, определенные для объектов в [справочнике по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md).</span><span class="sxs-lookup"><span data-stu-id="b7937-123">For more information, see the `load()` methods defined for objects in [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md).</span></span>
* <span data-ttu-id="b7937-p106">`loadOption` указывает объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в статье, посвященной [параметрам](/javascript/api/office/officeextension.loadoption) загрузки объектов.</span><span class="sxs-lookup"><span data-stu-id="b7937-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="b7937-126">Обратите внимание, что некоторые "Свойства" в объекте могут иметь такое же имя, что и другой объект.</span><span class="sxs-lookup"><span data-stu-id="b7937-126">Please be aware that some of the "properties" under an object may have the same name as another object.</span></span> <span data-ttu-id="b7937-127">Например, `format` — это свойство объекта range, но также имеется и объект `format`.</span><span class="sxs-lookup"><span data-stu-id="b7937-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="b7937-128">Поэтому если вы, например, вызываете `range.load("format")`, это эквивалентно `range.format.load()`, являющемуся пустым вызовом load(), который может стать причиной проблем с производительностью, как описано ранее.</span><span class="sxs-lookup"><span data-stu-id="b7937-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="b7937-129">Чтобы избежать этого, код должен загружать только "конечные узлы" в дереве объектов.</span><span class="sxs-lookup"><span data-stu-id="b7937-129">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="b7937-130">Временная приостановка процессов Excel</span><span class="sxs-lookup"><span data-stu-id="b7937-130">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="b7937-131">В Excel есть несколько фоновых задач, которые реагируют на ввод, выполняемый как пользователями, так и надстройкой.</span><span class="sxs-lookup"><span data-stu-id="b7937-131">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="b7937-132">Для повышения производительности можно управлять некоторыми из этих процессов Excel.</span><span class="sxs-lookup"><span data-stu-id="b7937-132">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="b7937-133">Это особенно полезно, если ваша надстройка работает с большими наборами данных.</span><span class="sxs-lookup"><span data-stu-id="b7937-133">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="b7937-134">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="b7937-134">Suspend calculation temporarily</span></span>

<span data-ttu-id="b7937-135">Если вы пытаетесь выполнить операцию с большим количеством ячеек (например, установка значения огромного объекта range) и не возражаете временно приостановить расчеты в Excel до завершения операции, рекомендуется приостановить вычисление до следующего вызова `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="b7937-135">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="b7937-136">Дополнительные сведения об использовании API `suspendApiCalculationUntilNextSync()` для приостановки и повторного включения вычислений удобным способом см. в справочном документе [Объект Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="b7937-136">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="b7937-137">В приведенном ниже коде показано, как временно приостановить вычисление:</span><span class="sxs-lookup"><span data-stu-id="b7937-137">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

<span data-ttu-id="b7937-138">Обратите внимание, что приостанавливаются только вычисления формул.</span><span class="sxs-lookup"><span data-stu-id="b7937-138">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="b7937-139">Все измененные ссылки все еще перестраиваются.</span><span class="sxs-lookup"><span data-stu-id="b7937-139">Any altered references are still rebuilt.</span></span> <span data-ttu-id="b7937-140">Например, при переименовании листа все ссылки в формулах будут обновляться на этом листе.</span><span class="sxs-lookup"><span data-stu-id="b7937-140">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="b7937-141">Приостановка обновления экрана</span><span class="sxs-lookup"><span data-stu-id="b7937-141">Suspend screen updating</span></span>

<span data-ttu-id="b7937-142">Excel отображает изменения, производимые вашей надстройкой, примерно по мере их выполнения в коде.</span><span class="sxs-lookup"><span data-stu-id="b7937-142">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="b7937-143">Для больших циклических наборов данных может не требоваться просмотр хода выполнения на экране в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="b7937-143">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="b7937-144">Параметр `Application.suspendScreenUpdatingUntilNextSync()` приостанавливает визуальные обновления для Excel до вызова надстройкой метода `context.sync()` или завершения метода `Excel.run` (неявно вызывающего `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="b7937-144">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="b7937-145">Необходимо учитывать, что Excel не будет проявлять признаков работы до следующей синхронизации. Ваша надстройка должна либо предоставить пользователям инструкции, оповещающие их об этой задержке, либо отобразить строку состояния, демонстрирующую активность.</span><span class="sxs-lookup"><span data-stu-id="b7937-145">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="b7937-146">Не вызывайте их `suspendScreenUpdatingUntilNextSync` повторно (например, в цикле).</span><span class="sxs-lookup"><span data-stu-id="b7937-146">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="b7937-147">Повторные вызовы приведут к мерцанию окна Excel.</span><span class="sxs-lookup"><span data-stu-id="b7937-147">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="b7937-148">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="b7937-148">Enable and disable events</span></span>

<span data-ttu-id="b7937-149">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="b7937-149">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="b7937-150">Пример кода, в котором показано, как включить и отключить события, см. в статье [Работа с событиями](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="b7937-150">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="b7937-151">Импорт данных в таблицы</span><span class="sxs-lookup"><span data-stu-id="b7937-151">Importing data into tables</span></span>

<span data-ttu-id="b7937-152">При попытке импортировать огромное количество данных непосредственно в объект [Table](/javascript/api/excel/excel.table) (например, с помощью `TableRowCollection.add()`) можно столкнуться с низкой производительностью.</span><span class="sxs-lookup"><span data-stu-id="b7937-152">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="b7937-153">Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону.</span><span class="sxs-lookup"><span data-stu-id="b7937-153">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="b7937-154">Если вы пытаетесь записать данные в существующую таблицу, запишите данные в объект range с помощью `table.getDataBodyRange()`, и таблица расширится автоматически.</span><span class="sxs-lookup"><span data-stu-id="b7937-154">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="b7937-155">Ниже приведен пример такого способа.</span><span class="sxs-lookup"><span data-stu-id="b7937-155">Here is an example of this approach:</span></span>

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
> <span data-ttu-id="b7937-156">Можно легко преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="b7937-156">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="b7937-157">Прекращение отслеживания ненужных диапазонов</span><span class="sxs-lookup"><span data-stu-id="b7937-157">Untrack unneeded ranges</span></span>

<span data-ttu-id="b7937-158">Слой JavaScript создает прокси-объекты для вашей надстройки для взаимодействия с книгой Excel и базовыми диапазонами.</span><span class="sxs-lookup"><span data-stu-id="b7937-158">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="b7937-159">Эти объекты хранятся в памяти до вызова `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="b7937-159">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="b7937-160">Операции с большими пакетами могут создавать много прокси-объектов, необходимых надстройке лишь один раз, которые можно удалить из памяти до выполнения пакетных действий.</span><span class="sxs-lookup"><span data-stu-id="b7937-160">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="b7937-161">Метод [Range.untrack()](/javascript/api/excel/excel.range#untrack--) удаляет объект Excel Range из памяти.</span><span class="sxs-lookup"><span data-stu-id="b7937-161">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="b7937-162">Вызов этого метода после завершения действий надстройки с диапазоном должен приводить к заметному повышению производительности при использовании большого количества объектов Range.</span><span class="sxs-lookup"><span data-stu-id="b7937-162">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span>

> [!NOTE]
> <span data-ttu-id="b7937-163">`Range.untrack()` — это ярлык для [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span><span class="sxs-lookup"><span data-stu-id="b7937-163">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="b7937-164">Отслеживание любого прокси-объекта можно прекратить, удалив его из списка отслеживаемых объектов в контексте.</span><span class="sxs-lookup"><span data-stu-id="b7937-164">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="b7937-165">Обычно объекты Range являются единственными объектами Excel, используемыми в достаточных количествах для применения прекращения отслеживания.</span><span class="sxs-lookup"><span data-stu-id="b7937-165">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="b7937-166">В приведенном ниже примере кода выбранный диапазон заполняется данными по одной ячейке.</span><span class="sxs-lookup"><span data-stu-id="b7937-166">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="b7937-167">После добавления значения в ячейку, диапазон отображает, что отслеживание ячейки прекращено.</span><span class="sxs-lookup"><span data-stu-id="b7937-167">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="b7937-168">Выполните этот код с выбранным диапазоном от 10 000 до 20 000 ячеек сначала со строкой `cell.untrack()`, а затем без нее.</span><span class="sxs-lookup"><span data-stu-id="b7937-168">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="b7937-169">Вы должны заметить, что код выполняется с использованием строки `cell.untrack()` быстрее, чем без нее.</span><span class="sxs-lookup"><span data-stu-id="b7937-169">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="b7937-170">Вы также можете заметить уменьшение времени отклика впоследствии, так как этап очистки занимает меньше времени.</span><span class="sxs-lookup"><span data-stu-id="b7937-170">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="b7937-171">См. также</span><span class="sxs-lookup"><span data-stu-id="b7937-171">See also</span></span>

- [<span data-ttu-id="b7937-172">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="b7937-172">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b7937-173">Дополнительные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="b7937-173">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="b7937-174">Ограничения ресурсов и оптимизация производительности надстроек Office</span><span class="sxs-lookup"><span data-stu-id="b7937-174">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- [<span data-ttu-id="b7937-175">Объект Worksheet Functions (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="b7937-175">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)

---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизация Excel надстройки с помощью API JavaScript.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 5313bb3fe25d165e49cc0508e81d58294db48798
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349387"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="bb8a6-103">Оптимизация производительности с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="bb8a6-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="bb8a6-104">Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="bb8a6-105">Вы обнаружите существенные различия в производительности между разными подходами.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="bb8a6-106">В этой статье приведены инструкции и примеры кода, показывающие, как эффективно выполнять стандартные задачи, используя API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bb8a6-107">Многие проблемы производительности можно устранить с помощью рекомендуемого использования и `load` `sync` вызовов.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-107">Many performance issues can be addressed through recommended usage of `load` and `sync` calls.</span></span> <span data-ttu-id="bb8a6-108">См. раздел "Улучшения производительности с помощью API для приложений" в разделе Ограничения ресурсов и оптимизация производительности для Office надстройки для консультаций по эффективной работе с API, определенными для приложений. [](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis)</span><span class="sxs-lookup"><span data-stu-id="bb8a6-108">See the "Performance improvements with the application-specific APIs" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) for advice on working with the application-specific APIs in an efficient way.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="bb8a6-109">Временная приостановка процессов Excel</span><span class="sxs-lookup"><span data-stu-id="bb8a6-109">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="bb8a6-110">В Excel есть несколько фоновых задач, которые реагируют на ввод, выполняемый как пользователями, так и надстройкой.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-110">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="bb8a6-111">Для повышения производительности можно управлять некоторыми из этих процессов Excel.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-111">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="bb8a6-112">Это особенно полезно, если ваша надстройка работает с большими наборами данных.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-112">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="bb8a6-113">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="bb8a6-113">Suspend calculation temporarily</span></span>

<span data-ttu-id="bb8a6-114">Если вы пытаетесь выполнить операцию с большим количеством ячеек (например, установка значения огромного объекта range) и не возражаете временно приостановить расчеты в Excel до завершения операции, рекомендуется приостановить вычисление до следующего вызова `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-114">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="bb8a6-115">Дополнительные сведения об использовании API `suspendApiCalculationUntilNextSync()` для приостановки и повторного включения вычислений удобным способом см. в справочном документе [Объект Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="bb8a6-115">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="bb8a6-116">В следующем коде показано, как временно приостановить вычисление.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-116">The following code demonstrates how to suspend calculation temporarily.</span></span>

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

<span data-ttu-id="bb8a6-117">Обратите внимание, что приостановлены только расчеты формул.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-117">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="bb8a6-118">Все измененные ссылки по-прежнему перестраиваются.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-118">Any altered references are still rebuilt.</span></span> <span data-ttu-id="bb8a6-119">Например, переименование таблицы по-прежнему обновляет все ссылки в формулах на этот список.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-119">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="bb8a6-120">Приостановка обновления экрана</span><span class="sxs-lookup"><span data-stu-id="bb8a6-120">Suspend screen updating</span></span>

<span data-ttu-id="bb8a6-121">Excel отображает изменения, производимые вашей надстройкой, примерно по мере их выполнения в коде.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-121">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="bb8a6-122">Для больших циклических наборов данных может не требоваться просмотр хода выполнения на экране в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-122">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="bb8a6-123">Параметр `Application.suspendScreenUpdatingUntilNextSync()` приостанавливает визуальные обновления для Excel до вызова надстройкой метода `context.sync()` или завершения метода `Excel.run` (неявно вызывающего `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="bb8a6-123">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="bb8a6-124">Необходимо учитывать, что Excel не будет проявлять признаков работы до следующей синхронизации. Ваша надстройка должна либо предоставить пользователям инструкции, оповещающие их об этой задержке, либо отобразить строку состояния, демонстрирующую активность.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-124">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="bb8a6-125">Не звони `suspendScreenUpdatingUntilNextSync` несколько раз (например, в цикле).</span><span class="sxs-lookup"><span data-stu-id="bb8a6-125">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="bb8a6-126">Повторные вызовы при Excel окне.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-126">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="bb8a6-127">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="bb8a6-127">Enable and disable events</span></span>

<span data-ttu-id="bb8a6-128">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-128">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="bb8a6-129">Пример кода, в котором показано, как включить и отключить события, см. в статье [Работа с событиями](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="bb8a6-129">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="bb8a6-130">Импорт данных в таблицы</span><span class="sxs-lookup"><span data-stu-id="bb8a6-130">Importing data into tables</span></span>

<span data-ttu-id="bb8a6-131">При попытке импортировать огромное количество данных непосредственно в объект [Table](/javascript/api/excel/excel.table) (например, с помощью `TableRowCollection.add()`) можно столкнуться с низкой производительностью.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-131">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="bb8a6-132">Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-132">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="bb8a6-133">Если вы пытаетесь записать данные в существующую таблицу, запишите данные в объект range с помощью `table.getDataBodyRange()`, и таблица расширится автоматически.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-133">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span>

<span data-ttu-id="bb8a6-134">Ниже приведен пример такого способа.</span><span class="sxs-lookup"><span data-stu-id="bb8a6-134">Here is an example of this approach:</span></span>

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
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
> <span data-ttu-id="bb8a6-135">Можно легко преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="bb8a6-135">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="bb8a6-136">См. также</span><span class="sxs-lookup"><span data-stu-id="bb8a6-136">See also</span></span>

* [<span data-ttu-id="bb8a6-137">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="bb8a6-137">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="bb8a6-138">Ограничения ресурсов и оптимизация производительности надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bb8a6-138">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
* [<span data-ttu-id="bb8a6-139">Объект Worksheet Functions (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="bb8a6-139">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)

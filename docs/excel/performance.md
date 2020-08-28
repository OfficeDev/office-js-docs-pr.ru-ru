---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизируйте производительность надстройки Excel с помощью API JavaScript.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: fdaccdca4779aaca64420794e382330994488606
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294103"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="90c5a-103">Оптимизация производительности с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="90c5a-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="90c5a-104">Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="90c5a-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="90c5a-105">Вы обнаружите существенные различия в производительности между разными подходами.</span><span class="sxs-lookup"><span data-stu-id="90c5a-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="90c5a-106">В этой статье приведены инструкции и примеры кода, показывающие, как эффективно выполнять стандартные задачи, используя API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="90c5a-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="90c5a-107">Многие проблемы, связанные с производительностью, можно устранить, выполняя Рекомендуемые `load` `sync` вызовы и вызовы.</span><span class="sxs-lookup"><span data-stu-id="90c5a-107">Many performance issues can be addressed through recommended usage of `load` and `sync` calls.</span></span> <span data-ttu-id="90c5a-108">Изучите раздел "улучшения производительности с помощью API для определенных приложений" в разделе [пределы ресурсов и оптимизация производительности для надстроек Office](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) , чтобы получить рекомендации по работе с API, зависящими от приложения.</span><span class="sxs-lookup"><span data-stu-id="90c5a-108">See the "Performance improvements with the application-specific APIs" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) for advice on working with the application-specific APIs in an efficient way.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="90c5a-109">Временная приостановка процессов Excel</span><span class="sxs-lookup"><span data-stu-id="90c5a-109">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="90c5a-110">В Excel есть несколько фоновых задач, которые реагируют на ввод, выполняемый как пользователями, так и надстройкой.</span><span class="sxs-lookup"><span data-stu-id="90c5a-110">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="90c5a-111">Для повышения производительности можно управлять некоторыми из этих процессов Excel.</span><span class="sxs-lookup"><span data-stu-id="90c5a-111">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="90c5a-112">Это особенно полезно, если ваша надстройка работает с большими наборами данных.</span><span class="sxs-lookup"><span data-stu-id="90c5a-112">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="90c5a-113">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="90c5a-113">Suspend calculation temporarily</span></span>

<span data-ttu-id="90c5a-114">Если вы пытаетесь выполнить операцию с большим количеством ячеек (например, установка значения огромного объекта range) и не возражаете временно приостановить расчеты в Excel до завершения операции, рекомендуется приостановить вычисление до следующего вызова `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="90c5a-114">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="90c5a-115">Дополнительные сведения об использовании API `suspendApiCalculationUntilNextSync()` для приостановки и повторного включения вычислений удобным способом см. в справочном документе [Объект Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="90c5a-115">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="90c5a-116">В приведенном ниже коде показано, как временно приостановить вычисление:</span><span class="sxs-lookup"><span data-stu-id="90c5a-116">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

<span data-ttu-id="90c5a-117">Обратите внимание, что приостанавливаются только вычисления формул.</span><span class="sxs-lookup"><span data-stu-id="90c5a-117">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="90c5a-118">Все измененные ссылки все еще перестраиваются.</span><span class="sxs-lookup"><span data-stu-id="90c5a-118">Any altered references are still rebuilt.</span></span> <span data-ttu-id="90c5a-119">Например, при переименовании листа все ссылки в формулах будут обновляться на этом листе.</span><span class="sxs-lookup"><span data-stu-id="90c5a-119">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="90c5a-120">Приостановка обновления экрана</span><span class="sxs-lookup"><span data-stu-id="90c5a-120">Suspend screen updating</span></span>

<span data-ttu-id="90c5a-121">Excel отображает изменения, производимые вашей надстройкой, примерно по мере их выполнения в коде.</span><span class="sxs-lookup"><span data-stu-id="90c5a-121">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="90c5a-122">Для больших циклических наборов данных может не требоваться просмотр хода выполнения на экране в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="90c5a-122">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="90c5a-123">Параметр `Application.suspendScreenUpdatingUntilNextSync()` приостанавливает визуальные обновления для Excel до вызова надстройкой метода `context.sync()` или завершения метода `Excel.run` (неявно вызывающего `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="90c5a-123">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="90c5a-124">Необходимо учитывать, что Excel не будет проявлять признаков работы до следующей синхронизации. Ваша надстройка должна либо предоставить пользователям инструкции, оповещающие их об этой задержке, либо отобразить строку состояния, демонстрирующую активность.</span><span class="sxs-lookup"><span data-stu-id="90c5a-124">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="90c5a-125">Не вызывайте их `suspendScreenUpdatingUntilNextSync` повторно (например, в цикле).</span><span class="sxs-lookup"><span data-stu-id="90c5a-125">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="90c5a-126">Повторные вызовы приведут к мерцанию окна Excel.</span><span class="sxs-lookup"><span data-stu-id="90c5a-126">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="90c5a-127">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="90c5a-127">Enable and disable events</span></span>

<span data-ttu-id="90c5a-128">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="90c5a-128">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="90c5a-129">Пример кода, в котором показано, как включить и отключить события, см. в статье [Работа с событиями](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="90c5a-129">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="90c5a-130">Импорт данных в таблицы</span><span class="sxs-lookup"><span data-stu-id="90c5a-130">Importing data into tables</span></span>

<span data-ttu-id="90c5a-131">При попытке импортировать огромное количество данных непосредственно в объект [Table](/javascript/api/excel/excel.table) (например, с помощью `TableRowCollection.add()`) можно столкнуться с низкой производительностью.</span><span class="sxs-lookup"><span data-stu-id="90c5a-131">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="90c5a-132">Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону.</span><span class="sxs-lookup"><span data-stu-id="90c5a-132">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="90c5a-133">Если вы пытаетесь записать данные в существующую таблицу, запишите данные в объект range с помощью `table.getDataBodyRange()`, и таблица расширится автоматически.</span><span class="sxs-lookup"><span data-stu-id="90c5a-133">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span>

<span data-ttu-id="90c5a-134">Ниже приведен пример такого способа.</span><span class="sxs-lookup"><span data-stu-id="90c5a-134">Here is an example of this approach:</span></span>

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
> <span data-ttu-id="90c5a-135">Можно легко преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="90c5a-135">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="90c5a-136">См. также</span><span class="sxs-lookup"><span data-stu-id="90c5a-136">See also</span></span>

* [<span data-ttu-id="90c5a-137">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="90c5a-137">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="90c5a-138">Ограничения ресурсов и оптимизация производительности надстроек Office</span><span class="sxs-lookup"><span data-stu-id="90c5a-138">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
* [<span data-ttu-id="90c5a-139">Объект Worksheet Functions (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="90c5a-139">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)

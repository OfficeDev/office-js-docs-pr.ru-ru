---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизируйте производительность с использованием API JavaScript для Excel
ms.date: 12/06/2018
ms.openlocfilehash: f076ad6f773725c878b404d1039271a2ac59be48
ms.sourcegitcommit: 0adc31ceaba92cb15dc6430c00fe7a96c107c9de
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/09/2018
ms.locfileid: "27210093"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="219d8-103">Оптимизация производительности с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="219d8-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="219d8-104">Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="219d8-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="219d8-105">Вы обнаружите существенные различия в производительности между разными подходами.</span><span class="sxs-lookup"><span data-stu-id="219d8-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="219d8-106">В этой статье приведены инструкции и примеры кода, показывающие, как эффективно выполнять стандартные задачи, используя API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="219d8-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="219d8-107">Минимизация количества вызовов sync()</span><span class="sxs-lookup"><span data-stu-id="219d8-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="219d8-108">В API JavaScript для Excel ```sync()``` является единственной асинхронной операцией и в некоторых обстоятельствах может выполняться медленно, особенно в случае с Excel Online.</span><span class="sxs-lookup"><span data-stu-id="219d8-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="219d8-109">Для оптимизации производительности минимизируйте количество вызовов ```sync()```, поставив в очередь максимально возможное количество изменений до ее вызова.</span><span class="sxs-lookup"><span data-stu-id="219d8-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="219d8-110">Примеры кода, использующие этот подход, см. в статье [Основные концепции — sync()](excel-add-ins-core-concepts.md#sync).</span><span class="sxs-lookup"><span data-stu-id="219d8-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="219d8-111">Минимизация количества созданных прокси-объектов</span><span class="sxs-lookup"><span data-stu-id="219d8-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="219d8-112">Избегайте повторного создания одного и того же прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="219d8-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="219d8-113">Вместо этого, если вам нужен одинаковый прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.</span><span class="sxs-lookup"><span data-stu-id="219d8-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="219d8-114">Загрузка только необходимых свойств</span><span class="sxs-lookup"><span data-stu-id="219d8-114">Load necessary properties only</span></span>

<span data-ttu-id="219d8-115">В API JavaScript для Excel необходимо явно загрузить свойства прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="219d8-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="219d8-116">Несмотря на то, что вы можете загрузить все свойства одновременно, сделав пустой вызов ```load()```, этот подход может значительно замедлить производительность.</span><span class="sxs-lookup"><span data-stu-id="219d8-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="219d8-117">Вместо этого предлагается загружать только необходимые свойства, особенно для объектов с большим количеством свойств.</span><span class="sxs-lookup"><span data-stu-id="219d8-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="219d8-118">Например, если вы собираетесь считать свойство **address** объекта range, при вызове метода **load()** укажите только это свойство:</span><span class="sxs-lookup"><span data-stu-id="219d8-118">For example, if you only intend to read the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="219d8-119">Вы можете вызвать метод **load()** любым из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="219d8-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="219d8-120">_Синтаксис:_</span><span class="sxs-lookup"><span data-stu-id="219d8-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="219d8-121">_Где:_</span><span class="sxs-lookup"><span data-stu-id="219d8-121">_Where:_</span></span>
 
* <span data-ttu-id="219d8-122">`properties` — это список свойств для загрузки, указанных как строки с разделителями-запятыми или как массив имен.</span><span class="sxs-lookup"><span data-stu-id="219d8-122">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names.</span></span> <span data-ttu-id="219d8-123">Дополнительные сведения см. в описаниях методов **load()**, определенных для объектов, в [справочнике по API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="219d8-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="219d8-p106">`loadOption` указывает объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в статье, посвященной [параметрам](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) загрузки объектов.</span><span class="sxs-lookup"><span data-stu-id="219d8-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="219d8-126">Имейте в виду, что некоторые "свойства" объекта могут совпадать с именем другого объекта.</span><span class="sxs-lookup"><span data-stu-id="219d8-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="219d8-127">Например, `format` — это свойство объекта range, но также имеется и объект `format`.</span><span class="sxs-lookup"><span data-stu-id="219d8-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="219d8-128">Поэтому если вы, например, вызываете `range.load("format")`, это эквивалентно `range.format.load()`, являющемуся пустым вызовом load(), который может стать причиной проблем с производительностью, как описано ранее.</span><span class="sxs-lookup"><span data-stu-id="219d8-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="219d8-129">Чтобы избежать этого, ваш код должен загружать только "конечные узлы" в дереве объектов.</span><span class="sxs-lookup"><span data-stu-id="219d8-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="219d8-130">Временная приостановка процессов Excel</span><span class="sxs-lookup"><span data-stu-id="219d8-130">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="219d8-131">В Excel есть несколько фоновых задач, которые реагируют на ввод, выполняемый как пользователями, так и надстройкой.</span><span class="sxs-lookup"><span data-stu-id="219d8-131">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="219d8-132">Для повышения производительности можно управлять некоторыми из этих процессов Excel.</span><span class="sxs-lookup"><span data-stu-id="219d8-132">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="219d8-133">Это особенно полезно, если ваша надстройка работает с большими наборами данных.</span><span class="sxs-lookup"><span data-stu-id="219d8-133">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="219d8-134">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="219d8-134">Suspend calculation temporarily</span></span>

<span data-ttu-id="219d8-135">Если вы пытаетесь выполнить операцию с большим количеством ячеек (например, установка значения огромного объекта range) и не возражаете временно приостановить расчеты в Excel до завершения операции, рекомендуется приостановить вычисление до следующего вызова `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="219d8-135">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="219d8-136">Дополнительные сведения об использовании API `suspendApiCalculationUntilNextSync()` для приостановки и повторного включения вычислений удобным способом см. в справочном документе [Объект Application](https://docs.microsoft.com/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="219d8-136">See [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="219d8-137">В приведенном ниже коде показано, как временно приостановить вычисление:</span><span class="sxs-lookup"><span data-stu-id="219d8-137">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

### <a name="suspend-screen-updating"></a><span data-ttu-id="219d8-138">Приостановка обновления экрана</span><span class="sxs-lookup"><span data-stu-id="219d8-138">Suspend screen updating</span></span>

> [!NOTE]
> <span data-ttu-id="219d8-139">Для метода `suspendScreenUpdatingUntilNextSync()`, описанного в этой статье, требуется бета-версия библиотеки JavaScript для Office из [сети CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="219d8-139">The `suspendScreenUpdatingUntilNextSync()` method described in this article requires the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="219d8-140">[Файл определения типа] (https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) также находится в сети CDN.</span><span class="sxs-lookup"><span data-stu-id="219d8-140">The [type definition file] (https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) is also found at the CDN.</span></span> <span data-ttu-id="219d8-141">Дополнительные сведения о наших планируемых API см. на странице [открытой спецификации](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec) на портале GitHub.</span><span class="sxs-lookup"><span data-stu-id="219d8-141">For more information on our upcoming APIs, please visit the [open spec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec) on GitHub.</span></span>

<span data-ttu-id="219d8-142">Excel отображает изменения, производимые вашей надстройкой, примерно по мере их выполнения в коде.</span><span class="sxs-lookup"><span data-stu-id="219d8-142">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="219d8-143">Для больших циклических наборов данных может не требоваться просмотр хода выполнения на экране в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="219d8-143">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="219d8-144">Параметр `Application.suspendScreenUpdatingUntilNextSync()` приостанавливает визуальные обновления для Excel до вызова надстройкой метода `context.sync()` или завершения метода `Excel.run` (неявно вызывающего `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="219d8-144">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="219d8-145">Необходимо учитывать, что Excel не будет проявлять признаков работы до следующей синхронизации. Ваша надстройка должна либо предоставить пользователям инструкции, оповещающие их об этой задержке, либо отобразить строку состояния, демонстрирующую активность.</span><span class="sxs-lookup"><span data-stu-id="219d8-145">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="219d8-146">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="219d8-146">Enable and disable events</span></span>

<span data-ttu-id="219d8-147">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="219d8-147">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="219d8-148">Пример кода, в котором показано, как включить и отключить события, см. в статье [Работа с событиями](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="219d8-148">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="219d8-149">Изменение всех ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="219d8-149">Update all cells in a range</span></span> 

<span data-ttu-id="219d8-150">Если нужно изменить все ячейки в диапазоне с использованием одинакового значения или свойства, это может занять много времени при применении двумерного массива, многократно задающего одно и то же значение, поскольку в этом способе Excel требуется выполнять итерации по всем ячейкам в диапазоне для установки каждой отдельно.</span><span class="sxs-lookup"><span data-stu-id="219d8-150">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="219d8-151">В Excel есть более эффективный способ изменения всех ячеек в диапазоне с использованием одинакового значения или свойства.</span><span class="sxs-lookup"><span data-stu-id="219d8-151">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="219d8-152">Если нужно применить одинаковое значение, одинаковый числовой формат или одинаковую формулу для диапазона ячеек, эффективнее указывать одно значение вместо массива значений.</span><span class="sxs-lookup"><span data-stu-id="219d8-152">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="219d8-153">Это значительно повысит производительность.</span><span class="sxs-lookup"><span data-stu-id="219d8-153">Doing so will significantly improve performance.</span></span> <span data-ttu-id="219d8-154">Пример кода, демонстрирующий этот способ в действии, см. в статье [Основные концепции — Изменение всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="219d8-154">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="219d8-155">Распространенным сценарием применения этого способа является установка разных числовых форматов в разных столбцах на листе.</span><span class="sxs-lookup"><span data-stu-id="219d8-155">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="219d8-156">В этом случае можно просто выполнить итерацию столбцов и установить числовой формат для каждого столбца с помощью одного значения.</span><span class="sxs-lookup"><span data-stu-id="219d8-156">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="219d8-157">Обработайте каждый столбец в качестве диапазона, как показано в примере кода [Изменение всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="219d8-157">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="219d8-158">При использовании TypeScript вы заметите ошибку компиляции с сообщением, что одно значение не может быть установлено в двумерный массив.</span><span class="sxs-lookup"><span data-stu-id="219d8-158">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="219d8-159">Это неизбежно, поскольку значения *являются* двумерным массивом при извлечении свойств, а TypeScript не допускает использования разных типов методов задания и получения.</span><span class="sxs-lookup"><span data-stu-id="219d8-159">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="219d8-160">Однако есть простой обходной путь — установка значений с суффиксом `as any`, например `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="219d8-160">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="219d8-161">Импорт данных в таблицы</span><span class="sxs-lookup"><span data-stu-id="219d8-161">Importing data into tables</span></span>

<span data-ttu-id="219d8-162">При попытке импортировать огромное количество данных непосредственно в объект [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) (например, с помощью `TableRowCollection.add()`) можно столкнуться с низкой производительностью.</span><span class="sxs-lookup"><span data-stu-id="219d8-162">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="219d8-163">Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону.</span><span class="sxs-lookup"><span data-stu-id="219d8-163">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="219d8-164">Если вы пытаетесь записать данные в существующую таблицу, запишите данные в объект range с помощью `table.getDataBodyRange()`, и таблица расширится автоматически.</span><span class="sxs-lookup"><span data-stu-id="219d8-164">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="219d8-165">Ниже приведен пример такого способа.</span><span class="sxs-lookup"><span data-stu-id="219d8-165">Here is an example of this approach:</span></span>

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
> <span data-ttu-id="219d8-166">Можно легко преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="219d8-166">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="219d8-167">Прекращение отслеживания ненужных диапазонов</span><span class="sxs-lookup"><span data-stu-id="219d8-167">Untrack unneeded ranges</span></span>

<span data-ttu-id="219d8-168">Слой JavaScript создает прокси-объекты для вашей надстройки для взаимодействия с книгой Excel и базовыми диапазонами.</span><span class="sxs-lookup"><span data-stu-id="219d8-168">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="219d8-169">Эти объекты хранятся в памяти до вызова `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="219d8-169">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="219d8-170">Операции с большими пакетами могут создавать много прокси-объектов, необходимых надстройке лишь один раз, которые можно удалить из памяти до выполнения пакетных действий.</span><span class="sxs-lookup"><span data-stu-id="219d8-170">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="219d8-171">Метод [Range.untrack()](/javascript/api/excel/excel.range#untrack--) удаляет объект Excel Range из памяти.</span><span class="sxs-lookup"><span data-stu-id="219d8-171">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="219d8-172">Вызов этого метода после завершения действий надстройки с диапазоном должен приводить к заметному повышению производительности при использовании большого количества объектов Range.</span><span class="sxs-lookup"><span data-stu-id="219d8-172">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span> 

> [!NOTE]
> <span data-ttu-id="219d8-173">`Range.untrack()` — это ярлык для [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span><span class="sxs-lookup"><span data-stu-id="219d8-173">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="219d8-174">Отслеживание любого прокси-объекта можно прекратить, удалив его из списка отслеживаемых объектов в контексте.</span><span class="sxs-lookup"><span data-stu-id="219d8-174">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="219d8-175">Обычно объекты Range являются единственными объектами Excel, используемыми в достаточных количествах для применения прекращения отслеживания.</span><span class="sxs-lookup"><span data-stu-id="219d8-175">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="219d8-176">В приведенном ниже примере кода выбранный диапазон заполняется данными по одной ячейке.</span><span class="sxs-lookup"><span data-stu-id="219d8-176">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="219d8-177">После добавления значения в ячейку, диапазон отображает, что отслеживание ячейки прекращено.</span><span class="sxs-lookup"><span data-stu-id="219d8-177">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="219d8-178">Выполните этот код с выбранным диапазоном от 10 000 до 20 000 ячеек сначала со строкой `cell.untrack()`, а затем без нее.</span><span class="sxs-lookup"><span data-stu-id="219d8-178">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="219d8-179">Вы должны заметить, что код выполняется с использованием строки `cell.untrack()` быстрее, чем без нее.</span><span class="sxs-lookup"><span data-stu-id="219d8-179">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="219d8-180">Вы также можете заметить уменьшение времени отклика впоследствии, так как этап очистки занимает меньше времени.</span><span class="sxs-lookup"><span data-stu-id="219d8-180">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="219d8-181">См. также</span><span class="sxs-lookup"><span data-stu-id="219d8-181">See also</span></span>

- [<span data-ttu-id="219d8-182">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="219d8-182">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="219d8-183">Дополнительные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="219d8-183">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="219d8-184">Ограничения ресурсов и оптимизация производительности надстроек Office</span><span class="sxs-lookup"><span data-stu-id="219d8-184">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- [<span data-ttu-id="219d8-185">Открытая спецификация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="219d8-185">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="219d8-186">Объект Worksheet Functions (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="219d8-186">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)

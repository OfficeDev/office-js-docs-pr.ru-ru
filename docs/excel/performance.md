---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизация производительности с помощью API Excel JavaScript
ms.date: 03/28/2018
ms.openlocfilehash: ee1687fcb1a5db74e65f5e73994653df235b4823
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505379"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="f22e5-103">Оптимизация производительности с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="f22e5-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="f22e5-p101">Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel. Вы найдете значительные различия в производительности между различными подходами. В этой статье приведены инструкции и примеры кода, показывающие, как выполнять стандартные задачи, эффективно используя интерфейс API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="f22e5-p101">There are multiple ways that you can perform common tasks with the Excel JavaScript API. You'll find significant performance differences between various approaches. This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="f22e5-107">Минимизация количества вызовов sync()</span><span class="sxs-lookup"><span data-stu-id="f22e5-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="f22e5-p102">В API JavaScript для Excel ```sync()``` является лишь асинхронной операцией, и может выполняться медленно при некоторых обстоятельствах, особенно в случае с Excel Online. Для оптимизации производительности минимизируйте количество вызовов ```sync()```, поставив в очередь максимально возможное количество изменений еще до вызова.</span><span class="sxs-lookup"><span data-stu-id="f22e5-p102">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online. To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="f22e5-110">Для образцов кода, которые используют этот подход, см. статью [Основные понятия - sync()](excel-add-ins-core-concepts.md#sync).</span><span class="sxs-lookup"><span data-stu-id="f22e5-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="f22e5-111">Минимизация количества созданных прокси-объектов</span><span class="sxs-lookup"><span data-stu-id="f22e5-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="f22e5-p103">Избегайте повторного создания одного и того же прокси-объекта. Вместо этого, если вам нужен один и тот же прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.</span><span class="sxs-lookup"><span data-stu-id="f22e5-p103">Avoid repeatedly creating the same proxy object. Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="f22e5-114">Загрузка только необходимых свойств</span><span class="sxs-lookup"><span data-stu-id="f22e5-114">Load necessary properties only</span></span>

<span data-ttu-id="f22e5-p104">В API JavaScript для Excel необходимо явно загрузить свойства прокси-объекта. Несмотря на то, что вы можете загрузить все свойства одновременно, сделав пустой вызов ```load()```, этот подход может иметь значительные эксплуатационные издержки. Вместо этого мы предлагаем вам загружать только необходимые свойства, особенно для тех объектов, которые имеют большое количество свойств.</span><span class="sxs-lookup"><span data-stu-id="f22e5-p104">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object. Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead. Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="f22e5-118">Например, если вы собираетесь считать свойство **address** объекта range, при вызове метода **load()** укажите только это свойство:</span><span class="sxs-lookup"><span data-stu-id="f22e5-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="f22e5-119">Вы можете вызвать метод **load()** любым из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="f22e5-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="f22e5-120">_Синтаксис:_</span><span class="sxs-lookup"><span data-stu-id="f22e5-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="f22e5-121">_Где:_</span><span class="sxs-lookup"><span data-stu-id="f22e5-121">_Where:_</span></span>
 
* <span data-ttu-id="f22e5-p105">`properties` — это список свойств для загрузки,указанных как строки с разделителями-запятыми или как массив имен. Для получения дополнительных сведений см. описания методов **load()**, определенных для объектов, в [справочнике по API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="f22e5-p105">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names. For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="f22e5-p106">`loadOption` указывает объект, описывающий параметры «выбрать», «развернуть», «сверху» и «пропустить». Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption).</span><span class="sxs-lookup"><span data-stu-id="f22e5-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="f22e5-p107">Имейте в виду, что некоторые «свойства» объекта могут совпадать с именем другого объекта. Например, `format` — это свойство объекта range, но также `format` сам по себе является объектом. Итак, если вы вызываете, например, `range.load("format")`, это эквивалентно `range.format.load()`, который представляет собой пустой вызов load(), который может стать причиной проблем с производительностью, как описано ранее. Чтобы избежать этого, ваш код должен загружать только «листовые узлов» в дереве объектов.</span><span class="sxs-lookup"><span data-stu-id="f22e5-p107">Please be aware that some of the “properties” under an object may have the same name as another object. For example, `format` is a property under range object, but `format` itself is an object as well. So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously. To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="f22e5-130">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="f22e5-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="f22e5-131">Если вы пытаетесь выполнить операцию с большим количеством ячеек (например, установив значение огромного объекта range) и не возражаете временно приостановить вычисление в Excel до завершения операции, мы рекомендуем приостановить расчет до следующего вызова ```context.sync()```.</span><span class="sxs-lookup"><span data-stu-id="f22e5-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="f22e5-p108">См. справочную документацию [Объект application](https://docs.microsoft.com/javascript/api/excel/excel.application) для получения дополнительных сведений об использовании API ```suspendApiCalculationUntilNextSync()``` для приостановки и повторного включения вычислений очень удобным способом. В следующем коде показано, как временно приостановить вычисления:</span><span class="sxs-lookup"><span data-stu-id="f22e5-p108">See [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way. The following code demonstrates how to suspend calculation temporarily:</span></span>

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

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="f22e5-134">Обновление всех ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="f22e5-134">Update all cells in a range</span></span> 

<span data-ttu-id="f22e5-p109">Если вам нужно обновить все ячейки в диапазоне с одним и тем же значением или свойством, может уйти много времени на выполнение этого с помощью двумерного массива, который многократно задает одно и то же значение, поскольку для этого подхода Excel требует итерации по всем ячейкам в диапазон для установки каждой отдельно. Excel имеет более эффективный способ обновления всех ячеек в диапазоне с тем же значением или свойством.</span><span class="sxs-lookup"><span data-stu-id="f22e5-p109">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately. Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="f22e5-p110">Если вам нужно применить одно и то же значение, тот же формат номера или ту же формулу для диапазона ячеек, более эффективно указывать одно значение вместо массива значений. Это значительно повысит производительность. Для примера кода, который показывает этот подход в действии, см.статью [Основные понятия - Обновления всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="f22e5-p110">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values. Doing so will significantly improve performance. For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="f22e5-p111">Обычный сценарий, в котором вы можете применить этот подход, — это установка разных форматов чисел в разных столбцах на листе. В этом случае вы можете просто выполнить итерацию столбцов и задать формат чисел для каждого столбца с одним значением. Обработайте каждый столбец в качестве диапазона, как показано в примере кода [Обновление всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="f22e5-p111">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet. In this case, you can simply iterate through the columns and set the number format on each column with a single value. Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="f22e5-p112">При использовании TypeScript вы заметите ошибку компиляции с сообщением, что одно значение не может быть установлено в 2D-массив.  Это неизбежно, поскольку значения *являются* 2D-массивом при извлечении свойств, а TypeScript не допускает использование разных типов setter и getter.  Тем не менее, есть простой обходной путь — установка значений с суффиксом `as any`, например `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="f22e5-p112">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.  This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.  However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="f22e5-146">Импорт данных в таблицы</span><span class="sxs-lookup"><span data-stu-id="f22e5-146">Importing data into tables</span></span>

<span data-ttu-id="f22e5-p113">При попытке импортировать огромное количество данных непосредственно в объект [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) (например, с помощью `TableRowCollection.add()`) вы можете столкнуться с низкой производительностью. Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону. Если вы пытаетесь записать данные в существующую таблицу, запишите данные в объект range с помощью `table.getDataBodyRange()`, и таблица расширится автоматически.</span><span class="sxs-lookup"><span data-stu-id="f22e5-p113">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance. If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range. If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="f22e5-150">Вот пример такого подхода:</span><span class="sxs-lookup"><span data-stu-id="f22e5-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="f22e5-151">Вы можете удобно преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) .</span><span class="sxs-lookup"><span data-stu-id="f22e5-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="enable-and-disable-events"></a><span data-ttu-id="f22e5-152">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="f22e5-152">Enable and disable events</span></span>

<span data-ttu-id="f22e5-p114">Производительность надстройки можно повысить с помощью отключения событий. Пример кода, в котором показано, как включить и отключить события, см. в статье [Работа с событиями](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="f22e5-p114">Performance of an add-in may be improved by disabling events. A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="f22e5-155">См. также</span><span class="sxs-lookup"><span data-stu-id="f22e5-155">See also</span></span>

- [<span data-ttu-id="f22e5-156">Основные принципы программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="f22e5-156">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f22e5-157">Углубленные принципы программирования с использованием интерфейса API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="f22e5-157">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="f22e5-158">Открытая спецификация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="f22e5-158">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="f22e5-159">Объект Worksheet Functions (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="f22e5-159">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)

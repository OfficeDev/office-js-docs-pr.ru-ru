---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизация Excel надстройки с помощью API JavaScript.
ms.date: 08/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: ade2ac02f22c93d920174f54e6fc2efed349e3d5
ms.sourcegitcommit: e4b83d43c117225898a60391ea06465ba490f895
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/08/2021
ms.locfileid: "60809065"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Оптимизация производительности с использованием API JavaScript для Excel

Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel. Вы обнаружите существенные различия в производительности между разными подходами. В этой статье приведены инструкции и примеры кода, показывающие, как эффективно выполнять стандартные задачи, используя API JavaScript для Excel.

> [!IMPORTANT]
> Многие проблемы производительности можно устранить с помощью рекомендуемого использования и `load` `sync` вызовов. См. раздел "Улучшения производительности с помощью API для приложений" в разделе Ограничения ресурсов и оптимизация производительности для Office надстройки для консультаций по эффективной работе с API, определенными для приложений. [](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis)

## <a name="suspend-excel-processes-temporarily"></a>Временная приостановка процессов Excel

В Excel есть несколько фоновых задач, которые реагируют на ввод, выполняемый как пользователями, так и надстройкой. Для повышения производительности можно управлять некоторыми из этих процессов Excel. Это особенно полезно, если ваша надстройка работает с большими наборами данных.

### <a name="suspend-calculation-temporarily"></a>Временная приостановка вычисления

Если вы пытаетесь выполнить операцию с большим количеством ячеек (например, установка значения огромного объекта range) и не возражаете временно приостановить расчеты в Excel до завершения операции, рекомендуется приостановить вычисление до следующего вызова `context.sync()`.

Дополнительные сведения об использовании API `suspendApiCalculationUntilNextSync()` для приостановки и повторного включения вычислений удобным способом см. в справочном документе [Объект Application](/javascript/api/excel/excel.application). В следующем коде показано, как временно приостановить вычисление.

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

Обратите внимание, что приостановлены только расчеты формул. Все измененные ссылки по-прежнему перестраиваются. Например, переименование таблицы по-прежнему обновляет все ссылки в формулах на этот список.

### <a name="suspend-screen-updating"></a>Приостановка обновления экрана

Excel отображает изменения, производимые вашей надстройкой, примерно по мере их выполнения в коде. Для больших циклических наборов данных может не требоваться просмотр хода выполнения на экране в режиме реального времени. Параметр `Application.suspendScreenUpdatingUntilNextSync()` приостанавливает визуальные обновления для Excel до вызова надстройкой метода `context.sync()` или завершения метода `Excel.run` (неявно вызывающего `context.sync`). Необходимо учитывать, что Excel не будет проявлять признаков работы до следующей синхронизации. Ваша надстройка должна либо предоставить пользователям инструкции, оповещающие их об этой задержке, либо отобразить строку состояния, демонстрирующую активность.

> [!NOTE]
> Не звони `suspendScreenUpdatingUntilNextSync` несколько раз (например, в цикле). Повторные вызовы при Excel окне.

### <a name="enable-and-disable-events"></a>Включение и отключение событий

Производительность надстройки можно повысить с помощью отключения событий. Пример кода, в котором показано, как включить и отключить события, см. в статье [Работа с событиями](excel-add-ins-events.md#enable-and-disable-events).

## <a name="importing-data-into-tables"></a>Импорт данных в таблицы

При попытке импортировать огромное количество данных непосредственно в объект [Table](/javascript/api/excel/excel.table) (например, с помощью `TableRowCollection.add()`) можно столкнуться с низкой производительностью. Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону. Если вы пытаетесь записать данные в существующую таблицу, запишите данные в объект range с помощью `table.getDataBodyRange()`, и таблица расширится автоматически.

Ниже приведен пример такого способа.

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
> Можно легко преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](/javascript/api/excel/excel.table#convertToRange__).

## <a name="payload-size-limit-best-practices"></a>Оптимальные методы ограничения размера полезной нагрузки

API Excel JavaScript имеет ограничения по размеру для звонков API. Excel в Интернете имеет ограничение размера полезной нагрузки для запросов и ответов в размере 5 МБ, а API возвращает ошибку, если это `RichAPI.Error` ограничение превышено. На всех платформах диапазон для получения операций ограничен пятью миллионами ячеек. Большие диапазоны обычно превышают оба этих ограничения.

Размер полезной нагрузки запроса — это сочетание следующих трех компонентов.

* Количество вызовов API
* Количество объектов, таких как `Range` объекты
* Длина значения для набора или получения

Если API возвращает ошибку, используйте стратегии наилучшей практики, задокументированные в этой статье, чтобы оптимизировать скрипт и `RequestPayloadSizeLimitExceeded` избежать ошибки.

### <a name="strategy-1-move-unchanged-values-out-of-loops"></a>Стратегия 1. Перемещение неизмененных значений из циклов

Ограничить количество процессов, которые происходят в циклах, чтобы повысить производительность. В следующем примере кода можно выйти из цикла, так как он не меняется `context.workbook.worksheets.getActiveWorksheet()` `for` в этом цикле.

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    var ranges = [];
    
    // This sample retrieves the worksheet every time the loop runs, which is bad for performance.
    for (let i = 0; i < 7500; i++) {
      var rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

В следующем примере кода показана логика, аналогичная предыдущему примеру кода, но с улучшенной стратегией производительности. Значение `context.workbook.worksheets.getActiveWorksheet()` извлекается перед циклом, так как это значение не требуется извлекать при каждом запуске `for` `for` цикла. В этом цикле должны быть извлечены только значения, которые изменяются в контексте цикла.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    var ranges = [];
    // Retrieve the worksheet outside the loop.
    var worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    // Only process the necessary values inside the loop.
    for (let i = 0; i < 7500; i++) {
      var rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### <a name="strategy-2-create-fewer-range-objects"></a>Стратегия 2. Создание меньше объектов диапазона

Создайте меньше объектов диапазона, чтобы повысить производительность и свести к минимуму размер полезной нагрузки. Два подхода к созданию меньшего диапазона объектов описаны в следующих разделах статьи и примерах кода.

#### <a name="split-each-range-array-into-multiple-arrays"></a>Разделение каждого массива диапазона на несколько массивов

Один из способов создания меньшего диапазона объектов — разделить каждый массив диапазона на несколько массивов, а затем обработать каждый новый массив циклом и новым `context.sync()` вызовом.

> [!IMPORTANT]
> Используйте эту стратегию только в том случае, если вы впервые определили превышение предельного размера запроса полезной нагрузки. Использование нескольких циклов может уменьшить размер каждого запроса полезной нагрузки, чтобы избежать превышения лимита в 5 МБ, но использование нескольких циклов и нескольких вызовов также отрицательно влияет на `context.sync()` производительность.

В следующем примере кода предпринимается попытка обработки большого массива диапазонов в одном цикле, а затем одного `context.sync()` вызова. Обработка слишком многких значений диапазона в одном вызове приводит к превышению размера запроса полезной нагрузки `context.sync()` на 5МБ.

```js
// This code sample does not show a recommended strategy.
// Calling 10,000 rows would likely exceed the 5MB payload size limit in a real-world situation.
async function run() {
  await Excel.run(async (context) => {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // This sample attempts to process too many ranges at once. 
    for (let row = 1; row < 10000; row++) {
      var range = sheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

В следующем примере кода показана логика, аналогичная предыдущему примеру кода, но стратегия, которая позволяет избежать превышения предельного размера запроса полезной нагрузки в 5 МБ. В следующем примере кода диапазоны обрабатываются в двух отдельных циклах, и за каждым циклом следует `context.sync()` вызов.

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops and `context.sync()` calls negatively impacts performance.
// Only use this strategy if you've determined that you're exceeding the payload request limit.
async function run() {
  await Excel.run(async (context) => {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();

    // Split the ranges into two loops, rows 1-5000 and then 5001-10000.
    for (let row = 1; row < 5000; row++) {
      var range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    // Sync after each loop. 
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
      var range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### <a name="set-range-values-in-an-array"></a>Набор значений диапазона в массиве

Другой способ создать меньше объектов диапазона — создать массив, использовать цикл, чтобы задать все данные в этом массиве, а затем передать значения массива в диапазон. Это дает преимущества как для производительности, так и для размера полезной нагрузки. Вместо вызова `range.values` для каждого диапазона в цикле, называется один раз за `range.values` пределами цикла.

В следующем примере кода показано, как создать массив, задайте значения этого массива в цикле, а затем передайте значения массива в диапазон за `for` пределами цикла.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    // Create an array.
    const array = new Array(10000);

    // Set the values of the array inside the loop.
    for (var i = 0; i < 10000; i++) {
      array[i] = [1];
    }

    // Pass the array values to a range outside the loop. 
    var range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## <a name="see-also"></a>См. также

* [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
* [Обработка ошибок с Excel API JavaScript](excel-add-ins-error-handling.md)
* [Ограничения ресурсов и оптимизация производительности надстроек Office](../concepts/resource-limits-and-performance-optimization.md)
* [Объект Worksheet Functions (API JavaScript для Excel)](/javascript/api/excel/excel.functions)

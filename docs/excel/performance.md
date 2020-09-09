---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизируйте производительность надстройки Excel с помощью API JavaScript.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 42ab5f28717f0f7dcd06461840de692a5daf60ce
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408616"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Оптимизация производительности с использованием API JavaScript для Excel

Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel. Вы обнаружите существенные различия в производительности между разными подходами. В этой статье приведены инструкции и примеры кода, показывающие, как эффективно выполнять стандартные задачи, используя API JavaScript для Excel.

> [!IMPORTANT]
> Многие проблемы, связанные с производительностью, можно устранить, выполняя Рекомендуемые `load` `sync` вызовы и вызовы. Изучите раздел "улучшения производительности с помощью API для определенных приложений" в разделе [пределы ресурсов и оптимизация производительности для надстроек Office](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) , чтобы получить рекомендации по работе с API, зависящими от приложения.

## <a name="suspend-excel-processes-temporarily"></a>Временная приостановка процессов Excel

В Excel есть несколько фоновых задач, которые реагируют на ввод, выполняемый как пользователями, так и надстройкой. Для повышения производительности можно управлять некоторыми из этих процессов Excel. Это особенно полезно, если ваша надстройка работает с большими наборами данных.

### <a name="suspend-calculation-temporarily"></a>Временная приостановка вычисления

Если вы пытаетесь выполнить операцию с большим количеством ячеек (например, установка значения огромного объекта range) и не возражаете временно приостановить расчеты в Excel до завершения операции, рекомендуется приостановить вычисление до следующего вызова `context.sync()`.

Дополнительные сведения об использовании API `suspendApiCalculationUntilNextSync()` для приостановки и повторного включения вычислений удобным способом см. в справочном документе [Объект Application](/javascript/api/excel/excel.application). В приведенном ниже коде показано, как временно приостановить вычисление:

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

Обратите внимание, что приостанавливаются только вычисления формул. Все измененные ссылки все еще перестраиваются. Например, при переименовании листа все ссылки в формулах будут обновляться на этом листе.

### <a name="suspend-screen-updating"></a>Приостановка обновления экрана

Excel отображает изменения, производимые вашей надстройкой, примерно по мере их выполнения в коде. Для больших циклических наборов данных может не требоваться просмотр хода выполнения на экране в режиме реального времени. Параметр `Application.suspendScreenUpdatingUntilNextSync()` приостанавливает визуальные обновления для Excel до вызова надстройкой метода `context.sync()` или завершения метода `Excel.run` (неявно вызывающего `context.sync`). Необходимо учитывать, что Excel не будет проявлять признаков работы до следующей синхронизации. Ваша надстройка должна либо предоставить пользователям инструкции, оповещающие их об этой задержке, либо отобразить строку состояния, демонстрирующую активность.

> [!NOTE]
> Не вызывайте их `suspendScreenUpdatingUntilNextSync` повторно (например, в цикле). Повторные вызовы приведут к мерцанию окна Excel.

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
> Можно легко преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).

## <a name="see-also"></a>См. также

* [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
* [Ограничения ресурсов и оптимизация производительности надстроек Office](../concepts/resource-limits-and-performance-optimization.md)
* [Объект Worksheet Functions (API JavaScript для Excel)](/javascript/api/excel/excel.functions)

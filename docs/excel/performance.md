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
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Оптимизация производительности с использованием API Excel JavaScript

Существует несколько способов выполнения общих задач с помощью API Excel JavaScript. Вы найдете значительные различия в производительности между различными подходами. В этой статье приведены примеры руководств и кода, чтобы показать вам, как эффективно выполнять общие задачи с помощью API Excel JavaScript.

## <a name="minimize-the-number-of-sync-calls"></a>Минимизировать количество вызовов sync ()

В Excel JavaScript API, ```sync()``` является единственной асинхронной операцией, и в некоторых случаях она может быть медленной, особенно для Excel Online. Чтобы оптимизировать производительность, минимизируйте количество вызовов ```sync()```, поставив в очередь столько изменений, сколько возможно, прежде чем вызвать его.

См.статью [Основные понятия - синхронизация ()](excel-add-ins-core-concepts.md#sync) для образцов кода, которые следуют этой практике.

## <a name="minimize-the-number-of-proxy-objects-created"></a>Минимизировать количество созданных прокси-объектов

Избегайте повторного создания одного и того же прокси-объекта. Вместо этого, если вам нужен один и тот же прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.

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

## <a name="load-necessary-properties-only"></a>Загрузка только необходимых свойств

В Excel JavaScript API вам необходимо явно загрузить свойства прокси-объекта. Хотя вы можете сразу загрузить все свойства с пустым ```load()``` вызовом, этот подход может иметь значительные накладные расходы. Вместо этого мы предлагаем вам загружать только необходимые свойства, особенно для тех объектов, которые имеют большое количество свойств.

Например, если вы собираетесь считать свойство **address** объекта range, при вызове метода **load()** укажите только это свойство:
 
```js
range.load('address');
```
 
Вы можете вызвать метод **load()** любым из следующих способов:
 
_Синтаксис:_
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
_Где:_
 
* `properties` Это список свойств для загрузки, указанных как строки с разделителями-запятыми или как массив имен. Дополнительные сведения см. в описаниях методов **load()**, определенных для объектов в [справочнике по API JavaScript для Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).
* `loadOption` указывает объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](https://dev.office.com/reference/add-ins/excel/loadoption).

Помните, что некоторые «свойства» под объектом могут иметь то же имя, что и другой объект. Например, `format` является свойством объекта диапазона, но `format` сам по себе является объектом. Итак, если вы вызываете например, `range.load("format")`, это эквивалентно `range.format.load()`, который представляет собой вызов пустой нагрузки (), который может вызвать проблемы с производительностью, как описано ранее. Чтобы избежать  этого, ваш код должен загружать только «листовые узлы» в представлении объектов. 

## <a name="suspend-calculation-temporarily"></a>Временно приостанавливать расчет

Если вы пытаетесь выполнить операцию на большом количестве ячеек (например, установив значение огромного объекта диапазона), и вы не возражаете временно приостановить вычисление в Excel во время завершения операции, мы рекомендуем приостановить расчет до следующего ```context.sync()``` вызова.

См.статью [Объект приложения](https://dev.office.com/reference/add-ins/excel/application), справочную документацию для получения информации о том, как использовать ```suspendApiCalculationUntilNextSync()``` API для приостановки и повторного включения вычислений очень удобным способом. Следующий код демонстрирует, как временно приостановить расчет:

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

## <a name="update-all-cells-in-a-range"></a>Изменение всех ячеек в диапазоне 

Когда вам нужно обновить все ячейки в диапазоне с одним и тем же значением или свойством, может уйти много времени на выполнение этого с помощью двумерного массива, который многократно задает одно и то же значение, поскольку для этого подхода Excel требует итерации по всем ячейкам в диапазон для установки каждой отдельно. Excel имеет более эффективный способ обновления всех ячеек в диапазоне с тем же значением или свойством.

Если вам нужно применить одно и то же значение, тот же формат номера или ту же формулу для диапазона ячеек, более эффективно указывать одно значение вместо массива значений. Это значительно улучшит производительность. Для примера кода, который показывает этот подход в действии, см.статью [Основные понятия -Обновление всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Обычный сценарий, в котором вы можете применить этот подход, - это установка разных форматов чисел в разных столбцах на листе. В этом случае вы можете просто выполнить итерацию столбцов и задавать формат чисел для каждого столбца с одним значением. Обрабатывайте каждый столбец как диапазон, как показано в [Обновление всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) образца кода.

> [!NOTE]
> Если вы используете TypeScript, вы заметите ошибку компиляции, заявив, что одно значение не может быть установлено в 2D-массив.  Это неизбежно, поскольку значения *находятся* 2D-массив при извлечении свойств, а TypeScript не допускает использование разных типов setter vs getter.  Однако простым обходным путем является установление значений с помощью суффикса`as any`, например, `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Импорт данных в таблицы

При попытке импортировать огромное количество данных непосредственно в [Таблицу](https://dev.office.com/reference/add-ins/excel/table) объекта (например, используя `TableRowCollection.add()`), вы можете столкнуться с низкой производительностью. Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону. Если вы пытаетесь записать данные в существующую таблицу, напишите данные в объект диапазона через `table.getDataBodyRange()`, и таблица будет автоматически расшириться. 

Вот пример такого подхода:

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
> Вы можете удобно преобразовать объект Table в объект Range, используя [метод Table.convertToRange ()](https://dev.office.com/reference/add-ins/excel/table#converttorange) .

## <a name="see-also"></a>См. также

- [Основные понятия API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Сложные понятия, связанные с API JavaScript для Excel](excel-add-ins-advanced-concepts.md)
- [Открытая спецификация по API JavaScript для Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Объект Worksheet Functions (API JavaScript для Excel)](https://dev.office.com/reference/add-ins/excel/functions)

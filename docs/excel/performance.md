---
title: Оптимизация производительности API JavaScript для Excel
description: Оптимизируйте производительность с использованием API JavaScript для Excel
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d041356129ad5e5db8c990daaafee4e583de1dfa
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325054"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Оптимизация производительности с использованием API JavaScript для Excel

Существует несколько способов выполнения стандартных задач с помощью API JavaScript для Excel. Вы обнаружите существенные различия в производительности между разными подходами. В этой статье приведены инструкции и примеры кода, показывающие, как эффективно выполнять стандартные задачи, используя API JavaScript для Excel.

## <a name="minimize-the-number-of-sync-calls"></a>Минимизация количества вызовов sync()

В API JavaScript для Excel ```sync()``` является единственной асинхронной операцией и в некоторых обстоятельствах может выполняться медленно, особенно в случае с Excel в Интернете. Для оптимизации производительности минимизируйте количество вызовов ```sync()```, поставив в очередь максимально возможное количество изменений до ее вызова.

Примеры кода, использующие этот подход, см. в статье [Основные концепции — sync()](excel-add-ins-core-concepts.md#sync).

## <a name="minimize-the-number-of-proxy-objects-created"></a>Минимизация количества созданных прокси-объектов

Избегайте повторного создания одного и того же прокси-объекта. Вместо этого, если вам нужен одинаковый прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.

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

## <a name="load-necessary-properties-only"></a>Загрузка только необходимых свойств

В API JavaScript для Excel необходимо явно загрузить свойства прокси-объекта. Несмотря на то, что вы можете загрузить все свойства одновременно, сделав пустой вызов ```load()```, этот подход может значительно замедлить производительность. Вместо этого предлагается загружать только необходимые свойства, особенно для объектов с большим количеством свойств.

Например, если требуется только прочитать `address` свойство объекта Range, при вызове `load()` метода укажите только это свойство:

```js
range.load('address');
```

Метод можно вызвать `load()` одним из следующих способов:

_Синтаксис:_

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

_Где:_

* `properties` — это список свойств для загрузки, указанных как строки с разделителями-запятыми или как массив имен. Дополнительные сведения приведены в статье методы `load()` , определенные для объектов в [справочнике по API JavaScript для Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).
* `loadOption` указывает объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в статье, посвященной [параметрам](/javascript/api/office/officeextension.loadoption) загрузки объектов.

Имейте в виду, что некоторые "свойства" объекта могут совпадать с именем другого объекта. Например, `format` — это свойство объекта range, но также имеется и объект `format`. Поэтому если вы, например, вызываете `range.load("format")`, это эквивалентно `range.format.load()`, являющемуся пустым вызовом load(), который может стать причиной проблем с производительностью, как описано ранее. Чтобы избежать этого, ваш код должен загружать только "конечные узлы" в дереве объектов. 

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

### <a name="suspend-screen-updating"></a>Приостановка обновления экрана

Excel отображает изменения, производимые вашей надстройкой, примерно по мере их выполнения в коде. Для больших циклических наборов данных может не требоваться просмотр хода выполнения на экране в режиме реального времени. Параметр `Application.suspendScreenUpdatingUntilNextSync()` приостанавливает визуальные обновления для Excel до вызова надстройкой метода `context.sync()` или завершения метода `Excel.run` (неявно вызывающего `context.sync`). Необходимо учитывать, что Excel не будет проявлять признаков работы до следующей синхронизации. Ваша надстройка должна либо предоставить пользователям инструкции, оповещающие их об этой задержке, либо отобразить строку состояния, демонстрирующую активность.

### <a name="enable-and-disable-events"></a>Включение и отключение событий

Производительность надстройки можно повысить с помощью отключения событий. Пример кода, в котором показано, как включить и отключить события, см. в статье [Работа с событиями](excel-add-ins-events.md#enable-and-disable-events).

## <a name="update-all-cells-in-a-range"></a>Изменение всех ячеек в диапазоне

Если нужно изменить все ячейки в диапазоне с использованием одинакового значения или свойства, это может занять много времени при применении двумерного массива, многократно задающего одно и то же значение, поскольку в этом способе Excel требуется выполнять итерации по всем ячейкам в диапазоне для установки каждой отдельно. В Excel есть более эффективный способ изменения всех ячеек в диапазоне с использованием одинакового значения или свойства.

Если нужно применить одинаковое значение, одинаковый числовой формат или одинаковую формулу для диапазона ячеек, эффективнее указывать одно значение вместо массива значений. Это значительно повысит производительность. Пример кода, демонстрирующий этот способ в действии, см. в статье [Основные концепции — Изменение всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Распространенным сценарием применения этого способа является установка разных числовых форматов в разных столбцах на листе. В этом случае можно просто выполнить итерацию столбцов и установить числовой формат для каждого столбца с помощью одного значения. Обработайте каждый столбец в качестве диапазона, как показано в примере кода [Изменение всех ячеек в диапазоне](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

> [!NOTE]
> При использовании TypeScript вы заметите ошибку компиляции с сообщением, что одно значение не может быть установлено в двумерный массив.  Это неизбежно, поскольку значения *являются* двумерным массивом при извлечении свойств, а TypeScript не допускает использования разных типов методов задания и получения.  Однако есть простой обходной путь — установка значений с суффиксом `as any`, например `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Импорт данных в таблицы

При попытке импортировать огромное количество данных непосредственно в объект [Table](/javascript/api/excel/excel.table) (например, с помощью `TableRowCollection.add()`) можно столкнуться с низкой производительностью. Если вы пытаетесь добавить новую таблицу, сначала необходимо заполнить данные, установив `range.values`, а затем выполнить вызов `worksheet.tables.add()` для создания таблицы по диапазону. Если вы пытаетесь записать данные в существующую таблицу, запишите данные в объект range с помощью `table.getDataBodyRange()`, и таблица расширится автоматически. 

Ниже приведен пример такого способа.

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
> Можно легко преобразовать объект Table в объект Range, используя метод [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).

## <a name="untrack-unneeded-ranges"></a>Прекращение отслеживания ненужных диапазонов

Слой JavaScript создает прокси-объекты для вашей надстройки для взаимодействия с книгой Excel и базовыми диапазонами. Эти объекты хранятся в памяти до вызова `context.sync()`. Операции с большими пакетами могут создавать много прокси-объектов, необходимых надстройке лишь один раз, которые можно удалить из памяти до выполнения пакетных действий.

Метод [Range.untrack()](/javascript/api/excel/excel.range#untrack--) удаляет объект Excel Range из памяти. Вызов этого метода после завершения действий надстройки с диапазоном должен приводить к заметному повышению производительности при использовании большого количества объектов Range.

> [!NOTE]
> `Range.untrack()` — это ярлык для [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-). Отслеживание любого прокси-объекта можно прекратить, удалив его из списка отслеживаемых объектов в контексте. Обычно объекты Range являются единственными объектами Excel, используемыми в достаточных количествах для применения прекращения отслеживания.

В приведенном ниже примере кода выбранный диапазон заполняется данными по одной ячейке. После добавления значения в ячейку, диапазон отображает, что отслеживание ячейки прекращено. Выполните этот код с выбранным диапазоном от 10 000 до 20 000 ячеек сначала со строкой `cell.untrack()`, а затем без нее. Вы должны заметить, что код выполняется с использованием строки `cell.untrack()` быстрее, чем без нее. Вы также можете заметить уменьшение времени отклика впоследствии, так как этап очистки занимает меньше времени.

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

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Дополнительные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-advanced-concepts.md)
- [Ограничения ресурсов и оптимизация производительности надстроек Office](../concepts/resource-limits-and-performance-optimization.md)
- [Открытая спецификация по API JavaScript для Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Объект Worksheet Functions (API JavaScript для Excel)](/javascript/api/excel/excel.functions)

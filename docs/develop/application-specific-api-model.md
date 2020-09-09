---
title: Использование модели API для определенных приложений
description: Сведения о модели API на основе обещаний для Excel, OneNote и надстроек Word.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: fb25201174dcd97b40ccf6be69b238951103db07
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408602"
---
# <a name="using-the-application-specific-api-model"></a>Использование модели API для определенных приложений

В этой статье описывается, как использовать модель API для создания надстроек в Excel, Word и OneNote. В нем представлены основные концепции использования API на основе Promise.

> [!NOTE]
> Эта модель не поддерживается клиентами Office 2013. Используйте [общую модель API](office-javascript-api-object-model.md) для работы с этими версиями Office. Чтобы ознакомиться с полными сведениями о доступности платформы, ознакомьтесь с разделом [клиентские приложения и платформы Office для надстроек Office](../overview/office-add-in-availability.md).

> [!TIP]
> В примерах на этой странице используются API JavaScript для Excel, но эти понятия также относятся к API-интерфейсам OneNote, Visio и Word JavaScript.

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>Асинхронная природа интерфейсов API на основе обещаний

Надстройки Office — это веб-сайты, которые отображаются внутри контейнера браузера в приложениях Office, таких как Excel. Этот контейнер внедряется в приложение Office на платформах на настольных компьютерах, таких как Office в Windows, и запускается в элементе iFrame HTML в Office в Интернете. Из-за соображений производительности интерфейсы API Office.js не могут синхронно взаимодействовать с приложениями Office на всех платформах. Таким образом, `sync()` вызов API в Office.js возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) , которое разрешается при выполнении приложением Office запрошенных действий чтения или записи. Кроме того, можно поставить в очередь несколько действий, таких как установка свойств или вызов методов, и запускать их как пакет команд с одним вызовом `sync()` , а не отправлять отдельный запрос для каждого действия. В следующих разделах описано, как это сделать с помощью `run()` `sync()` API-интерфейсов.

## <a name="run-function"></a>функция *. Run

`Excel.run`, `Word.run` и `OneNote.run` выполните функцию, которая определяет действия, выполняемые с помощью Excel, Word и OneNote. `*.run` автоматически создает контекст запроса, который можно использовать для взаимодействия с объектами Office. По `*.run` завершении обещание разрешается, и все объекты, которые были выделены во время выполнения, автоматически освобождаются.

В приведенном ниже примере показано, как использовать `Excel.run` . Такой же шаблон также используется с Word и OneNote.

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a>Контекст запроса

Приложение Office и надстройка запускаются в двух различных процессах. Так как они используют разные среды выполнения, надстройкам требуется `RequestContext` объект, чтобы подключить надстройку к объектам в Office, таким как листы, диапазоны, абзацы и таблицы. Этот `RequestContext` объект предоставляется в качестве аргумента при вызове `*.run` .

## <a name="proxy-objects"></a>Прокси-объекты

Объекты JavaScript для Office, объявляемые и используемые с помощью API на основе Promise, являются прокси-объектами. Все методы, которые вы вызываете, либо свойства, которые вы настраиваете либо загружаете, в прокси-объектах просто добавляются в очередь команд, ожидающих выполнения. При вызове `sync()` метода в контексте запроса (например, `context.sync()` ) команды, поставленные в очередь, отправляются в приложение Office и запускаются. Эти API основаны на пакетной основе. Вы можете поместить в очередь любое количество изменений, которое требуется в контексте запроса, а затем вызвать `sync()` метод для запуска пакета команд в очереди.

Например, в приведенном ниже фрагменте кода объявляется локальный объект JavaScript [Excel. Range](/javascript/api/excel/excel.range) , `selectedRange` для ссылки на выбранный диапазон в книге Excel, а затем задаются некоторые свойства этого объекта. `selectedRange`Объект является прокси-объектом, поэтому заданные свойства и метод, вызываемый для этого объекта, не будут отражены в документе Excel до вызова надстройки `context.sync()` .

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>Совет по производительности: Минимизируйте число созданных прокси-объектов

Избегайте повторного создания одного и того же прокси-объекта. Вместо этого, если вам нужен одинаковый прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
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

### <a name="sync"></a>sync()

При вызове `sync()` метода в контексте запроса выполняется синхронизация состояния между объектами прокси-сервера и объектами в документе Office. `sync()`Метод выполняет все команды, помещенные в очередь в контексте запроса, и получает значения для всех свойств, которые должны быть загружены в прокси-объекты. `sync()`Метод выполняется асинхронно и возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается по `sync()` завершении метода.

В следующем примере показана Пакетная функция, которая определяет локальный прокси-сервер JavaScript ( `selectedRange` ), загружает свойство этого объекта, а затем использует шаблон JavaScript для синхронизации для `context.sync()` синхронизации состояния между прокси-объектами и объектами в документе Excel.

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

В предыдущем примере `selectedRange` установлен, и его параметр `address` загружается при вызове `context.sync()`.

Так как `sync()` это асинхронная операция, всегда следует возвращать `Promise` объект, чтобы убедиться, что `sync()` операция завершается, прежде чем продолжить выполнение скрипта. Если вы используете TypeScript или ES6 + JavaScript, вы можете `await` `context.sync()` позвонить вместо возврата обещаний.

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>Совет по производительности: Минимизируйте число вызовов синхронизации

В API JavaScript для Excel `sync()` является единственной асинхронной операцией и в некоторых обстоятельствах может выполняться медленно, особенно в случае с Excel в Интернете. Для оптимизации производительности минимизируйте количество вызовов `sync()`, поставив в очередь максимально возможное количество изменений до ее вызова. Чтобы получить дополнительные сведения о оптимизации производительности с помощью `sync()` , [не используйте метод Context. Sync в циклах](../concepts/correlated-objects-pattern.md).

### <a name="load"></a>load()

Перед чтением свойств прокси-объекта необходимо явно загрузить свойства для заполнения прокси-объекта данными из документа Office и затем вызвать метод `context.sync()` . Например, если вы создаете прокси-объект для ссылки на выбранный диапазон, а затем хотите прочитать свойство выбранного диапазона, необходимо `address` загрузить `address` свойство, прежде чем его можно будет прочитать. Чтобы запросить свойства прокси-объекта, вызовите `load()` метод для объекта и укажите свойства для загрузки. В следующем примере показано `Range.address` свойство, для которого выполняется загрузка `myRange` .

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> Если вы вызываете только методы или задаете свойства прокси-объекта, вам не нужно вызывать `load()` метод. `load()`Метод требуется только в том случае, если необходимо прочитать свойства прокси-объекта.

Аналогично запросам для задания свойств или вызова методов в прокси-объектах, запросы на загрузку свойств в прокси-объектах добавляются в очередь команд, ожидающих выполнения, в контексте запроса, который будет запущен, когда вы в следующий раз вызовете метод `sync()`. В очередь можно поставить сколько угодно вызовов `load()` в контексте запроса.

#### <a name="scalar-and-navigation-properties"></a>Скалярные и навигационные свойства

Существует две категории свойств: **скалярные** и **навигационные**. К скалярным свойствам относятся назначаемые типы, такие как строки, целые числа и структуры JSON. Свойства навигации — это объекты, доступные только для чтения, и коллекции объектов, которым назначены поля, а не непосредственное назначение свойства. Например, `name` `position` элементы в объекте [Excel. лист](/javascript/api/excel/excel.worksheet) являются скалярными свойствами, в то время как `protection` `tables` Свойства навигации.

Надстройка может использовать свойства навигации в качестве пути для загрузки определенных скалярных свойств. Приведенный ниже код ставит в очередь `load` команду для имени шрифта `Excel.Range` , используемого объектом, без загрузки каких бы то ни было других сведений.

```js
someRange.load("format/font/name")
```

Кроме того, можно задать скалярные свойства свойства навигации, обходим путь. Например, можно задать размер шрифта для элемента с помощью параметра `Excel.Range` `someRange.format.font.size = 10;` . Вам не нужно загружать свойство перед его заданием.

Обратите внимание, что некоторые свойства объекта могут иметь то же имя, что и другой объект. Например, `format` является свойством `Excel.Range` объекта, но `format` само по себе также является объектом. Таким образом, при совершении такого вызова, как `range.load("format")` , это эквивалентно `range.format.load()` (нежелательный пустой `load()` оператор). Чтобы избежать этого, код должен загружать только "конечные узлы" в дереве объектов.

#### <a name="calling-load-without-parameters-not-recommended"></a>Вызов `load` без параметров (не рекомендуется)

При вызове `load()` метода для объекта (или коллекции) без указания каких-либо параметров будут загружены все скалярные свойства объекта или коллекции. Загрузка ненужных данных приведет к снижению производительности надстройки. Всегда следует явно указывать свойства для загрузки.

> [!IMPORTANT]
> Объем данных, возвращаемых оператором `load` без параметров, может превышать ограничения по размерам для службы. Чтобы сократить риски для старых надстроек, некоторые свойства не возвращаются методом `load` без их явного запроса. Следующие свойства исключены из таких операций загрузки:
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

Методы в API на основе обещания, возвращающие примитивные типы, имеют похожий шаблон для `load` / `sync` парадигмы. Например, `Excel.TableCollection.getCount` получает количество таблиц в коллекции. `getCount` Возвращает значение `ClientResult<number>` , означающее, что `value` возвращаемое свойство [`ClientResult`](/javascript/api/office/officeextension.clientresult) является числом. Скрипт не может получить доступ к этому значению, пока не вызовет `context.sync()`.

Приведенный ниже код получает общее количество таблиц в книге Excel и записывает их в консоль.

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a>set()

Установка свойств объекта с вложенными свойствами навигации может быть трудоемкой задачей. В качестве альтернативы для установки отдельных свойств с помощью путей навигации, описанных выше, можно использовать `object.set()` метод, доступный для объектов в API JavaScript на основе Promise. С помощью этого метода можно задать сразу несколько свойств объекта, передавая другой объект того же типа Office.js или объект JavaScript со свойствами, сходными по структуре со свойствами объекта, для которого вызывается метод.

В приведенном ниже примере кода показано, как задать несколько свойств формата диапазона, вызвав метод `set()` и передав в него объект JavaScript, имена и типы свойств которого повторяют структуру свойств объекта `Range`. В этом примере предполагается, что данные находятся в диапазоне **B2:E2**.

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="some-properties-cannot-be-set-directly"></a>Некоторые свойства невозможно задать напрямую

Некоторые свойства не могут быть заданы, несмотря на то, что они доступны для записи. Эти свойства являются частью родительского свойства, которое должно быть задано как один объект. Это связано с тем, что родительское свойство использует вложенные свойства с определенными логическими связями. Эти родительские свойства должны быть заданы с помощью нотации литерала объекта, чтобы задать весь объект, а не задавать отдельные вложенные свойства этого объекта. Один из примеров этого примера находится в файле [PageLayout](/javascript/api/excel/excel.pagelayout). `zoom`Свойство должно быть задано с помощью одного объекта [пажелайаутзумоптионс](/javascript/api/excel/excel.pagelayoutzoomoptions) , как показано ниже:

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

В предыдущем примере вы ***не*** сможете напрямую присвоить `zoom` значение: `sheet.pageLayout.zoom.scale = 200;` . Этот оператор выдает ошибку, так как `zoom` не загружен. Даже если `zoom` были загружены, набор масштабов не вступит в силу. Все операции контекста выполняются `zoom` , обновляя прокси-объект в надстройке и перезаписывая локально заданные значения.

Это поведение отличается от [свойств навигации](application-specific-api-model.md#scalar-and-navigation-properties) , таких как [Range. Format](/javascript/api/excel/excel.range#format). Свойства `format` можно задать с помощью навигации по объектам, как показано ниже:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Можно определить свойство, для которого не могут быть заданы вложенные свойства, путем проверки модификатора только для чтения. Все свойства, доступные только для чтения, могут иметь нередактируемые вложенные свойства, не предназначенные только для чтения. Записываемые свойства, такие как `PageLayout.zoom` , должны быть заданы на уровне объекта. В сводке:

- Свойство только для чтения: вложенные свойства можно задать с помощью навигации.
- Записываемое свойство: подсвойства невозможно задать с помощью навигации (необходимо задать в качестве части исходного назначения родительского объекта).



## <a name="42ornullobject-methods-and-properties"></a>&#42;методы и свойства Орнуллобжект

Некоторые методы и свойства метода доступа создают исключение, если нужный объект не существует. Например, если вы попытаетесь получить лист Excel, указав имя листа, которого нет в книге, `getItem()` метод создаст `ItemNotFound` исключение. Библиотеки, зависящие от приложения, позволяют коду проверять наличие сущностей документа, не требуя кода обработки исключений. Это достигается с помощью `*OrNullObject` вариантов методов и свойств. Эти варианты возвращают объект, `isNullObject` свойству которого присвоено значение `true` , если указанный элемент не существует, а не создает исключение.

Например, вы можете вызвать `getItemOrNullObject()` метод для коллекции, например, с помощью **листов** , чтобы получить элемент из коллекции. `getItemOrNullObject()`Метод возвращает указанный элемент, если он существует; в противном случае возвращает объект, `isNullObject` свойству которого присвоено значение `true` . Затем код может оценить это свойство, чтобы определить, существует ли объект.

> [!NOTE]
> `*OrNullObject`Варианты никогда не возвращают значение JavaScript `null` . Они возвращают обычные прокси-объекты Office. Если объект, который представляет объект, не существует, то `isNullObject` для свойства объекта задано значение `true` . Не проверяйте возвращаемый объект на значение null или фалсити. Он никогда `null` `false` или `undefined` .

Следующий пример кода пытается извлечь лист Excel с именем "Data" с помощью `getItemOrNullObject()` метода. Если лист с таким именем не существует, создается новый лист. Обратите внимание, что код не загружает `isNullObject` свойство. Office автоматически загружает это свойство при `context.sync` его вызове, поэтому нет необходимости явно загружать его с аналогичным действием `datasheet.load('isNullObject')` .

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a>См. также

* [Общая объектная модель API JavaScript](office-javascript-api-object-model.md)
* [Ограничения ресурсов и оптимизация производительности надстроек Office](../concepts/resource-limits-and-performance-optimization.md)

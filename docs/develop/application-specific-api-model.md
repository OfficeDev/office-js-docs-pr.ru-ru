---
title: Использование модели API для определенных приложений
description: Сведения о модели API на основе обещаний для надстроек Excel, OneNote и Word.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 568494dc0b92f1a4f9c6556b169293e68ae0bce9
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773498"
---
# <a name="application-specific-api-model"></a>Модель API для конкретных приложений

В этой статье описано, как использовать модель API для создания надстроек в Excel, Word и OneNote. Здесь представлены основные понятия, лежащие в основе использования API на основе обещаний.

> [!NOTE]
> Эта модель не поддерживается клиентами Office 2013. Используйте [общую модель API](office-javascript-api-object-model.md) для работы с этими версиями Office. Полные сведения о доступности платформ см. в статье [Доступность клиентских приложений и платформ Office для надстроек Office](../overview/office-add-in-availability.md).

> [!TIP]
> В примерах на этой странице используются API JavaScript для Excel, но эти понятия также относятся к API JavaScript для OneNote, Visio и Word.

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>Асинхронный характер API на основе обещаний

Надстройки Office — это веб-сайты, отображающиеся внутри контейнера браузера в приложениях Office, таких как Excel. Этот контейнер внедряется в приложение Office на платформах для классических ПК, например Office для Windows, и запускается в элементе iFrame HTML в Office для Интернета. Из-за соображений производительности интерфейсы API Office.js не могут синхронно взаимодействовать с приложениями Office на всех платформах. Таким образом, вызов API `sync()` в Office.js возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается, когда приложение Office выполняет запрошенные действия чтения или записи. Кроме того, вы можете поместить в очередь несколько действий, например действия настройки свойств или вызова методов, а затем запустить их в виде пакета команд в одном вызове метода `sync()`, а не отправлять отдельные запросы для каждого действия. В разделах ниже описано, как сделать это, используя API `run()` и `sync()`.

## <a name="run-function"></a>Функция *.run

`Excel.run`, `Word.run` и `OneNote.run` исполняют функцию, определяющую действия, выполняемые в Excel, Word и OneNote. `*.run` автоматически создает контекст запроса, который можно использовать для взаимодействия с объектами Office. Когда `*.run` завершает работу, обещание разрешается, и все объекты, которые были выделены в среде выполнения, будут автоматически разблокированы.

В следующем примере показано, как использовать шаблон `Excel.run`. Такой же шаблон используется в Word и OneNote.

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

Приложение Office и ваша надстройка работают в двух разных процессах. Так как они используют разные среды выполнения, надстройкам требуется объект `RequestContext`, чтобы можно было подключать надстройку к объектам в Office, например к листам, диапазонам, абзацам и таблицам. Этот объект `RequestContext` предоставляется в качестве аргумента при вызове `*.run`.

## <a name="proxy-objects"></a>Прокси-объекты

Объекты JavaScript для Office, объявляемые и используемые с помощью API на основе обещаний, являются прокси-объектами.  Все методы, которые вы вызываете, либо свойства, которые вы настраиваете либо загружаете, в прокси-объектах просто добавляются в очередь команд, ожидающих выполнения. Когда вы вызываете метод `sync()` в контексте запроса (например, `context.sync()`), команды, помещенные в очередь, передаются в приложение Office и выполняются. По существу, эти API ориентированы на работу с пакетами. Вы можете поместить в очередь любое количество изменений в контексте запроса, а затем вызвать метод `sync()`, чтобы запустить пакет команд, помещенных в очередь.

Например, во фрагменте кода ниже показано, как объявить локальный объект JavaScript [Excel.Range](/javascript/api/excel/excel.range) (`selectedRange`) для ссылки на выделенный диапазон в книге Excel, а затем задать ряд свойств для этого объекта. Объект `selectedRange` представляет собой прокси-объект, поэтому свойства, заданные в этом объекте, и метод, вызываемый в этом объекте, не будут отображены в документе Excel, пока надстройка не вызовет метод `context.sync()`.

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>Совет по производительности: минимизируйте количество созданных прокси-объектов

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

При вызове метода `sync()` в контексте запроса будет синхронизировано состояние прокси-объектов и объектов в документе Office. Метод `sync()` запускает любые команды, помещенные в очередь в контексте запроса, и получает значения для любых свойств, которые следует загрузить в прокси-объектах. Метод `sync()` выполняется асинхронно и возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается по завершении работы метода `sync()`.

В примере ниже показана пакетная функция, которая определяет локальный прокси-объект JavaScript (`selectedRange`), загружает свойство этого объекта, а затем использует шаблон обещаний JavaScript для вызова метода `context.sync()` и, соответственно, синхронизации состояния прокси-объектов и объектов в документе Excel.

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

В предыдущем примере настроен параметр `selectedRange`, и его свойство `address` загружается при вызове `context.sync()`.

Так как `sync()` — это асинхронная операция, всегда следует возвращать объект `Promise`, чтобы завершить операцию `sync()`, прежде чем продолжить выполнение сценария. Если вы используете TypeScript или JavaScript ES6+, вы можете `await` вызов `context.sync()` вместо возврата обещания.

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>Совет по производительности: минимизируйте количество вызовов синхронизации

В API JavaScript для Excel `sync()` является единственной асинхронной операцией и в некоторых обстоятельствах может выполняться медленно, особенно в случае с Excel в Интернете. Для оптимизации производительности минимизируйте количество вызовов `sync()`, поставив в очередь максимально возможное количество изменений до ее вызова. Дополнительные сведения об оптимизации производительности с помощью `sync()` см. в статье [Избегайте использования метода context.sync в циклах](../concepts/correlated-objects-pattern.md).

### <a name="load"></a>load()

Чтобы можно было считывать свойства прокси-объекта, вам необходимо явно загрузить их и заполнить прокси-объект данными из документа Office, а затем вызвать метод `context.sync()`. Например, вы создали прокси-объект для ссылки на выделенный диапазон, а затем вам потребовалось считать свойство `address` выделенного диапазона. Прежде чем вы сможете считать свойство `address`, вам потребуется загрузить его. Чтобы запросить загрузку свойств прокси-объекта, вызовите метод `load()` в объекте и укажите свойства, которые необходимо загрузить. В следующем примере показана загрузка свойства `Range.address` для `myRange`.

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
> Если вы вызываете методы или задаете свойства только в прокси-объекте, вам не нужно вызывать метод `load()`. Метод `load()` требуется только тогда, когда вам необходимо считать свойства в прокси-объекте.

Аналогично запросам для задания свойств или вызова методов в прокси-объектах, запросы на загрузку свойств в прокси-объектах добавляются в очередь команд, ожидающих выполнения, в контексте запроса, который будет запущен, когда вы в следующий раз вызовете метод `sync()`. В очередь можно поставить сколько угодно вызовов `load()` в контексте запроса.

#### <a name="scalar-and-navigation-properties"></a>Скалярные и навигационные свойства

Существует две категории свойств: **скалярные** и **навигационные**. К скалярным свойствам относятся назначаемые типы, такие как строки, целые числа и структуры JSON. Свойства навигации — это объекты и коллекции объектов только для чтения, которым назначаются поля вместо прямого назначения свойства. Например, элементы `name` и `position` объекта [Excel.Worksheet](/javascript/api/excel/excel.worksheet) являются скалярными свойствами, а `protection` и `tables` — свойствами навигации.

Надстройка может использовать свойства навигации в качестве пути для загрузки определенных скалярных свойств. Следующий код помещает в очередь команду `load` для имени шрифта, используемого объектом `Excel.Range`, без загрузки каких-либо других сведений.

```js
someRange.load("format/font/name")
```

Вы также можете задавать скалярные свойства из свойства навигации по пути к ним. Например, вы можете задать размер шрифта для `Excel.Range` с помощью команды `someRange.format.font.size = 10;`. Чтобы задать свойство, необязательно загружать его.

Имейте в виду, что некоторые свойства объекта могут совпадать с именем другого объекта. Например, `format` — это свойство объекта `Excel.Range`, но также имеется и объект `format`. Поэтому если вы, например, вызываете `range.load("format")`, это эквивалентно `range.format.load()` (нежелательный пустой оператор `load()`). Чтобы избежать этого, ваш код должен загружать только "конечные узлы" в дереве объектов.

#### <a name="calling-load-without-parameters-not-recommended"></a>Вызов метода `load` без параметров (не рекомендуется)

Если вызвать метод `load()` для объекта (или коллекции), не указывая параметры, будут загружены все скалярные свойства объекта или объектов в коллекции. Загрузка ненужных данных замедлит вашу надстройку. Необходимо всегда явным образом указывать свойства для загрузки.

> [!IMPORTANT]
> Объем данных, возвращаемых оператором `load` без параметров, может превышать ограничения по размерам для службы. Чтобы сократить риски для старых надстроек, некоторые свойства не возвращаются методом `load` без их явного запроса. Следующие свойства исключаются из таких операций нагрузки.
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

Методы в API на основе обещаний, возвращающие примитивные типы, используют шаблон, похожий на парадигму `load`/`sync`. Например, `Excel.TableCollection.getCount` получает количество таблиц в коллекции. `getCount` возвращает `ClientResult<number>`. Это означает, что свойство `value` возвращаемого [`ClientResult`](/javascript/api/office/officeextension.clientresult) выражено числом. Сценарий не может получить доступ к этому значению, пока не вызовет `context.sync()`.

Следующий код получает общее количество таблиц в книге Excel и записывает его в консоль.

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

Установка свойств объекта с вложенными свойствами навигации может быть трудоемкой задачей. Вместо того чтобы задавать отдельные свойства с помощью путей навигации, как описано выше, вы можете использовать метод `object.set()`, доступный для объектов в API JavaScript на основе обещаний. С помощью этого метода можно задать сразу несколько свойств объекта, передавая другой объект того же типа Office.js или объект JavaScript со свойствами, сходными по структуре со свойствами объекта, для которого вызывается метод.

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

Некоторые свойства невозможно задать, хотя они и поддерживают запись. Эти свойства являются частью родительского свойства, которое должно быть задано как один объект. Это связано с тем, что родительское свойство использует вложенные свойства с определенными логическими связями. Эти родительские свойства должны быть заданы с помощью нотации литерала объекта, чтобы задать весь объект, а не отдельные вложенные свойства этого объекта.  Один из примеров доступен в разделе [PageLayout](/javascript/api/excel/excel.pagelayout). Свойство `zoom` должно быть установлено с помощью одного объекта [PageLayoutZoomOptions,](/javascript/api/excel/excel.pagelayoutzoomoptions) как показано здесь.

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

В предыдущем примере вы ***не*** сможете напрямую присвоить значение `zoom`: `sheet.pageLayout.zoom.scale = 200;`. Этот оператор выдает ошибку, так как `zoom` не загружен. Даже если `zoom` загружен, масштабный набор не будет работать. Все контекстные операции происходят в `zoom`, обновляя прокси-объект в надстройке и переписывая локально установленные значения.

Это поведение отличается от [свойств навигации](application-specific-api-model.md#scalar-and-navigation-properties), например [Range.format](/javascript/api/excel/excel.range#format). Свойства объектов `format` можно установить с помощью объектной навигации, как показано здесь.

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Вы можете определить свойство, для которого невозможно напрямую задать его вложенные свойства, путем проверки модификатора только для чтения. Для всех свойств, доступных только для чтения, можно напрямую задать их вложенные свойства, использующиеся не только для чтения. Записываемые свойства, например `PageLayout.zoom`, должны быть заданы на уровне объекта. Сводка:

- Свойство только для чтения: вложенные свойства можно задать с помощью навигации.
- Записываемое свойство: вложенные свойства нельзя задать с помощью навигации (необходимо установить их в рамках начального назначения родительского объекта).

## <a name="42ornullobject-methods-and-properties"></a>Методы и свойства &#42;OrNullObject

Некоторые методы и свойства доступа создают исключение, если нужный объект не существует. Например, если для получения листа Excel указать имя листа, не существующее в книге, метод `getItem()` создаст исключение `ItemNotFound`. Библиотеки конкретных приложений позволяют коду проверять наличие сущностей документа, не требуя кода обработки исключений.  Это достигается с помощью вариантов методов и свойств `*OrNullObject`.  Эти варианты вместо создания исключения возвращают объект, свойству `isNullObject` которого присвоено значение `true`, если указанный элемент не существует.

Например, вы можете вызвать метод `getItemOrNullObject()` для коллекции, такой как **Worksheets**, чтобы получить элемент из коллекции. Метод `getItemOrNullObject()` возвращает указанный элемент, если он существует. В противном случае возвращается объект, свойству `isNullObject` которого присвоено значение `true`. Затем код может оценить это свойство, чтобы определить, существует ли объект.

> [!NOTE]
> Варианты `*OrNullObject` никогда не возвращают значение JavaScript `null`. Они возвращают обычные прокси-объекты Office. Если сущность, представляемая объектом, не существует, свойству `isNullObject` объекта присваивается значение `true`. Не проверяйте возвращенный объект на нулевое значение или ложность. Для него никогда не используются значения `null`, `false` или `undefined`.

В следующем примере кода осуществляется попытка получить лист Excel с именем Data с помощью метода `getItemOrNullObject()`. Если лист с таким именем не существует, создается новый лист. Обратите внимание, что код не загружает свойство `isNullObject`. Office автоматически загружает это свойство, когда вызывается `context.sync`, поэтому вам не нужно явным образом загружать его с помощью `datasheet.load('isNullObject')`.

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

---
title: Работа с событиями при помощи API JavaScript для Excel
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 08653a84c051709d16371d89672d3f7ebe2030b7
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872020"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Работа с событиями при помощи API JavaScript для Excel

В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel. 

## <a name="events-in-excel"></a>События в Excel

Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Ниже перечислены поддерживаемые в настоящее время события.

| Событие | Описание | Поддерживаемые объекты |
|:---------------|:-------------|:-----------|
| `onAdded` | Событие, возникающее при добавлении объекта. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | Событие, возникающее при удалении объекта. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | Событие, возникающее при активации объекта. | [**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet) |
| `onDeactivated` | Событие, возникающее при отключении объекта. | [**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet) |
| `onCalculated` | Событие, возникающее после завершения вычислений на листе (или на всех листах коллекции). | [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet) |
| `onChanged` | Событие, возникающее при изменении данных в ячейках. | [**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection) |
| `onDataChanged` | Событие, возникающее при изменении данных или форматирования в привязке. | [**Binding**](/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | Событие, возникающее при изменении активной ячейки или выбранного диапазона. | [**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**Binding**](/javascript/api/excel/excel.binding) |
| `onSettingsChanged` | Событие, возникающее при изменении параметров в документе. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a>Триггеры событий

События в книге Excel могут вызываться:

- при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;
- из кода (JavaScript) надстройки Office, вносящего изменения в книгу;
- из кода (макроса) надстройки VBA, вносящего изменения в книгу.

Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.

### <a name="lifecycle-of-an-event-handler"></a>Жизненный цикл обработчика событий

Обработчик событий создается при его регистрации надстройкой. Он удаляется при отмене его регистрации надстройкой или при обновлении, перезагрузке или закрытии надстройки. Обработчики событий не остаются в составе файла Excel или между сеансами с Excel Online.

> [!CAUTION]
> Когда объект, для которого зарегистрированы события, удаляется (например, таблица с зарегистрированным событием `onChanged`), обработчик событий больше не запускается, но остается в памяти, пока сеанс надстройки или Excel не обновится или не закроется.

### <a name="events-and-coauthoring"></a>События и совместное редактирование

Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, в частности `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Регистрация обработчика событий

В приведенном ниже примере кода регистрируется обработчик события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>Обработка событий

Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль. 

```js
function handleChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>Удаление обработчика события

В приведенном ниже примере кода регистрируется обработчик событий `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события. В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();

        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="enable-and-disable-events"></a>Включение и отключение событий

Производительность надстройки можно повысить с помощью отключения событий. Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей.

События включаются и отключаются на уровне [среды выполнения](/javascript/api/excel/excel.runtime).
Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики.

В приведенном ниже примере кода показано, как включать и отключать события.

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)

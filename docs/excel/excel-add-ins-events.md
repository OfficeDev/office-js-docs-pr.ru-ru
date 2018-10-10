---
title: Работа с событиями с помощью API JavaScript для Excel
description: ''
ms.date: 09/21/2018
ms.openlocfilehash: b56d25e7e0306b4881115397d4136e63ddc03e5c
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459177"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Работа с событиями с помощью API JavaScript для Excel 

В этой статье описаны важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий с помощью API JavaScript для Excel. 

## <a name="events-in-excel"></a>События в Excel

Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Далее перечислены поддерживаемые в настоящее время события.

| Событие | Описание | Поддерживаемые объекты |
|:---------------|:-------------|:-----------|
| `onAdded` | Событие, происходящее при добавлении объекта. | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | Событие, происходящее  при удалении объекта. | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | Событие, происходящее  при активации объекта. | [**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onDeactivated` | Событие, происходящее  при отключении объекта. | [**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onCalculated` | Событие, происходящее после завершения расчета на листе (или на всех листах коллекции). | [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onChanged` | Событие, происходящее при изменении данных в ячейках. | [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection) |
| `onDataChanged` | Событие, происходящее при изменении данных или форматирования в привязке. | [**Привязка**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | Событие, происходящее при изменении активной ячейки или выбранного диапазона. | [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSettingsChanged` | Событие, происходящее при изменении параметров в документе. | [**SettingCollection**](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a>Триггеры событий

События в книге Excel могут вызываться:

- изменениями, вносимыми в книгу пользователем с помощью пользовательского интерфейса Excel;
- изменениями, вносимыми в книгу кодом надстройки Office (JavaScript);
- изменениями, вносимыми в книгу кодом (макросом) надстройки VBA.

Любое изменение, которое отвечает требованиям реакции на событие Excel по умолчанию, вызывает соответствующие события в книге.

### <a name="lifecycle-of-an-event-handler"></a>Жизненный цикл обработчика событий

Обработчик событий создается при его регистрации надстройкой и удаляется при отмене его регистрации или закрытии надстройки. Обработчики событий не остаются в составе файла Excel.

### <a name="events-and-coauthoring"></a>События и совместное редактирование

Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, например `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Регистрация обработчика событий

В приведенном ниже примере кода регистрируется обработчик событий для события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.

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

Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Можно настроить эту функцию на выполнение любых действий, необходимых для сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль. 

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

## <a name="remove-an-event-handler"></a>Удаление обработчика событий

В приведенном ниже примере кода регистрируется обработчик событий для события `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события. В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.

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

Производительность надстройки можно повысить, отключив событие. Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей. 

События включаются и отключаются на уровне [среды выполнения](https://docs.microsoft.com/javascript/api/excel/excel.runtime). Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики. 

Следующий пример кода показывает, как включать и отключать события.

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

- [Основные принципы программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)
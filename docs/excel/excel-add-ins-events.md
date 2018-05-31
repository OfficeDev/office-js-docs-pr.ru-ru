---
title: Работа с событиями при помощи API JavaScript для Excel
description: ''
ms.date: 01/29/2018
ms.openlocfilehash: 4e04b31e7a130f21d6a9c94d041dc2a122a5890e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437474"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Работа с событиями при помощи API JavaScript для Excel 

В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel. 

> [!IMPORTANT]
> В настоящее время API, описанные в этой статье, представлены только в общедоступной ознакомительной бета-версии и не предназначены для использования в рабочей среде. Чтобы запускать содержащиеся в этой статье примеры кода, необходимо использовать достаточно позднюю сборку Office и ссылаться на бета-версию библиотеки в сети CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

## <a name="events-in-excel"></a>События в Excel

Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Ниже перечислены поддерживаемые в настоящее время события.

| Событие | Описание | Поддерживаемые объекты |
|:---------------|:-------------|:-----------|
| `onAdded` | Событие, возникающее при добавлении объекта. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetaddedeventargs.md) |
| `onActivated` | Событие, возникающее при активации объекта. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md) |
| `onDeactivated` | Событие, возникающее при отключении объекта. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md) |
| `onChanged` | Событие, возникающее при изменении данных в ячейках. | [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**TableCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**Binding**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingdatachangedeventargs.md) |
| `onSelectionChanged` | Событие, возникающее при изменении активной ячейки или выбранного диапазона. | [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tableselectionchangedeventargs.md), [**Binding**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingselectionchangedeventargs.md) |

### <a name="event-triggers"></a>Триггеры событий

События в книге Excel могут вызываться:

- при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;
- из кода (JavaScript) надстройки Office, вносящего изменения в книгу;
- из кода (макроса) надстройки VBA, вносящего изменения в книгу.

Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.

### <a name="lifecycle-of-an-event-handler"></a>Жизненный цикл обработчика событий

Обработчик событий создается, когда надстройка регистрирует его, и удаляется при отмене его регистрации или закрытии надстройки. Обработчики событий не остаются в составе файла Excel.

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

## <a name="see-also"></a>См. также

- [Основные понятия API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Открытая спецификация по API JavaScript для Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Общие сведения о функциях событий Excel (ознакомительная версия)](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)

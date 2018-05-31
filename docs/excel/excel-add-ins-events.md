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
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="ef662-102">Работа с событиями при помощи API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ef662-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="ef662-103">В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="ef662-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="ef662-104">В настоящее время API, описанные в этой статье, представлены только в общедоступной ознакомительной бета-версии и не предназначены для использования в рабочей среде.</span><span class="sxs-lookup"><span data-stu-id="ef662-104">The APIs described in this article are currently available only in public preview (beta) and are not intended for use in production environments.</span></span> <span data-ttu-id="ef662-105">Чтобы запускать содержащиеся в этой статье примеры кода, необходимо использовать достаточно позднюю сборку Office и ссылаться на бета-версию библиотеки в сети CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="ef662-105">To run the code samples that this article contains, you must use a sufficiently recent build of Office and reference the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="ef662-106">События в Excel</span><span class="sxs-lookup"><span data-stu-id="ef662-106">Events in Excel</span></span>

<span data-ttu-id="ef662-107">Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии.</span><span class="sxs-lookup"><span data-stu-id="ef662-107">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="ef662-108">С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события.</span><span class="sxs-lookup"><span data-stu-id="ef662-108">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="ef662-109">Ниже перечислены поддерживаемые в настоящее время события.</span><span class="sxs-lookup"><span data-stu-id="ef662-109">The following events are currently supported.</span></span>

| <span data-ttu-id="ef662-110">Событие</span><span class="sxs-lookup"><span data-stu-id="ef662-110">Event</span></span> | <span data-ttu-id="ef662-111">Описание</span><span class="sxs-lookup"><span data-stu-id="ef662-111">Description</span></span> | <span data-ttu-id="ef662-112">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="ef662-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="ef662-113">Событие, возникающее при добавлении объекта.</span><span class="sxs-lookup"><span data-stu-id="ef662-113">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="ef662-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="ef662-114">**WorksheetCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetaddedeventargs.md) |
| `onActivated` | <span data-ttu-id="ef662-115">Событие, возникающее при активации объекта.</span><span class="sxs-lookup"><span data-stu-id="ef662-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="ef662-116">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="ef662-116">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="ef662-117">Событие, возникающее при отключении объекта.</span><span class="sxs-lookup"><span data-stu-id="ef662-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="ef662-118">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="ef662-118">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span></span> |
| `onChanged` | <span data-ttu-id="ef662-119">Событие, возникающее при изменении данных в ячейках.</span><span class="sxs-lookup"><span data-stu-id="ef662-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="ef662-120">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**TableCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**Binding**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingdatachangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="ef662-120">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), **TableCollection**, [Binding](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="ef662-121">Событие, возникающее при изменении активной ячейки или выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="ef662-121">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="ef662-122">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tableselectionchangedeventargs.md), [**Binding**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingselectionchangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="ef662-122">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), **Binding**</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="ef662-123">Триггеры событий</span><span class="sxs-lookup"><span data-stu-id="ef662-123">Event triggers</span></span>

<span data-ttu-id="ef662-124">События в книге Excel могут вызываться:</span><span class="sxs-lookup"><span data-stu-id="ef662-124">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="ef662-125">при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="ef662-125">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="ef662-126">из кода (JavaScript) надстройки Office, вносящего изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="ef662-126">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="ef662-127">из кода (макроса) надстройки VBA, вносящего изменения в книгу.</span><span class="sxs-lookup"><span data-stu-id="ef662-127">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="ef662-128">Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.</span><span class="sxs-lookup"><span data-stu-id="ef662-128">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="ef662-129">Жизненный цикл обработчика событий</span><span class="sxs-lookup"><span data-stu-id="ef662-129">Lifecycle of an event handler</span></span>

<span data-ttu-id="ef662-p103">Обработчик событий создается, когда надстройка регистрирует его, и удаляется при отмене его регистрации или закрытии надстройки. Обработчики событий не остаются в составе файла Excel.</span><span class="sxs-lookup"><span data-stu-id="ef662-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="ef662-132">События и совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="ef662-132">Events and coauthoring</span></span>

<span data-ttu-id="ef662-p104">Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, в частности `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="ef662-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="ef662-135">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="ef662-135">Register an event handler</span></span>

<span data-ttu-id="ef662-136">В приведенном ниже примере кода регистрируется обработчик события `onChanged` на листе под названием **Sample**.</span><span class="sxs-lookup"><span data-stu-id="ef662-136">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="ef662-137">В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.</span><span class="sxs-lookup"><span data-stu-id="ef662-137">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="ef662-138">Обработка событий</span><span class="sxs-lookup"><span data-stu-id="ef662-138">Handle an event</span></span>

<span data-ttu-id="ef662-139">Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события.</span><span class="sxs-lookup"><span data-stu-id="ef662-139">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="ef662-140">Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария.</span><span class="sxs-lookup"><span data-stu-id="ef662-140">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="ef662-141">В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.</span><span class="sxs-lookup"><span data-stu-id="ef662-141">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="ef662-142">Удаление обработчика события</span><span class="sxs-lookup"><span data-stu-id="ef662-142">Remove an event handler</span></span>

<span data-ttu-id="ef662-143">В приведенном ниже примере кода регистрируется обработчик событий `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события.</span><span class="sxs-lookup"><span data-stu-id="ef662-143">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="ef662-144">В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="ef662-144">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="ef662-145">См. также</span><span class="sxs-lookup"><span data-stu-id="ef662-145">See also</span></span>

- [<span data-ttu-id="ef662-146">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ef662-146">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ef662-147">Открытая спецификация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ef662-147">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="ef662-148">Общие сведения о функциях событий Excel (ознакомительная версия)</span><span class="sxs-lookup"><span data-stu-id="ef662-148">Introduction to Excel Event Features (preview)</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)

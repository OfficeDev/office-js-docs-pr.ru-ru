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
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="0344e-102">Работа с событиями с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="0344e-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="0344e-103">В этой статье описаны важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="0344e-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="0344e-104">События в Excel</span><span class="sxs-lookup"><span data-stu-id="0344e-104">Events in Excel</span></span>

<span data-ttu-id="0344e-p101">Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Далее перечислены поддерживаемые в настоящее время события.</span><span class="sxs-lookup"><span data-stu-id="0344e-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="0344e-108">Событие</span><span class="sxs-lookup"><span data-stu-id="0344e-108">Event</span></span> | <span data-ttu-id="0344e-109">Описание</span><span class="sxs-lookup"><span data-stu-id="0344e-109">Description</span></span> | <span data-ttu-id="0344e-110">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="0344e-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="0344e-111">Событие, происходящее при добавлении объекта.</span><span class="sxs-lookup"><span data-stu-id="0344e-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="0344e-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="0344e-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="0344e-113">Событие, происходящее  при удалении объекта.</span><span class="sxs-lookup"><span data-stu-id="0344e-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="0344e-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="0344e-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="0344e-115">Событие, происходящее  при активации объекта.</span><span class="sxs-lookup"><span data-stu-id="0344e-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="0344e-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="0344e-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="0344e-117">Событие, происходящее  при отключении объекта.</span><span class="sxs-lookup"><span data-stu-id="0344e-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="0344e-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="0344e-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="0344e-119">Событие, происходящее после завершения расчета на листе (или на всех листах коллекции).</span><span class="sxs-lookup"><span data-stu-id="0344e-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="0344e-120">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="0344e-120">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="0344e-121">Событие, происходящее при изменении данных в ячейках.</span><span class="sxs-lookup"><span data-stu-id="0344e-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="0344e-122">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="0344e-122">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **TableCollection**, [Binding](https://docs.microsoft.com/javascript/api/excel/excel.table)</span></span> |
| `onDataChanged` | <span data-ttu-id="0344e-123">Событие, происходящее при изменении данных или форматирования в привязке.</span><span class="sxs-lookup"><span data-stu-id="0344e-123">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="0344e-124">**Привязка**</span><span class="sxs-lookup"><span data-stu-id="0344e-124">**Binding**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="0344e-125">Событие, происходящее при изменении активной ячейки или выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="0344e-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="0344e-126">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="0344e-126">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="0344e-127">Событие, происходящее при изменении параметров в документе.</span><span class="sxs-lookup"><span data-stu-id="0344e-127">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="0344e-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="0344e-128">**SettingCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="0344e-129">Триггеры событий</span><span class="sxs-lookup"><span data-stu-id="0344e-129">Event triggers</span></span>

<span data-ttu-id="0344e-130">События в книге Excel могут вызываться:</span><span class="sxs-lookup"><span data-stu-id="0344e-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="0344e-131">изменениями, вносимыми в книгу пользователем с помощью пользовательского интерфейса Excel;</span><span class="sxs-lookup"><span data-stu-id="0344e-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="0344e-132">изменениями, вносимыми в книгу кодом надстройки Office (JavaScript);</span><span class="sxs-lookup"><span data-stu-id="0344e-132">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="0344e-133">изменениями, вносимыми в книгу кодом (макросом) надстройки VBA.</span><span class="sxs-lookup"><span data-stu-id="0344e-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="0344e-134">Любое изменение, которое отвечает требованиям реакции на событие Excel по умолчанию, вызывает соответствующие события в книге.</span><span class="sxs-lookup"><span data-stu-id="0344e-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="0344e-135">Жизненный цикл обработчика событий</span><span class="sxs-lookup"><span data-stu-id="0344e-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="0344e-p102">Обработчик событий создается при его регистрации надстройкой и удаляется при отмене его регистрации или закрытии надстройки. Обработчики событий не остаются в составе файла Excel.</span><span class="sxs-lookup"><span data-stu-id="0344e-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="0344e-138">События и совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="0344e-138">Events and coauthoring</span></span>

<span data-ttu-id="0344e-p103">Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, например `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="0344e-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="0344e-141">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="0344e-141">Register an event handler</span></span>

<span data-ttu-id="0344e-p104">В приведенном ниже примере кода регистрируется обработчик событий для события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.</span><span class="sxs-lookup"><span data-stu-id="0344e-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="0344e-144">Обработка событий</span><span class="sxs-lookup"><span data-stu-id="0344e-144">Handle an event</span></span>

<span data-ttu-id="0344e-p105">Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Можно настроить эту функцию на выполнение любых действий, необходимых для сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.</span><span class="sxs-lookup"><span data-stu-id="0344e-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="0344e-148">Удаление обработчика событий</span><span class="sxs-lookup"><span data-stu-id="0344e-148">Remove an event handler</span></span>

<span data-ttu-id="0344e-p106">В приведенном ниже примере кода регистрируется обработчик событий для события `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события. В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="0344e-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="0344e-151">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="0344e-151">Enable and disable events</span></span>

<span data-ttu-id="0344e-p107">Производительность надстройки можно повысить, отключив событие. Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей.</span><span class="sxs-lookup"><span data-stu-id="0344e-p107">The performance of an add-in may be improved by disabling events. For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span> 

<span data-ttu-id="0344e-p108">События включаются и отключаются на уровне [среды выполнения](https://docs.microsoft.com/javascript/api/excel/excel.runtime). Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики.</span><span class="sxs-lookup"><span data-stu-id="0344e-p108">Events are enabled and disabled at the [runtime](https://docs.microsoft.com/javascript/api/excel/excel.runtime) level. The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> 

<span data-ttu-id="0344e-156">Следующий пример кода показывает, как включать и отключать события.</span><span class="sxs-lookup"><span data-stu-id="0344e-156">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="0344e-157">См. также</span><span class="sxs-lookup"><span data-stu-id="0344e-157">See also</span></span>

- [<span data-ttu-id="0344e-158">Основные принципы программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="0344e-158">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
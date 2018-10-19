---
title: Работа с событиями с помощью API JavaScript для Excel
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: c3fbdf27dcbedf0d006973e6ebc2e01b02e6cec2
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639940"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="73226-102">Работа с событиями с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="73226-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="73226-103">В этой статье описаны важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="73226-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="73226-104">События в Excel</span><span class="sxs-lookup"><span data-stu-id="73226-104">Events in Excel</span></span>

<span data-ttu-id="73226-p101">Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Далее перечислены поддерживаемые в настоящее время события.</span><span class="sxs-lookup"><span data-stu-id="73226-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="73226-108">Событие</span><span class="sxs-lookup"><span data-stu-id="73226-108">Event</span></span> | <span data-ttu-id="73226-109">Описание</span><span class="sxs-lookup"><span data-stu-id="73226-109">Description</span></span> | <span data-ttu-id="73226-110">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="73226-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="73226-111">Событие, происходящее при добавлении объекта.</span><span class="sxs-lookup"><span data-stu-id="73226-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="73226-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="73226-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="73226-113">Событие, происходящее  при удалении объекта.</span><span class="sxs-lookup"><span data-stu-id="73226-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="73226-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="73226-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="73226-115">Событие, происходящее  при активации объекта.</span><span class="sxs-lookup"><span data-stu-id="73226-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="73226-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="73226-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="73226-117">Событие, происходящее  при отключении объекта.</span><span class="sxs-lookup"><span data-stu-id="73226-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="73226-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="73226-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="73226-119">Событие, происходящее после завершения расчета на листе (или на всех листах коллекции).</span><span class="sxs-lookup"><span data-stu-id="73226-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="73226-120">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="73226-120">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="73226-121">Событие, происходящее при изменении данных в ячейках.</span><span class="sxs-lookup"><span data-stu-id="73226-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="73226-122">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="73226-122">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **TableCollection**, [Binding](https://docs.microsoft.com/javascript/api/excel/excel.table)</span></span> |
| `onDataChanged` | <span data-ttu-id="73226-123">Событие, происходящее при изменении данных или форматирования в привязке.</span><span class="sxs-lookup"><span data-stu-id="73226-123">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="73226-124">**Привязка**</span><span class="sxs-lookup"><span data-stu-id="73226-124">**Binding**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="73226-125">Событие, происходящее при изменении активной ячейки или выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="73226-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="73226-126">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="73226-126">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="73226-127">Событие, происходящее при изменении параметров в документе.</span><span class="sxs-lookup"><span data-stu-id="73226-127">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="73226-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="73226-128">**settingCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="73226-129">Триггеры событий</span><span class="sxs-lookup"><span data-stu-id="73226-129">Event triggers</span></span>

<span data-ttu-id="73226-130">События в книге Excel могут вызываться:</span><span class="sxs-lookup"><span data-stu-id="73226-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="73226-131">изменениями, вносимыми в книгу пользователем с помощью пользовательского интерфейса Excel;</span><span class="sxs-lookup"><span data-stu-id="73226-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="73226-132">изменениями, вносимыми в книгу кодом надстройки Office (JavaScript);</span><span class="sxs-lookup"><span data-stu-id="73226-132">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="73226-133">изменениями, вносимыми в книгу кодом (макросом) надстройки VBA.</span><span class="sxs-lookup"><span data-stu-id="73226-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="73226-134">Любое изменение, которое отвечает требованиям реакции на событие Excel по умолчанию, вызывает соответствующие события в книге.</span><span class="sxs-lookup"><span data-stu-id="73226-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="73226-135">Жизненный цикл обработчика событий</span><span class="sxs-lookup"><span data-stu-id="73226-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="73226-p102">Обработчик событий создается при его регистрации надстройкой. Он удаляется при отмене его регистрации надстройкой или при обновлении, перезагрузке или закрытии надстройки. Обработчики событий не остаются в составе файла Excel или между сеансами с Excel Online.</span><span class="sxs-lookup"><span data-stu-id="73226-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

> [!CAUTION]
> <span data-ttu-id="73226-139">Когда объект, к которому зарегистрированы события, удаляется (например, таблица с зарегистрированным событием `onChanged`), обработчик событий более не запускается, а остается в памяти до тех пор, пока сеанс надстройки или Excel не обновится или не закроется.</span><span class="sxs-lookup"><span data-stu-id="73226-139">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="73226-140">События и совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="73226-140">Events and coauthoring</span></span>

<span data-ttu-id="73226-p103">Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, например `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="73226-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="73226-143">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="73226-143">Register an event handler</span></span>

<span data-ttu-id="73226-p104">В приведенном ниже примере кода регистрируется обработчик событий для события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.</span><span class="sxs-lookup"><span data-stu-id="73226-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="73226-146">Обработка событий</span><span class="sxs-lookup"><span data-stu-id="73226-146">Handle an event</span></span>

<span data-ttu-id="73226-p105">Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Можно настроить эту функцию на выполнение любых действий, необходимых для сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.</span><span class="sxs-lookup"><span data-stu-id="73226-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="73226-150">Удаление обработчика событий</span><span class="sxs-lookup"><span data-stu-id="73226-150">Remove an event handler</span></span>

<span data-ttu-id="73226-p106">В приведенном ниже примере кода регистрируется обработчик событий для события `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события. В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="73226-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="73226-153">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="73226-153">Enable and disable events</span></span>

<span data-ttu-id="73226-p107">Производительность надстройки можно повысить, отключив событие. Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей.</span><span class="sxs-lookup"><span data-stu-id="73226-p107">The performance of an add-in may be improved by disabling events. For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span> 

<span data-ttu-id="73226-p108">События включаются и отключаются на уровне [среды выполнения](https://docs.microsoft.com/javascript/api/excel/excel.runtime). Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики.</span><span class="sxs-lookup"><span data-stu-id="73226-p108">Events are enabled and disabled at the [runtime](https://docs.microsoft.com/javascript/api/excel/excel.runtime) level. The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> 

<span data-ttu-id="73226-158">Следующий пример кода показывает, как включать и отключать события.</span><span class="sxs-lookup"><span data-stu-id="73226-158">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="73226-159">См. также</span><span class="sxs-lookup"><span data-stu-id="73226-159">See also</span></span>

- [<span data-ttu-id="73226-160">Основные принципы программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="73226-160">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
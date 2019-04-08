---
title: Работа с событиями при помощи API JavaScript для Excel
description: ''
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: 7f05263f5220c2d60d0cebcfc686e1fed3f07900
ms.sourcegitcommit: 63219bcc1bb5e3bed7eb6c6b0adb73a4829c7e8f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/05/2019
ms.locfileid: "31479713"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="33540-102">Работа с событиями при помощи API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="33540-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="33540-103">В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="33540-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="33540-104">События в Excel</span><span class="sxs-lookup"><span data-stu-id="33540-104">Events in Excel</span></span>

<span data-ttu-id="33540-p101">Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Ниже перечислены поддерживаемые в настоящее время события.</span><span class="sxs-lookup"><span data-stu-id="33540-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="33540-108">Событие</span><span class="sxs-lookup"><span data-stu-id="33540-108">Event</span></span> | <span data-ttu-id="33540-109">Описание</span><span class="sxs-lookup"><span data-stu-id="33540-109">Description</span></span> | <span data-ttu-id="33540-110">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="33540-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="33540-111">Возникает при активации объекта.</span><span class="sxs-lookup"><span data-stu-id="33540-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="33540-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="33540-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAdded` | <span data-ttu-id="33540-113">Возникает при добавлении объекта.</span><span class="sxs-lookup"><span data-stu-id="33540-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="33540-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="33540-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onCalculated` | <span data-ttu-id="33540-115">Возникает после завершения вычислений на листе (или на всех листах коллекции).</span><span class="sxs-lookup"><span data-stu-id="33540-115">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="33540-116">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="33540-116">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="33540-117">Возникает при изменении данных в ячейках.</span><span class="sxs-lookup"><span data-stu-id="33540-117">Occurs when data within the binding is changed.</span></span> | <span data-ttu-id="33540-118">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="33540-118">[**Worksheet**](/javascript/api/excel/excel.table), [**Table**](/javascript/api/excel/excel.tablecollection), [**TableCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDataChanged` | <span data-ttu-id="33540-119">Возникает при изменении данных или форматирования в привязке.</span><span class="sxs-lookup"><span data-stu-id="33540-119">Occurs when data or formatting within the binding is changed.</span></span> | [**<span data-ttu-id="33540-120">Binding</span><span class="sxs-lookup"><span data-stu-id="33540-120">Binding</span></span>**](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="33540-121">Возникает при отключении объекта.</span><span class="sxs-lookup"><span data-stu-id="33540-121">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="33540-122">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="33540-122">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="33540-123">Возникает при удалении объекта.</span><span class="sxs-lookup"><span data-stu-id="33540-123">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="33540-124">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="33540-124">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="33540-125">Возникает при изменении активной ячейки или выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="33540-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="33540-126">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="33540-126">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="33540-127">Возникает при изменении параметров в документе.</span><span class="sxs-lookup"><span data-stu-id="33540-127">Occurs when the Settings in the document are changed.</span></span> | [**<span data-ttu-id="33540-128">SettingCollection</span><span class="sxs-lookup"><span data-stu-id="33540-128">SettingCollection</span></span>**](/javascript/api/excel/excel.settingcollection) |

### <a name="events-in-preview"></a><span data-ttu-id="33540-129">События в предварительной версии</span><span class="sxs-lookup"><span data-stu-id="33540-129">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="33540-130">Следующие события в настоящее время доступны только в общедоступной предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="33540-130">The  and  methods are currently available only in public preview (beta).</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="33540-131">Событие</span><span class="sxs-lookup"><span data-stu-id="33540-131">Event</span></span> | <span data-ttu-id="33540-132">Описание</span><span class="sxs-lookup"><span data-stu-id="33540-132">Description</span></span> | <span data-ttu-id="33540-133">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="33540-133">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="33540-134">Возникает, если фигура активирована.</span><span class="sxs-lookup"><span data-stu-id="33540-134">Occurs when the shape is activated.</span></span> | [**<span data-ttu-id="33540-135">Shape</span><span class="sxs-lookup"><span data-stu-id="33540-135">Shape</span></span>**](/javascript/api/excel/excel.shape)|
| `onAdded` | <span data-ttu-id="33540-136">Возникает, если в книгу добавлена новая таблица.</span><span class="sxs-lookup"><span data-stu-id="33540-136">Occurs when new table is added in a workbook.</span></span> | [**<span data-ttu-id="33540-137">TableCollection</span><span class="sxs-lookup"><span data-stu-id="33540-137">tableCollection</span></span>**](/javascript/api/excel/excel.tablecollection)|
| `onAutoSaveSettingChanged` | <span data-ttu-id="33540-138">Возникает при изменении параметра `autoSave` для книги.</span><span class="sxs-lookup"><span data-stu-id="33540-138">Occurs when the autoSave setting is changed on the workbook.</span></span> | [**<span data-ttu-id="33540-139">Workbook</span><span class="sxs-lookup"><span data-stu-id="33540-139">Workbook</span></span>**](/javascript/api/excel/excel.workbook) |
| `onChanged` | <span data-ttu-id="33540-140">Возникает при изменении любого листа в книге.</span><span class="sxs-lookup"><span data-stu-id="33540-140">Occurs when any worksheet in the workbook is changed.</span></span> | [**<span data-ttu-id="33540-141">WorksheetCollection</span><span class="sxs-lookup"><span data-stu-id="33540-141">worksheetCollection</span></span>**](/javascript/api/excel/excel.worksheetcollection)|
| `onDeactivated` | <span data-ttu-id="33540-142">Возникает, если фигура деактивирована.</span><span class="sxs-lookup"><span data-stu-id="33540-142">Occurs when the shape is deactivated.</span></span> | [**<span data-ttu-id="33540-143">Shape</span><span class="sxs-lookup"><span data-stu-id="33540-143">Shape</span></span>**](/javascript/api/excel/excel.shape)|
| `onDeleted` | <span data-ttu-id="33540-144">Возникает, если указанная таблица удалена из книги.</span><span class="sxs-lookup"><span data-stu-id="33540-144">Occurs when the specified table is deleted in a workbook.</span></span> | [**<span data-ttu-id="33540-145">TableCollection</span><span class="sxs-lookup"><span data-stu-id="33540-145">tableCollection</span></span>**](/javascript/api/excel/excel.tablecollection)|
| `onFiltered` | <span data-ttu-id="33540-146">Возникает при применении фильтра к объекту.</span><span class="sxs-lookup"><span data-stu-id="33540-146">Occurs when filter is applied on an object.</span></span> | <span data-ttu-id="33540-147">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="33540-147">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="33540-148">Возникает при изменении формата на листе.</span><span class="sxs-lookup"><span data-stu-id="33540-148">Occurs when format changed on a specific worksheet.</span></span> | <span data-ttu-id="33540-149">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="33540-149">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="33540-150">Возникает при изменениях выделения на любом листе.</span><span class="sxs-lookup"><span data-stu-id="33540-150">Occurs when the selection changes on any worksheet.</span></span> | [**<span data-ttu-id="33540-151">WorksheetCollection</span><span class="sxs-lookup"><span data-stu-id="33540-151">worksheetCollection</span></span>**](/javascript/api/excel/excel.worksheetcollection) |

### <a name="event-triggers"></a><span data-ttu-id="33540-152">Триггеры событий</span><span class="sxs-lookup"><span data-stu-id="33540-152">Event triggers</span></span>

<span data-ttu-id="33540-153">События в книге Excel могут вызываться:</span><span class="sxs-lookup"><span data-stu-id="33540-153">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="33540-154">при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="33540-154">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="33540-155">из кода (JavaScript) надстройки Office, вносящего изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="33540-155">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="33540-156">из кода (макроса) надстройки VBA, вносящего изменения в книгу.</span><span class="sxs-lookup"><span data-stu-id="33540-156">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="33540-157">Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.</span><span class="sxs-lookup"><span data-stu-id="33540-157">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="33540-158">Жизненный цикл обработчика событий</span><span class="sxs-lookup"><span data-stu-id="33540-158">Lifecycle of an event handler</span></span>

<span data-ttu-id="33540-159">Обработчик событий создается при его регистрации надстройкой.</span><span class="sxs-lookup"><span data-stu-id="33540-159">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="33540-160">Он удаляется при отмене его регистрации надстройкой или при обновлении, перезагрузке или закрытии надстройки.</span><span class="sxs-lookup"><span data-stu-id="33540-160">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="33540-161">Обработчики событий не остаются в составе файла Excel или между сеансами с Excel Online.</span><span class="sxs-lookup"><span data-stu-id="33540-161">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="33540-162">Когда объект, для которого зарегистрированы события, удаляется (например, таблица с зарегистрированным событием `onChanged`), обработчик событий больше не запускается, но остается в памяти, пока сеанс надстройки или Excel не обновится или не закроется.</span><span class="sxs-lookup"><span data-stu-id="33540-162">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="33540-163">События и совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="33540-163">Events and coauthoring</span></span>

<span data-ttu-id="33540-p104">Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, в частности `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="33540-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="33540-166">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="33540-166">Register an event handler</span></span>

<span data-ttu-id="33540-p105">В приведенном ниже примере кода регистрируется обработчик события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.</span><span class="sxs-lookup"><span data-stu-id="33540-p105">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="33540-169">Обработка событий</span><span class="sxs-lookup"><span data-stu-id="33540-169">Handle an event</span></span>

<span data-ttu-id="33540-p106">Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.</span><span class="sxs-lookup"><span data-stu-id="33540-p106">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="33540-173">Удаление обработчика события</span><span class="sxs-lookup"><span data-stu-id="33540-173">Remove an event handler</span></span>

<span data-ttu-id="33540-p107">В приведенном ниже примере кода регистрируется обработчик событий `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события. В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="33540-p107">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="33540-176">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="33540-176">Enable and disable events</span></span>

<span data-ttu-id="33540-177">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="33540-177">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="33540-178">Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей.</span><span class="sxs-lookup"><span data-stu-id="33540-178">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="33540-179">События включаются и отключаются на уровне [среды выполнения](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="33540-179">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="33540-180">Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики.</span><span class="sxs-lookup"><span data-stu-id="33540-180">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="33540-181">В приведенном ниже примере кода показано, как включать и отключать события.</span><span class="sxs-lookup"><span data-stu-id="33540-181">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="33540-182">См. также</span><span class="sxs-lookup"><span data-stu-id="33540-182">See also</span></span>

- [<span data-ttu-id="33540-183">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="33540-183">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

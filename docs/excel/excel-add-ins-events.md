---
title: Работа с событиями при помощи API JavaScript для Excel
description: ''
ms.date: 09/09/2019
localization_priority: Priority
ms.openlocfilehash: 338085b297e760057987fd94d7caf96b5134ce6a
ms.sourcegitcommit: 24303ca235ebd7144a1d913511d8e4fb7c0e8c0d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2019
ms.locfileid: "36838482"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="82fd0-102">Работа с событиями при помощи API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="82fd0-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="82fd0-103">В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="82fd0-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="82fd0-104">События в Excel</span><span class="sxs-lookup"><span data-stu-id="82fd0-104">Events in Excel</span></span>

<span data-ttu-id="82fd0-p101">Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Ниже перечислены поддерживаемые в настоящее время события.</span><span class="sxs-lookup"><span data-stu-id="82fd0-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="82fd0-108">Событие</span><span class="sxs-lookup"><span data-stu-id="82fd0-108">Event</span></span> | <span data-ttu-id="82fd0-109">Описание</span><span class="sxs-lookup"><span data-stu-id="82fd0-109">Description</span></span> | <span data-ttu-id="82fd0-110">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="82fd0-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="82fd0-111">Возникает при активации объекта.</span><span class="sxs-lookup"><span data-stu-id="82fd0-111">Occurs when an object is activated.</span></span> | <span data-ttu-id="82fd0-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.shape), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onAdded` | <span data-ttu-id="82fd0-113">Возникает при добавлении объекта в коллекцию.</span><span class="sxs-lookup"><span data-stu-id="82fd0-113">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="82fd0-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="82fd0-115">Возникает при изменении параметра `autoSave` для книги.</span><span class="sxs-lookup"><span data-stu-id="82fd0-115">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="82fd0-116">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="82fd0-116">**Workbook**</span></span>](/javascript/api/excel/excel.workbook) |
| `onCalculated` | <span data-ttu-id="82fd0-117">Возникает после завершения вычислений на листе (или на всех листах коллекции).</span><span class="sxs-lookup"><span data-stu-id="82fd0-117">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="82fd0-118">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-118">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="82fd0-119">Возникает при изменении данных в ячейках.</span><span class="sxs-lookup"><span data-stu-id="82fd0-119">Occurs when data within cells is changed.</span></span> | <span data-ttu-id="82fd0-120">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-120">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="82fd0-121">Возникает при изменении данных или форматирования в привязке.</span><span class="sxs-lookup"><span data-stu-id="82fd0-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="82fd0-122">**Binding**</span><span class="sxs-lookup"><span data-stu-id="82fd0-122">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="82fd0-123">Возникает при отключении объекта.</span><span class="sxs-lookup"><span data-stu-id="82fd0-123">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="82fd0-124">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-124">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.shape), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeleted` | <span data-ttu-id="82fd0-125">Возникает при удалении объекта из коллекции.</span><span class="sxs-lookup"><span data-stu-id="82fd0-125">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="82fd0-126">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-126">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="82fd0-127">Возникает при изменении формата на листе.</span><span class="sxs-lookup"><span data-stu-id="82fd0-127">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="82fd0-128">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-128">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="82fd0-129">Возникает при изменении активной ячейки или выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="82fd0-129">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="82fd0-130">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-130">[**Table**](/javascript/api/excel/excel.binding), [**TableCollection**](/javascript/api/excel/excel.table), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="82fd0-131">Возникает при изменении параметров в документе.</span><span class="sxs-lookup"><span data-stu-id="82fd0-131">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="82fd0-132">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="82fd0-132">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |

### <a name="events-in-preview"></a><span data-ttu-id="82fd0-133">События в предварительной версии</span><span class="sxs-lookup"><span data-stu-id="82fd0-133">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="82fd0-134">Следующие события в настоящее время доступны только в общедоступной предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="82fd0-134">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="82fd0-135">Событие</span><span class="sxs-lookup"><span data-stu-id="82fd0-135">Event</span></span> | <span data-ttu-id="82fd0-136">Описание</span><span class="sxs-lookup"><span data-stu-id="82fd0-136">Description</span></span> | <span data-ttu-id="82fd0-137">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="82fd0-137">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onColumnSorted` | <span data-ttu-id="82fd0-138">Возникает при сортировке одного или нескольких столбцов.</span><span class="sxs-lookup"><span data-stu-id="82fd0-138">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="82fd0-139">Происходит в результате операции сортировки слева направо.</span><span class="sxs-lookup"><span data-stu-id="82fd0-139">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="82fd0-140">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-140">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFiltered` | <span data-ttu-id="82fd0-141">Возникает при применении фильтра к объекту.</span><span class="sxs-lookup"><span data-stu-id="82fd0-141">Occurs when filter is applied on an object.</span></span> | <span data-ttu-id="82fd0-142">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-142">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="82fd0-143">Возникает при изменении состояния скрытия строки на определенном листе.</span><span class="sxs-lookup"><span data-stu-id="82fd0-143">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="82fd0-144">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-144">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowSorted` | <span data-ttu-id="82fd0-145">Возникает при сортировке одной или нескольких строк.</span><span class="sxs-lookup"><span data-stu-id="82fd0-145">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="82fd0-146">Происходит в результате операции сортировки сверху вниз.</span><span class="sxs-lookup"><span data-stu-id="82fd0-146">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="82fd0-147">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-147">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSingleClicked` | <span data-ttu-id="82fd0-148">Возникает, когда происходит щелчок левой кнопкой мыши или нажатие на листе.</span><span class="sxs-lookup"><span data-stu-id="82fd0-148">Occurs when left-clicked/tapped operation happens in the worksheet.</span></span> | <span data-ttu-id="82fd0-149">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="82fd0-149">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="82fd0-150">Триггеры событий</span><span class="sxs-lookup"><span data-stu-id="82fd0-150">Event triggers</span></span>

<span data-ttu-id="82fd0-151">События в книге Excel могут вызываться:</span><span class="sxs-lookup"><span data-stu-id="82fd0-151">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="82fd0-152">при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="82fd0-152">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="82fd0-153">из кода (JavaScript) надстройки Office, вносящего изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="82fd0-153">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="82fd0-154">из кода (макроса) надстройки VBA, вносящего изменения в книгу.</span><span class="sxs-lookup"><span data-stu-id="82fd0-154">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="82fd0-155">Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.</span><span class="sxs-lookup"><span data-stu-id="82fd0-155">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="82fd0-156">Жизненный цикл обработчика событий</span><span class="sxs-lookup"><span data-stu-id="82fd0-156">Lifecycle of an event handler</span></span>

<span data-ttu-id="82fd0-157">Обработчик событий создается при его регистрации надстройкой.</span><span class="sxs-lookup"><span data-stu-id="82fd0-157">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="82fd0-158">Он удаляется при отмене его регистрации надстройкой или при обновлении, перезагрузке или закрытии надстройки.</span><span class="sxs-lookup"><span data-stu-id="82fd0-158">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="82fd0-159">Обработчики событий не остаются в составе файла Excel или между сеансами с интернет-версией Excel.</span><span class="sxs-lookup"><span data-stu-id="82fd0-159">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="82fd0-160">Когда объект, для которого зарегистрированы события, удаляется (например, таблица с зарегистрированным событием `onChanged`), обработчик событий больше не запускается, но остается в памяти, пока сеанс надстройки или Excel не обновится или не закроется.</span><span class="sxs-lookup"><span data-stu-id="82fd0-160">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="82fd0-161">События и совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="82fd0-161">Events and coauthoring</span></span>

<span data-ttu-id="82fd0-p106">Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, в частности `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="82fd0-p106">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="82fd0-164">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="82fd0-164">Register an event handler</span></span>

<span data-ttu-id="82fd0-p107">В приведенном ниже примере кода регистрируется обработчик события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.</span><span class="sxs-lookup"><span data-stu-id="82fd0-p107">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="82fd0-167">Обработка событий</span><span class="sxs-lookup"><span data-stu-id="82fd0-167">Handle an event</span></span>

<span data-ttu-id="82fd0-p108">Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.</span><span class="sxs-lookup"><span data-stu-id="82fd0-p108">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="82fd0-171">Удаление обработчика события</span><span class="sxs-lookup"><span data-stu-id="82fd0-171">Remove an event handler</span></span>

<span data-ttu-id="82fd0-p109">В приведенном ниже примере кода регистрируется обработчик событий `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события. В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="82fd0-p109">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="82fd0-174">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="82fd0-174">Enable and disable events</span></span>

<span data-ttu-id="82fd0-175">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="82fd0-175">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="82fd0-176">Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей.</span><span class="sxs-lookup"><span data-stu-id="82fd0-176">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="82fd0-177">События включаются и отключаются на уровне [среды выполнения](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="82fd0-177">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="82fd0-178">Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики.</span><span class="sxs-lookup"><span data-stu-id="82fd0-178">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="82fd0-179">В приведенном ниже примере кода показано, как включать и отключать события.</span><span class="sxs-lookup"><span data-stu-id="82fd0-179">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="82fd0-180">См. также</span><span class="sxs-lookup"><span data-stu-id="82fd0-180">See also</span></span>

- [<span data-ttu-id="82fd0-181">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="82fd0-181">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

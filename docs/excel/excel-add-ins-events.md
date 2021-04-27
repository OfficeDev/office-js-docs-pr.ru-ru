---
title: Работа с событиями при помощи API JavaScript для Excel
description: Список событий для объектов JavaScript Excel. Это включает сведения об использовании обработчиков событий и связанных шаблонов.
ms.date: 04/23/2021
localization_priority: Normal
ms.openlocfilehash: 7c4a00736ef7a39e6d34153b113f1f274a2d3708
ms.sourcegitcommit: 78fb861afe7d7c3ee7fe3186150b3fed20994222
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2021
ms.locfileid: "52024313"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="67e1c-104">Работа с событиями при помощи API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="67e1c-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="67e1c-105">В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="67e1c-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="67e1c-106">События в Excel</span><span class="sxs-lookup"><span data-stu-id="67e1c-106">Events in Excel</span></span>

<span data-ttu-id="67e1c-p102">Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Ниже перечислены поддерживаемые в настоящее время события.</span><span class="sxs-lookup"><span data-stu-id="67e1c-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="67e1c-110">Событие</span><span class="sxs-lookup"><span data-stu-id="67e1c-110">Event</span></span> | <span data-ttu-id="67e1c-111">Описание</span><span class="sxs-lookup"><span data-stu-id="67e1c-111">Description</span></span> | <span data-ttu-id="67e1c-112">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="67e1c-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="67e1c-113">Возникает при активации объекта.</span><span class="sxs-lookup"><span data-stu-id="67e1c-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="67e1c-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span><span class="sxs-lookup"><span data-stu-id="67e1c-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span></span> |
| `onAdded` | <span data-ttu-id="67e1c-115">Возникает при добавлении объекта в коллекцию.</span><span class="sxs-lookup"><span data-stu-id="67e1c-115">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="67e1c-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span><span class="sxs-lookup"><span data-stu-id="67e1c-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="67e1c-117">Возникает при изменении параметра `autoSave` для книги.</span><span class="sxs-lookup"><span data-stu-id="67e1c-117">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="67e1c-118">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="67e1c-118">**Workbook**</span></span>](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | <span data-ttu-id="67e1c-119">Возникает после завершения вычислений на листе (или на всех листах коллекции).</span><span class="sxs-lookup"><span data-stu-id="67e1c-119">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="67e1c-120">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span><span class="sxs-lookup"><span data-stu-id="67e1c-120">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span></span> |
| `onChanged` | <span data-ttu-id="67e1c-121">Происходит, когда изменились данные отдельных ячеек или комментариев.</span><span class="sxs-lookup"><span data-stu-id="67e1c-121">Occurs when the data of individual cells or comments has changed.</span></span> | <span data-ttu-id="67e1c-122">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Таблица**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Таблица**](/javascript/api/excel/excel.worksheet#onchanged), Таблица , [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span><span class="sxs-lookup"><span data-stu-id="67e1c-122">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span></span> |
| `onColumnSorted` | <span data-ttu-id="67e1c-123">Возникает при сортировке одного или нескольких столбцов.</span><span class="sxs-lookup"><span data-stu-id="67e1c-123">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="67e1c-124">Происходит в результате операции сортировки слева направо.</span><span class="sxs-lookup"><span data-stu-id="67e1c-124">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="67e1c-125">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span><span class="sxs-lookup"><span data-stu-id="67e1c-125">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span></span> |
| `onDataChanged` | <span data-ttu-id="67e1c-126">Возникает при изменении данных или форматирования в привязке.</span><span class="sxs-lookup"><span data-stu-id="67e1c-126">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="67e1c-127">**Binding**</span><span class="sxs-lookup"><span data-stu-id="67e1c-127">**Binding**</span></span>](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | <span data-ttu-id="67e1c-128">Возникает при отключении объекта.</span><span class="sxs-lookup"><span data-stu-id="67e1c-128">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="67e1c-129">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="67e1c-129">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span></span> |
| `onDeleted` | <span data-ttu-id="67e1c-130">Возникает при удалении объекта из коллекции.</span><span class="sxs-lookup"><span data-stu-id="67e1c-130">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="67e1c-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span><span class="sxs-lookup"><span data-stu-id="67e1c-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span></span> |
| `onFormatChanged` | <span data-ttu-id="67e1c-132">Возникает при изменении формата на листе.</span><span class="sxs-lookup"><span data-stu-id="67e1c-132">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="67e1c-133">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span><span class="sxs-lookup"><span data-stu-id="67e1c-133">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span></span> |
| `onRowSorted` | <span data-ttu-id="67e1c-134">Возникает при сортировке одной или нескольких строк.</span><span class="sxs-lookup"><span data-stu-id="67e1c-134">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="67e1c-135">Происходит в результате операции сортировки сверху вниз.</span><span class="sxs-lookup"><span data-stu-id="67e1c-135">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="67e1c-136">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span><span class="sxs-lookup"><span data-stu-id="67e1c-136">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="67e1c-137">Возникает при изменении активной ячейки или выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="67e1c-137">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="67e1c-138">[**Привязка**](/javascript/api/excel/excel.binding#onselectionchanged), [**таблица**](/javascript/api/excel/excel.table#onselectionchanged), книга , [**таблица**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged) [](/javascript/api/excel/excel.workbook#onselectionchanged)</span><span class="sxs-lookup"><span data-stu-id="67e1c-138">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="67e1c-139">Возникает при изменении состояния скрытия строки на определенном листе.</span><span class="sxs-lookup"><span data-stu-id="67e1c-139">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="67e1c-140">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span><span class="sxs-lookup"><span data-stu-id="67e1c-140">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="67e1c-141">Возникает при изменении параметров в документе.</span><span class="sxs-lookup"><span data-stu-id="67e1c-141">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="67e1c-142">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="67e1c-142">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | <span data-ttu-id="67e1c-143">Возникает, когда происходит щелчок левой кнопкой мыши или нажатие на листе.</span><span class="sxs-lookup"><span data-stu-id="67e1c-143">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="67e1c-144">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span><span class="sxs-lookup"><span data-stu-id="67e1c-144">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span></span> |

### <a name="events-in-preview"></a><span data-ttu-id="67e1c-145">События в предварительной версии</span><span class="sxs-lookup"><span data-stu-id="67e1c-145">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="67e1c-146">Следующие события в настоящее время доступны только в общедоступной предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="67e1c-146">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="67e1c-147">Событие</span><span class="sxs-lookup"><span data-stu-id="67e1c-147">Event</span></span> | <span data-ttu-id="67e1c-148">Описание</span><span class="sxs-lookup"><span data-stu-id="67e1c-148">Description</span></span> | <span data-ttu-id="67e1c-149">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="67e1c-149">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="67e1c-150">Возникает при применении фильтра к объекту.</span><span class="sxs-lookup"><span data-stu-id="67e1c-150">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="67e1c-151">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span><span class="sxs-lookup"><span data-stu-id="67e1c-151">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="67e1c-152">Триггеры событий</span><span class="sxs-lookup"><span data-stu-id="67e1c-152">Event triggers</span></span>

<span data-ttu-id="67e1c-153">События в книге Excel могут вызываться:</span><span class="sxs-lookup"><span data-stu-id="67e1c-153">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="67e1c-154">при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="67e1c-154">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="67e1c-155">из кода (JavaScript) надстройки Office, вносящего изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="67e1c-155">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="67e1c-156">из кода (макроса) надстройки VBA, вносящего изменения в книгу.</span><span class="sxs-lookup"><span data-stu-id="67e1c-156">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="67e1c-157">Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.</span><span class="sxs-lookup"><span data-stu-id="67e1c-157">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="67e1c-158">Жизненный цикл обработчика событий</span><span class="sxs-lookup"><span data-stu-id="67e1c-158">Lifecycle of an event handler</span></span>

<span data-ttu-id="67e1c-159">Обработчик событий создается при его регистрации надстройкой.</span><span class="sxs-lookup"><span data-stu-id="67e1c-159">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="67e1c-160">Он удаляется при отмене его регистрации надстройкой или при обновлении, перезагрузке или закрытии надстройки.</span><span class="sxs-lookup"><span data-stu-id="67e1c-160">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="67e1c-161">Обработчики событий не остаются в составе файла Excel или между сеансами с интернет-версией Excel.</span><span class="sxs-lookup"><span data-stu-id="67e1c-161">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="67e1c-162">Когда объект, для которого зарегистрированы события, удаляется (например, таблица с зарегистрированным событием `onChanged`), обработчик событий больше не запускается, но остается в памяти, пока сеанс надстройки или Excel не обновится или не закроется.</span><span class="sxs-lookup"><span data-stu-id="67e1c-162">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="67e1c-163">События и совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="67e1c-163">Events and coauthoring</span></span>

<span data-ttu-id="67e1c-p107">Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, в частности `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="67e1c-p107">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="67e1c-166">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="67e1c-166">Register an event handler</span></span>

<span data-ttu-id="67e1c-p108">В приведенном ниже примере кода регистрируется обработчик события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleChange`.</span><span class="sxs-lookup"><span data-stu-id="67e1c-p108">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="67e1c-169">Обработка событий</span><span class="sxs-lookup"><span data-stu-id="67e1c-169">Handle an event</span></span>

<span data-ttu-id="67e1c-p109">Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.</span><span class="sxs-lookup"><span data-stu-id="67e1c-p109">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="67e1c-173">Удаление обработчика события</span><span class="sxs-lookup"><span data-stu-id="67e1c-173">Remove an event handler</span></span>

<span data-ttu-id="67e1c-174">В приведенном ниже примере кода регистрируется обработчик событий `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события.</span><span class="sxs-lookup"><span data-stu-id="67e1c-174">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="67e1c-175">В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="67e1c-175">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="67e1c-176">Обратите внимание, что для его удаления требуется использовать обработник `RequestContext` событий.</span><span class="sxs-lookup"><span data-stu-id="67e1c-176">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="67e1c-177">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="67e1c-177">Enable and disable events</span></span>

<span data-ttu-id="67e1c-178">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="67e1c-178">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="67e1c-179">Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей.</span><span class="sxs-lookup"><span data-stu-id="67e1c-179">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="67e1c-180">События включаются и отключаются на уровне [среды выполнения](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="67e1c-180">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="67e1c-181">Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики.</span><span class="sxs-lookup"><span data-stu-id="67e1c-181">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="67e1c-182">В приведенном ниже примере кода показано, как включать и отключать события.</span><span class="sxs-lookup"><span data-stu-id="67e1c-182">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="67e1c-183">См. также</span><span class="sxs-lookup"><span data-stu-id="67e1c-183">See also</span></span>

- [<span data-ttu-id="67e1c-184">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="67e1c-184">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)

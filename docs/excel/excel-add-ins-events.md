---
title: Работа с событиями при помощи API JavaScript для Excel
description: Список событий для объектов JavaScript Excel. Сюда входят сведения об использовании обработчиков событий и связанных с ними шаблонов.
ms.date: 03/31/2020
localization_priority: Normal
ms.openlocfilehash: 31f4123e812eae8a9a57abf7151b3500ce1ed296
ms.sourcegitcommit: 9c0e581b210eb31a7076496bbc7cd7d59192d799
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/03/2020
ms.locfileid: "43122903"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="a87d1-104">Работа с событиями при помощи API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a87d1-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="a87d1-105">В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="a87d1-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="a87d1-106">События в Excel</span><span class="sxs-lookup"><span data-stu-id="a87d1-106">Events in Excel</span></span>

<span data-ttu-id="a87d1-p102">Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Ниже перечислены поддерживаемые в настоящее время события.</span><span class="sxs-lookup"><span data-stu-id="a87d1-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="a87d1-110">Событие</span><span class="sxs-lookup"><span data-stu-id="a87d1-110">Event</span></span> | <span data-ttu-id="a87d1-111">Описание</span><span class="sxs-lookup"><span data-stu-id="a87d1-111">Description</span></span> | <span data-ttu-id="a87d1-112">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="a87d1-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="a87d1-113">Возникает при активации объекта.</span><span class="sxs-lookup"><span data-stu-id="a87d1-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="a87d1-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span><span class="sxs-lookup"><span data-stu-id="a87d1-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span></span> |
| `onAdded` | <span data-ttu-id="a87d1-115">Возникает при добавлении объекта в коллекцию.</span><span class="sxs-lookup"><span data-stu-id="a87d1-115">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="a87d1-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span><span class="sxs-lookup"><span data-stu-id="a87d1-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="a87d1-117">Возникает при изменении параметра `autoSave` для книги.</span><span class="sxs-lookup"><span data-stu-id="a87d1-117">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="a87d1-118">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="a87d1-118">**Workbook**</span></span>](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | <span data-ttu-id="a87d1-119">Возникает после завершения вычислений на листе (или на всех листах коллекции).</span><span class="sxs-lookup"><span data-stu-id="a87d1-119">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="a87d1-120">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span><span class="sxs-lookup"><span data-stu-id="a87d1-120">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span></span> |
| `onChanged` | <span data-ttu-id="a87d1-121">Возникает при изменении данных в ячейках.</span><span class="sxs-lookup"><span data-stu-id="a87d1-121">Occurs when data within cells is changed.</span></span> | <span data-ttu-id="a87d1-122">[**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span><span class="sxs-lookup"><span data-stu-id="a87d1-122">[**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span></span> |
| `onColumnSorted` | <span data-ttu-id="a87d1-123">Возникает при сортировке одного или нескольких столбцов.</span><span class="sxs-lookup"><span data-stu-id="a87d1-123">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="a87d1-124">Происходит в результате операции сортировки слева направо.</span><span class="sxs-lookup"><span data-stu-id="a87d1-124">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="a87d1-125">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span><span class="sxs-lookup"><span data-stu-id="a87d1-125">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span></span> |
| `onDataChanged` | <span data-ttu-id="a87d1-126">Возникает при изменении данных или форматирования в привязке.</span><span class="sxs-lookup"><span data-stu-id="a87d1-126">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="a87d1-127">**Binding**</span><span class="sxs-lookup"><span data-stu-id="a87d1-127">**Binding**</span></span>](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | <span data-ttu-id="a87d1-128">Возникает при отключении объекта.</span><span class="sxs-lookup"><span data-stu-id="a87d1-128">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="a87d1-129">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="a87d1-129">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span></span> |
| `onDeleted` | <span data-ttu-id="a87d1-130">Возникает при удалении объекта из коллекции.</span><span class="sxs-lookup"><span data-stu-id="a87d1-130">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="a87d1-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span><span class="sxs-lookup"><span data-stu-id="a87d1-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span></span> |
| `onFormatChanged` | <span data-ttu-id="a87d1-132">Возникает при изменении формата на листе.</span><span class="sxs-lookup"><span data-stu-id="a87d1-132">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="a87d1-133">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span><span class="sxs-lookup"><span data-stu-id="a87d1-133">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span></span> |
| `onRowSorted` | <span data-ttu-id="a87d1-134">Возникает при сортировке одной или нескольких строк.</span><span class="sxs-lookup"><span data-stu-id="a87d1-134">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="a87d1-135">Происходит в результате операции сортировки сверху вниз.</span><span class="sxs-lookup"><span data-stu-id="a87d1-135">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="a87d1-136">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span><span class="sxs-lookup"><span data-stu-id="a87d1-136">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="a87d1-137">Возникает при изменении активной ячейки или выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="a87d1-137">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="a87d1-138">[**Привязка**](/javascript/api/excel/excel.binding#onselectionchanged), [**Таблица**](/javascript/api/excel/excel.table#onselectionchanged), [**Книга**](/javascript/api/excel/excel.workbook#onselectionchanged), [**лист**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span><span class="sxs-lookup"><span data-stu-id="a87d1-138">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="a87d1-139">Возникает при изменении параметров в документе.</span><span class="sxs-lookup"><span data-stu-id="a87d1-139">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="a87d1-140">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="a87d1-140">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | <span data-ttu-id="a87d1-141">Возникает, когда происходит щелчок левой кнопкой мыши или нажатие на листе.</span><span class="sxs-lookup"><span data-stu-id="a87d1-141">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="a87d1-142">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span><span class="sxs-lookup"><span data-stu-id="a87d1-142">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span></span> |

> [!WARNING]
> <span data-ttu-id="a87d1-143">Событие `onSelectionChanged` в настоящее время нестабильно.</span><span class="sxs-lookup"><span data-stu-id="a87d1-143">`onSelectionChanged` is currently unstable.</span></span> <span data-ttu-id="a87d1-144">Существует временное решение для надежного использования `onSelectionChanged`.</span><span class="sxs-lookup"><span data-stu-id="a87d1-144">There is a workaround to reliably use `onSelectionChanged`.</span></span> <span data-ttu-id="a87d1-145">Добавьте следующий код в раздел `<head>` своей главной страницы HTML:</span><span class="sxs-lookup"><span data-stu-id="a87d1-145">Add the following code to the `<head>` section of your HTML home page:</span></span>
>
> ```HTML
> <script> MutationObserver=null; </script>
> ```
>
> <span data-ttu-id="a87d1-146">Полное обсуждение проблемы находится в [репозитории GitHub office-js](https://github.com/OfficeDev/office-js/issues/533).</span><span class="sxs-lookup"><span data-stu-id="a87d1-146">A full discussion of the issue can be found on the [office-js GitHub repo](https://github.com/OfficeDev/office-js/issues/533).</span></span>

### <a name="events-in-preview"></a><span data-ttu-id="a87d1-147">События в предварительной версии</span><span class="sxs-lookup"><span data-stu-id="a87d1-147">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="a87d1-148">Следующие события в настоящее время доступны только в общедоступной предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="a87d1-148">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="a87d1-149">Событие</span><span class="sxs-lookup"><span data-stu-id="a87d1-149">Event</span></span> | <span data-ttu-id="a87d1-150">Описание</span><span class="sxs-lookup"><span data-stu-id="a87d1-150">Description</span></span> | <span data-ttu-id="a87d1-151">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="a87d1-151">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="a87d1-152">Возникает при применении фильтра к объекту.</span><span class="sxs-lookup"><span data-stu-id="a87d1-152">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="a87d1-153">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span><span class="sxs-lookup"><span data-stu-id="a87d1-153">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="a87d1-154">Возникает при изменении состояния скрытия строки на определенном листе.</span><span class="sxs-lookup"><span data-stu-id="a87d1-154">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="a87d1-155">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span><span class="sxs-lookup"><span data-stu-id="a87d1-155">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="a87d1-156">Триггеры событий</span><span class="sxs-lookup"><span data-stu-id="a87d1-156">Event triggers</span></span>

<span data-ttu-id="a87d1-157">События в книге Excel могут вызываться:</span><span class="sxs-lookup"><span data-stu-id="a87d1-157">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="a87d1-158">при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="a87d1-158">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="a87d1-159">из кода (JavaScript) надстройки Office, вносящего изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="a87d1-159">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="a87d1-160">из кода (макроса) надстройки VBA, вносящего изменения в книгу.</span><span class="sxs-lookup"><span data-stu-id="a87d1-160">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="a87d1-161">Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.</span><span class="sxs-lookup"><span data-stu-id="a87d1-161">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="a87d1-162">Жизненный цикл обработчика событий</span><span class="sxs-lookup"><span data-stu-id="a87d1-162">Lifecycle of an event handler</span></span>

<span data-ttu-id="a87d1-163">Обработчик событий создается при его регистрации надстройкой.</span><span class="sxs-lookup"><span data-stu-id="a87d1-163">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="a87d1-164">Он удаляется при отмене его регистрации надстройкой или при обновлении, перезагрузке или закрытии надстройки.</span><span class="sxs-lookup"><span data-stu-id="a87d1-164">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="a87d1-165">Обработчики событий не остаются в составе файла Excel или между сеансами с интернет-версией Excel.</span><span class="sxs-lookup"><span data-stu-id="a87d1-165">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="a87d1-166">Когда объект, для которого зарегистрированы события, удаляется (например, таблица с зарегистрированным событием `onChanged`), обработчик событий больше не запускается, но остается в памяти, пока сеанс надстройки или Excel не обновится или не закроется.</span><span class="sxs-lookup"><span data-stu-id="a87d1-166">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="a87d1-167">События и совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="a87d1-167">Events and coauthoring</span></span>

<span data-ttu-id="a87d1-p108">Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, в частности `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="a87d1-p108">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="a87d1-170">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="a87d1-170">Register an event handler</span></span>

<span data-ttu-id="a87d1-p109">В приведенном ниже примере кода регистрируется обработчик события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.</span><span class="sxs-lookup"><span data-stu-id="a87d1-p109">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="a87d1-173">Обработка событий</span><span class="sxs-lookup"><span data-stu-id="a87d1-173">Handle an event</span></span>

<span data-ttu-id="a87d1-p110">Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.</span><span class="sxs-lookup"><span data-stu-id="a87d1-p110">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="a87d1-177">Удаление обработчика события</span><span class="sxs-lookup"><span data-stu-id="a87d1-177">Remove an event handler</span></span>

<span data-ttu-id="a87d1-178">В приведенном ниже примере кода регистрируется обработчик событий `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события.</span><span class="sxs-lookup"><span data-stu-id="a87d1-178">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="a87d1-179">В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="a87d1-179">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="a87d1-180">Обратите внимание `RequestContext` , что для удаления обработчика событий необходимо, чтобы он использовался для создания обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="a87d1-180">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="a87d1-181">Включение и отключение событий</span><span class="sxs-lookup"><span data-stu-id="a87d1-181">Enable and disable events</span></span>

<span data-ttu-id="a87d1-182">Производительность надстройки можно повысить с помощью отключения событий.</span><span class="sxs-lookup"><span data-stu-id="a87d1-182">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="a87d1-183">Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей.</span><span class="sxs-lookup"><span data-stu-id="a87d1-183">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="a87d1-184">События включаются и отключаются на уровне [среды выполнения](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="a87d1-184">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="a87d1-185">Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики.</span><span class="sxs-lookup"><span data-stu-id="a87d1-185">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="a87d1-186">В приведенном ниже примере кода показано, как включать и отключать события.</span><span class="sxs-lookup"><span data-stu-id="a87d1-186">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a87d1-187">См. также</span><span class="sxs-lookup"><span data-stu-id="a87d1-187">See also</span></span>

- [<span data-ttu-id="a87d1-188">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a87d1-188">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

---
title: Работа с событиями при помощи API JavaScript для Excel
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: 575e4112ed5f55356020eed8327d309fc58cd643
ms.sourcegitcommit: 9685fd83136bd2106f4c5595bda0010bc1b1950b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/19/2018
ms.locfileid: "20596521"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="57f2d-102">Работа с событиями при помощи API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="57f2d-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="57f2d-103">В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="57f2d-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="57f2d-104">События в Excel</span><span class="sxs-lookup"><span data-stu-id="57f2d-104">Events in Excel</span></span>

<span data-ttu-id="57f2d-105">Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии.</span><span class="sxs-lookup"><span data-stu-id="57f2d-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="57f2d-106">С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события.</span><span class="sxs-lookup"><span data-stu-id="57f2d-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="57f2d-107">Ниже перечислены поддерживаемые в настоящее время события.</span><span class="sxs-lookup"><span data-stu-id="57f2d-107">The following events are currently supported.</span></span>

| <span data-ttu-id="57f2d-108">Событие</span><span class="sxs-lookup"><span data-stu-id="57f2d-108">Event</span></span> | <span data-ttu-id="57f2d-109">Описание</span><span class="sxs-lookup"><span data-stu-id="57f2d-109">Description</span></span> | <span data-ttu-id="57f2d-110">Поддерживаемые объекты</span><span class="sxs-lookup"><span data-stu-id="57f2d-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="57f2d-111">Событие, возникающее при добавлении объекта.</span><span class="sxs-lookup"><span data-stu-id="57f2d-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="57f2d-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="57f2d-112">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | <span data-ttu-id="57f2d-113">Событие, возникающее при удалении объекта.</span><span class="sxs-lookup"><span data-stu-id="57f2d-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="57f2d-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="57f2d-114">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | <span data-ttu-id="57f2d-115">Событие, возникающее при активации объекта.</span><span class="sxs-lookup"><span data-stu-id="57f2d-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="57f2d-116">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Лист**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="57f2d-116">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="57f2d-117">Событие, возникающее при отключении объекта.</span><span class="sxs-lookup"><span data-stu-id="57f2d-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="57f2d-118">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Лист**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="57f2d-118">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="57f2d-119">Событие, возникающее при изменении данных в ячейках.</span><span class="sxs-lookup"><span data-stu-id="57f2d-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="57f2d-120">[**Лист**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Таблица**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span><span class="sxs-lookup"><span data-stu-id="57f2d-120">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **TableCollection**, [Binding](https://dev.office.com/reference/add-ins/excel/table)</span></span> |
| `onDataChanged` | <span data-ttu-id="57f2d-121">Событие, возникающее при изменении данных или форматирования в привязке.</span><span class="sxs-lookup"><span data-stu-id="57f2d-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="57f2d-122">**Привязка**</span><span class="sxs-lookup"><span data-stu-id="57f2d-122">**Binding**</span></span>](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | <span data-ttu-id="57f2d-123">Событие, возникающее при изменении активной ячейки или выбранного диапазона.</span><span class="sxs-lookup"><span data-stu-id="57f2d-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="57f2d-124">[**Лист**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Таблица**](https://dev.office.com/reference/add-ins/excel/table), [**Привязка**](https://dev.office.com/reference/add-ins/excel/binding)</span><span class="sxs-lookup"><span data-stu-id="57f2d-124">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="57f2d-125">Событие, возникающее при изменении Параметров в документе.</span><span class="sxs-lookup"><span data-stu-id="57f2d-125">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="57f2d-126">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="57f2d-126">**SettingCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="57f2d-127">Триггеры событий</span><span class="sxs-lookup"><span data-stu-id="57f2d-127">Event triggers</span></span>

<span data-ttu-id="57f2d-128">События в книге Excel могут вызываться:</span><span class="sxs-lookup"><span data-stu-id="57f2d-128">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="57f2d-129">при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="57f2d-129">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="57f2d-130">из кода (JavaScript) надстройки Office, вносящего изменения в книгу;</span><span class="sxs-lookup"><span data-stu-id="57f2d-130">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="57f2d-131">из кода (макроса) надстройки VBA, вносящего изменения в книгу.</span><span class="sxs-lookup"><span data-stu-id="57f2d-131">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="57f2d-132">Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.</span><span class="sxs-lookup"><span data-stu-id="57f2d-132">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="57f2d-133">Жизненный цикл обработчика событий</span><span class="sxs-lookup"><span data-stu-id="57f2d-133">Lifecycle of an event handler</span></span>

<span data-ttu-id="57f2d-p102">Обработчик событий создается, когда надстройка регистрирует его, и удаляется при отмене его регистрации или закрытии надстройки. Обработчики событий не остаются в составе файла Excel.</span><span class="sxs-lookup"><span data-stu-id="57f2d-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="57f2d-136">События и совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="57f2d-136">Events and coauthoring</span></span>

<span data-ttu-id="57f2d-p103">Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, в частности `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="57f2d-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="57f2d-139">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="57f2d-139">Register an event handler</span></span>

<span data-ttu-id="57f2d-140">В приведенном ниже примере кода регистрируется обработчик события `onChanged` на листе под названием **Sample**.</span><span class="sxs-lookup"><span data-stu-id="57f2d-140">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="57f2d-141">В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleDataChange`.</span><span class="sxs-lookup"><span data-stu-id="57f2d-141">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="57f2d-142">Обработка событий</span><span class="sxs-lookup"><span data-stu-id="57f2d-142">Handle an event</span></span>

<span data-ttu-id="57f2d-143">Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события.</span><span class="sxs-lookup"><span data-stu-id="57f2d-143">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="57f2d-144">Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария.</span><span class="sxs-lookup"><span data-stu-id="57f2d-144">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="57f2d-145">В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.</span><span class="sxs-lookup"><span data-stu-id="57f2d-145">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="57f2d-146">Удаление обработчика события</span><span class="sxs-lookup"><span data-stu-id="57f2d-146">Remove an event handler</span></span>

<span data-ttu-id="57f2d-147">В приведенном ниже примере кода регистрируется обработчик событий `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события.</span><span class="sxs-lookup"><span data-stu-id="57f2d-147">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="57f2d-148">В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="57f2d-148">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="57f2d-149">См. также</span><span class="sxs-lookup"><span data-stu-id="57f2d-149">See also</span></span>

- [<span data-ttu-id="57f2d-150">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="57f2d-150">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="57f2d-151">Открытая спецификация по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="57f2d-151">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
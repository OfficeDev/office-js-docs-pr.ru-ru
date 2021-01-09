---
title: Показать или скрыть области задач надстройки Office
description: Узнайте, как программным образом скрывать или показывать пользовательский интерфейс надстройки во время ее непрерывной работы.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789250"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a><span data-ttu-id="1c8c0-103">Показать или скрыть области задач надстройки Office</span><span class="sxs-lookup"><span data-stu-id="1c8c0-103">Show or hide the task pane of your Office Add-in</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="1c8c0-104">Вы можете отдемонстрировать области задач надстройки Office, вызывая `Office.addin.showAsTaskpane()` функцию.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-104">You can show the task pane of your Office Add-in by calling the `Office.addin.showAsTaskpane()` function.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="1c8c0-105">В предыдущем коде предполагается, что существует таблица Excel **с именем CurrentQuarterSales.**</span><span class="sxs-lookup"><span data-stu-id="1c8c0-105">The previous code assumes a scenario where there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="1c8c0-106">Надстройка делает области задач видимыми при активации этого таблицы.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-106">The add-in will make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="1c8c0-107">Этот метод `onCurrentQuarter` является обработом события [Office.Worksheet.onActivated,](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) зарегистрированного для этого таблицы.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-107">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) event which has been registered for the worksheet.</span></span>

<span data-ttu-id="1c8c0-108">Вы также можете скрыть области задач, вызывая `Office.addin.hide()` функцию.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-108">You can also hide the task pane by calling the `Office.addin.hide()` function.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

<span data-ttu-id="1c8c0-109">Предыдущий код — это обработец, зарегистрированный для [события Office.Worksheet.onDeactivated.](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="1c8c0-109">The previous code is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) event.</span></span>

## <a name="additional-details-on-showing-the-task-pane"></a><span data-ttu-id="1c8c0-110">Дополнительные сведения о от показании области задач</span><span class="sxs-lookup"><span data-stu-id="1c8c0-110">Additional details on showing the task pane</span></span>

<span data-ttu-id="1c8c0-111">При вызове Office отобразит в области задач файл, который назначен в качестве значения ресурса `Office.addin.showAsTaskpane()` ( `resid` ) области задач.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-111">When you call `Office.addin.showAsTaskpane()`, Office will display in a task pane the file that you assigned as the resource ID (`resid`) value of the task pane.</span></span> <span data-ttu-id="1c8c0-112">Это `resid` значение можно навести или  изменить, открыв файлmanifest.xmlи выявив его `<SourceLocation>` внутри `<Action xsi:type="ShowTaskpane">` элемента.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-112">This `resid` value can be assigned or changed by opening your **manifest.xml** file and locating `<SourceLocation>` inside the `<Action xsi:type="ShowTaskpane">` element.</span></span>
<span data-ttu-id="1c8c0-113">[(Дополнительные сведения см.](configure-your-add-in-to-use-a-shared-runtime.md) в настройках надстройки Office для использования общей времени работы.)</span><span class="sxs-lookup"><span data-stu-id="1c8c0-113">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md) for additional details.)</span></span>

<span data-ttu-id="1c8c0-114">Так `Office.addin.showAsTaskpane()` как это асинхронный метод, код будет работать до завершения работы функции.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-114">Since `Office.addin.showAsTaskpane()` is an asynchronous method, your code will continue running until the function is complete.</span></span> <span data-ttu-id="1c8c0-115">Дождись завершения с помощью ключевого слова или метода в зависимости от используемого синтаксиса `await` `then()` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-115">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span>

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a><span data-ttu-id="1c8c0-116">Настройка надстройки для использования общей времени работы</span><span class="sxs-lookup"><span data-stu-id="1c8c0-116">Configure your add-in to use the shared runtime</span></span>

<span data-ttu-id="1c8c0-117">Для использования этих `showAsTaskpane()` `hide()` методов надстройка должна использовать общую времени работы.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-117">To use the `showAsTaskpane()` and `hide()` methods, your add-in must use the shared runtime.</span></span> <span data-ttu-id="1c8c0-118">Дополнительные сведения см. в настройках [надстройки Office для использования общей времени работы.](configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="1c8c0-118">For more information, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="1c8c0-119">Сохранение прослушивателей состояния и событий</span><span class="sxs-lookup"><span data-stu-id="1c8c0-119">Preservation of state and event listeners</span></span>

<span data-ttu-id="1c8c0-120">Методы `hide()` `showAsTaskpane()` и методы изменяют только *видимость* области задач.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-120">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="1c8c0-121">Они не выгружают и не перезагружают его (или повторно ициализируют его состояние).</span><span class="sxs-lookup"><span data-stu-id="1c8c0-121">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="1c8c0-122">Рассмотрим следующий сценарий: в области задач есть вкладки.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-122">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="1c8c0-123">Вкладка **"Главная"** открывается при первом запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-123">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="1c8c0-124">Предположим, что  пользователь открывает вкладку "Параметры", а затем код в области задач вызывается в ответ `hide()` на какое-либо событие.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-124">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="1c8c0-125">Тем не менее позже код `showAsTaskpane()` вызывается в ответ на другое событие.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-125">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="1c8c0-126">В области задач снова появится  вкладка "Параметры".</span><span class="sxs-lookup"><span data-stu-id="1c8c0-126">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![Снимок экрана области задач с четырьмя вкладками "Главная", "Параметры", "Избранное" и "Учетные записи".](../images/TaskpaneWithTabs.png)

<span data-ttu-id="1c8c0-128">Кроме того, все прослушиватели событий, зарегистрированные в области задач, продолжают работать даже при скрытии области задач.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-128">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="1c8c0-129">Рассмотрим следующий сценарий: в области задач есть зарегистрированный обработитель для Excel и события для `Worksheet.onActivated` `Worksheet.onDeactivated` листа **Sheet1.**</span><span class="sxs-lookup"><span data-stu-id="1c8c0-129">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="1c8c0-130">Активированный обработок приводит к появления зеленой точки в области задач.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-130">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="1c8c0-131">Деактивированный обработок включает красный цвет точки (которое является состоянием по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="1c8c0-131">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="1c8c0-132">Предположим, что код `hide()` вызывается, **когда Лист1** не активирован, а точка является красной.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-132">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="1c8c0-133">Несмотря на то что области задач скрыты, **лист1** активируется.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-133">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="1c8c0-134">Последующие вызовы кода `showAsTaskpane()` в ответ на какое-либо событие.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-134">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="1c8c0-135">Когда откроется области задач, точка будет зеленой, так как прослушиватели событий и обработчики запустились, даже если она была скрыта.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-135">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

## <a name="handle-the-visibility-changed-event"></a><span data-ttu-id="1c8c0-136">Обработка события изменения видимости</span><span class="sxs-lookup"><span data-stu-id="1c8c0-136">Handle the visibility changed event</span></span>

<span data-ttu-id="1c8c0-137">Когда код изменяет видимость области задач с помощью или `showAsTaskpane()` `hide()` , Office активирует `VisibilityModeChanged` событие.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-137">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="1c8c0-138">Это событие может оказаться полезным.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-138">It can be useful to handle this event.</span></span> <span data-ttu-id="1c8c0-139">Например, предположим, что в области задач отображается список всех листов в книге.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-139">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="1c8c0-140">Если новый лист добавляется при скрытии области задач, то, чтобы сделать ее видимой, она сама по себе не добавит новое имя в список.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-140">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="1c8c0-141">Но ваш код может реагировать на событие, чтобы перезагрузить Worksheet.name всех таблиц в коллекции `VisibilityModeChanged` [Workbook.worksheets,](/javascript/api/excel/excel.workbook#worksheets) как показано в примере кода ниже. [](/javascript/api/excel/excel.worksheet#name)</span><span class="sxs-lookup"><span data-stu-id="1c8c0-141">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="1c8c0-142">Чтобы зарегистрировать обработатель для события, не используйте метод add handler, как в большинстве контекстов JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-142">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="1c8c0-143">Вместо этого существует специальная функция, которой вы передаете обработчивую функцию: [Office.addin.onVisibilityModeChanged.](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)</span><span class="sxs-lookup"><span data-stu-id="1c8c0-143">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="1c8c0-144">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-144">The following is an example.</span></span> <span data-ttu-id="1c8c0-145">Обратите `args.visibilityMode` внимание, что свойство имеет тип [VisibilityMode.](/javascript/api/office/office.visibilitymode)</span><span class="sxs-lookup"><span data-stu-id="1c8c0-145">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="1c8c0-146">Функция возвращает другую функцию, которая *отрегистрировать* обработитель.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-146">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="1c8c0-147">Вот простой, но не надежный пример:</span><span class="sxs-lookup"><span data-stu-id="1c8c0-147">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="1c8c0-148">Метод является асинхронным и возвращает обещание, которое означает, что ваш код должен ожидать выполнения обещания, прежде чем он сможет вызвать обработитель `onVisibilityModeChanged` дерегистрации. </span><span class="sxs-lookup"><span data-stu-id="1c8c0-148">The `onVisibilityModeChanged` method is asynchronous and returns a promise, which means that your code needs to await the fulfillment of the promise before it can call the **deregister** handler.</span></span>

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="1c8c0-149">Функция дерегистрации также является асинхронной и возвращает обещание.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-149">The deregister function is also asynchronous and returns a promise.</span></span> <span data-ttu-id="1c8c0-150">Таким образом, если у вас есть код, который не должен запускаться до завершения регистрации, следует дождаться обещания, возвращенного функцией дерегистрации.</span><span class="sxs-lookup"><span data-stu-id="1c8c0-150">So, if you have code that should not run until after the deregistration is complete, then you should await the promise returned by the deregister function.</span></span>

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a><span data-ttu-id="1c8c0-151">См. также</span><span class="sxs-lookup"><span data-stu-id="1c8c0-151">See also</span></span>

- [<span data-ttu-id="1c8c0-152">Настройка надстройки Office для использования общей времени работы JavaScript</span><span class="sxs-lookup"><span data-stu-id="1c8c0-152">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="1c8c0-153">Запуск кода в надстройки Office при запуске документа</span><span class="sxs-lookup"><span data-stu-id="1c8c0-153">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)

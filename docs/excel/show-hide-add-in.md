---
title: Отображение и скрытие надстройки Office в общей среде выполнения
description: Сведения о том, как программно скрыть или отобразить пользовательский интерфейс надстройки, когда он работает постоянно
ms.date: 05/17/2020
localization_priority: Normal
ms.openlocfilehash: 9b6c3384fda32854e26cc4852d5bd27d77fae544
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610335"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime"></a><span data-ttu-id="eeff1-103">Отображение и скрытие надстройки Office в общей среде выполнения</span><span class="sxs-lookup"><span data-stu-id="eeff1-103">Show or hide an Office Add-in in a shared runtime</span></span>

<span data-ttu-id="eeff1-104">Надстройка Office может включать любые из следующих частей:</span><span class="sxs-lookup"><span data-stu-id="eeff1-104">An Office Add-in can include any of the following parts:</span></span>

- <span data-ttu-id="eeff1-105">Область задач</span><span class="sxs-lookup"><span data-stu-id="eeff1-105">A task pane</span></span>
- <span data-ttu-id="eeff1-106">Файл функции без пользовательского интерфейса (пользовательские функции, которые не используют область задач или другие элементы пользовательского интерфейса)</span><span class="sxs-lookup"><span data-stu-id="eeff1-106">A UI-less function file (custom functions which do not use a task pane or other user interface elements)</span></span>
- <span data-ttu-id="eeff1-107">Пользовательская функция Excel</span><span class="sxs-lookup"><span data-stu-id="eeff1-107">An Excel custom function</span></span>

<span data-ttu-id="eeff1-108">По умолчанию каждая часть выполняется в отдельной среде выполнения JavaScript с собственным глобальным объектом и глобальными переменными.</span><span class="sxs-lookup"><span data-stu-id="eeff1-108">By default, each part runs in its own separate JavaScript runtime, with its own global object and global variables.</span></span>

<span data-ttu-id="eeff1-109">Надстройки могут совместно использовать общую среду выполнения JavaScript с двумя или более частями.</span><span class="sxs-lookup"><span data-stu-id="eeff1-109">It's possible for add-ins with two or more parts to share a common JavaScript runtime.</span></span> <span data-ttu-id="eeff1-110">Эта общая функция среды выполнения включает новые API, которые скрывают и повторно открывают область задач во время выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="eeff1-110">This shared runtime feature enables new APIs that hide and reopen the task pane while the add-in runs.</span></span>

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="eeff1-111">Настройка надстройки для использования общей среды выполнения</span><span class="sxs-lookup"><span data-stu-id="eeff1-111">Configure an add-in to use a shared runtime</span></span>

<span data-ttu-id="eeff1-112">Чтобы настроить надстройку для использования общей среды выполнения, ознакомьтесь со статьей [Настройка надстройки Office для использования общей среды выполнения](configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="eeff1-112">To configure the add-in to use a shared runtime, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="show-and-hide-the-task-pane"></a><span data-ttu-id="eeff1-113">Отображение и скрытие области задач</span><span class="sxs-lookup"><span data-stu-id="eeff1-113">Show and hide the task pane</span></span>

<span data-ttu-id="eeff1-114">Новые API находятся в `Office.addin` свойстве.</span><span class="sxs-lookup"><span data-stu-id="eeff1-114">The new APIs are in the `Office.addin` property.</span></span> <span data-ttu-id="eeff1-115">Чтобы отобразить область задач, вызывается код `Office.addin.showAsTaskpane()` .</span><span class="sxs-lookup"><span data-stu-id="eeff1-115">To show the task pane, your code calls `Office.addin.showAsTaskpane()`.</span></span> <span data-ttu-id="eeff1-116">В области задач Office будет отображаться страница, которая была назначена ИДЕНТИФИКАТОРу ресурса ( `resid` ) для области задач.</span><span class="sxs-lookup"><span data-stu-id="eeff1-116">Office will display in a task pane the page that you assigned to the resource ID (`resid`) for the task pane.</span></span> <span data-ttu-id="eeff1-117">Это то `resid` , которое было назначено в `<SourceLocation>` `<Action xsi:type="ShowTaskpane">` манифесте.</span><span class="sxs-lookup"><span data-stu-id="eeff1-117">This is the `resid` that you assigned to the `<SourceLocation>` of the `<Action xsi:type="ShowTaskpane">` in the manifest.</span></span> <span data-ttu-id="eeff1-118">(См. [Настройка надстройки Office для использования совместно используемой среды выполнения](configure-your-add-in-to-use-a-shared-runtime.md)).</span><span class="sxs-lookup"><span data-stu-id="eeff1-118">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).)</span></span>

<span data-ttu-id="eeff1-119">Это асинхронный метод, поэтому код должен ожидать его, если следующий код не будет выполняться, пока он не будет завершен.</span><span class="sxs-lookup"><span data-stu-id="eeff1-119">This is an asynchronous method, so your code should await it when the subsequent code should not run until it completes.</span></span> <span data-ttu-id="eeff1-120">Дождитесь этого завершения с помощью `await` ключевого слова или метода в зависимости от используемого `then()` синтаксиса JavaScript.</span><span class="sxs-lookup"><span data-stu-id="eeff1-120">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span> <span data-ttu-id="eeff1-121">Ниже предполагается, что имеется лист Excel с именем **курренткуартерсалес**.</span><span class="sxs-lookup"><span data-stu-id="eeff1-121">The following assumes that there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="eeff1-122">Надстройка должна сделать область задач видимой при активации этого листа.</span><span class="sxs-lookup"><span data-stu-id="eeff1-122">The add-in should make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="eeff1-123">Метод `onCurrentQuarter` является обработчиком для события [Office. лист. OnActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) , зарегистрированного для листа.</span><span class="sxs-lookup"><span data-stu-id="eeff1-123">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) event which has been registered for the worksheet.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="eeff1-124">Чтобы скрыть область задач, вызывается код `Office.addin.hide()` .</span><span class="sxs-lookup"><span data-stu-id="eeff1-124">To hide the task pane, your code calls `Office.addin.hide()`.</span></span> <span data-ttu-id="eeff1-125">В следующем примере показан обработчик, зарегистрированный для события [Office. лист. OnDeactivate](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) .</span><span class="sxs-lookup"><span data-stu-id="eeff1-125">The following example is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) event.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="eeff1-126">Сохранение прослушивателей состояний и событий</span><span class="sxs-lookup"><span data-stu-id="eeff1-126">Preservation of state and event listeners</span></span>

<span data-ttu-id="eeff1-127">`hide()`Методы and `showAsTaskpane()` изменяют только *видимость* области задач.</span><span class="sxs-lookup"><span data-stu-id="eeff1-127">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="eeff1-128">Они не выгружают и не загружают их (или повторно инициализируют состояние).</span><span class="sxs-lookup"><span data-stu-id="eeff1-128">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="eeff1-129">Рассмотрим следующий сценарий: область задач разработана с использованием вкладок.</span><span class="sxs-lookup"><span data-stu-id="eeff1-129">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="eeff1-130">Вкладка **Главная** открывается при первом запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="eeff1-130">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="eeff1-131">Предположим, что пользователь открывает вкладку **Параметры** , а затем код в области задач вызывается `hide()` в ответ на событие.</span><span class="sxs-lookup"><span data-stu-id="eeff1-131">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="eeff1-132">Все еще позже вызывается код `showAsTaskpane()` в ответ на другое событие.</span><span class="sxs-lookup"><span data-stu-id="eeff1-132">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="eeff1-133">Область задач будет снова отображаться, а вкладка **Параметры** все еще будет выбрана.</span><span class="sxs-lookup"><span data-stu-id="eeff1-133">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![Снимок экрана с областью задач с четырьмя вкладками "Главная", "Параметры", "Избранное" и "учетные записи".](../images/TaskpaneWithTabs.png)

<span data-ttu-id="eeff1-135">Кроме того, все прослушиватели событий, зарегистрированные в области задач, продолжают выполняться, даже если область задач скрыта.</span><span class="sxs-lookup"><span data-stu-id="eeff1-135">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="eeff1-136">Рассмотрим следующий сценарий: область задач содержит зарегистрированный обработчик для Excel `Worksheet.onActivated` и `Worksheet.onDeactivated` событий для листа с именем **Лист1**.</span><span class="sxs-lookup"><span data-stu-id="eeff1-136">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="eeff1-137">Активированный обработчик вызывает отображение зеленой точки в области задач.</span><span class="sxs-lookup"><span data-stu-id="eeff1-137">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="eeff1-138">Отключенный обработчик включает красную точку (ее состояние по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="eeff1-138">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="eeff1-139">Предположим, что код вызывается, `hide()` когда **Лист1** не активирован, а точка имеет красный цвет.</span><span class="sxs-lookup"><span data-stu-id="eeff1-139">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="eeff1-140">Когда область задач скрыта, **Лист1** активируется.</span><span class="sxs-lookup"><span data-stu-id="eeff1-140">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="eeff1-141">Последующие вызовы кода `showAsTaskpane()` в ответ на событие.</span><span class="sxs-lookup"><span data-stu-id="eeff1-141">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="eeff1-142">Когда откроется область задач, точка будет зеленым, так как прослушиватели и обработчики событий запускаются несмотря на то, что область задач скрыта.</span><span class="sxs-lookup"><span data-stu-id="eeff1-142">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

### <a name="handle-visibility-changed-event"></a><span data-ttu-id="eeff1-143">Обработка события изменения видимости</span><span class="sxs-lookup"><span data-stu-id="eeff1-143">Handle visibility changed event</span></span>

<span data-ttu-id="eeff1-144">Когда код изменяет видимость области задач с `showAsTaskpane()` или `hide()` , Office запускает `VisibilityModeChanged` событие.</span><span class="sxs-lookup"><span data-stu-id="eeff1-144">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="eeff1-145">Это может пригодиться для обработки этого события.</span><span class="sxs-lookup"><span data-stu-id="eeff1-145">It can be useful to handle this event.</span></span> <span data-ttu-id="eeff1-146">Например, предположим, что в области задач отображается список всех листов в книге.</span><span class="sxs-lookup"><span data-stu-id="eeff1-146">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="eeff1-147">Если новый лист добавляется в то время, когда область задач скрыта, то при отображении видимой области задач в списке добавляется имя нового листа.</span><span class="sxs-lookup"><span data-stu-id="eeff1-147">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="eeff1-148">Но ваш код может ответить на `VisibilityModeChanged` событие, чтобы перегрузить свойство [Worksheet.Name](/javascript/api/excel/excel.worksheet#name) для всех листов в коллекции [Workbook. листы](/javascript/api/excel/excel.workbook#worksheets) , как показано в приведенном ниже примере кода.</span><span class="sxs-lookup"><span data-stu-id="eeff1-148">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="eeff1-149">Чтобы зарегистрировать обработчик для события, не используйте метод "добавить обработчик", как в большинстве контекстов Office JavaScript.</span><span class="sxs-lookup"><span data-stu-id="eeff1-149">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="eeff1-150">Вместо этого используется специальная функция, для которой вы передаете свой обработчик: [Office. AddIn. онвисибилитимодечанжед](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span><span class="sxs-lookup"><span data-stu-id="eeff1-150">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="eeff1-151">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="eeff1-151">The following is an example.</span></span> <span data-ttu-id="eeff1-152">Обратите внимание, что `args.visibilityMode` свойство имеет тип [висибилитимоде](/javascript/api/office/office.visibilitymode).</span><span class="sxs-lookup"><span data-stu-id="eeff1-152">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="eeff1-153">Функция возвращает другую функцию, которая *отменяет регистрацию* обработчика.</span><span class="sxs-lookup"><span data-stu-id="eeff1-153">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="eeff1-154">Вот простой, но не надежный, пример:</span><span class="sxs-lookup"><span data-stu-id="eeff1-154">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="eeff1-155">`onVisibilityModeChanged`Метод является асинхронным, что означает, что если ваш код вызывает обработчик *отмены регистрации* , который `onVisibilityModeChanged` возвращает значение, необходимо убедиться, что `onVisibilityModeChanged` оно завершено до вызова обработчика отмены регистрации.</span><span class="sxs-lookup"><span data-stu-id="eeff1-155">The `onVisibilityModeChanged` method is asynchronous which means that if your code calls the *deregister* handler that `onVisibilityModeChanged` returns, you should ensure that `onVisibilityModeChanged` has completed before calling the deregister handler.</span></span> <span data-ttu-id="eeff1-156">Один из способов сделать это — использовать `await` ключевое слово для вызова метода, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="eeff1-156">One way to do that is to use the `await` keyword on the method call as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="eeff1-157">Если вы хотите использовать только пред ES2015 JavaScript, ваш код может использовать `then` метод, чтобы дождаться, пока возвращенный объект Promise не будет разрешен и присвоить возвращаемую функцию глобальной переменной, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="eeff1-157">If you want to use only pre-ES2015 JavaScript, your code can use the `then` method to wait until the returned Promise object has resolved and assign the returned function to a global variable as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="eeff1-158">Функция отмены регистрации сама по себе асинхронна.</span><span class="sxs-lookup"><span data-stu-id="eeff1-158">The deregister function is itself asynchronous.</span></span> <span data-ttu-id="eeff1-159">Таким образом, если код не должен выполняться до завершения отмены регистрации, функция дерегистрации должна также ожидаться с помощью `await` ключевого слова или `then` метода, как показано в следующих примерах.</span><span class="sxs-lookup"><span data-stu-id="eeff1-159">So, if you have code that should not run until after the deregistration is complete, then the deregister function should also be awaited with either the `await` keyword or with a `then` method as in the following examples.</span></span>

<span data-ttu-id="eeff1-160">Чтобы отменить регистрацию обработчика:</span><span class="sxs-lookup"><span data-stu-id="eeff1-160">To deregister the handler:</span></span>

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```

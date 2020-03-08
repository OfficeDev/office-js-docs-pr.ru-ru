---
title: Передача данных и сообщений в диалоговое окно с главной страницы
description: Узнайте, как передавать данные в диалоговое окно с главной страницы с помощью API Мессажечилд и Диалогпарентмессажерецеивед.
ms.date: 03/11/2020
localization_priority: Normal
ms.openlocfilehash: 03d89a2e5ffb9060edb25dd8e0c3c71c0dd274eb
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561875"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a><span data-ttu-id="79ced-103">Передача данных и сообщений в диалоговое окно с главной страницы (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="79ced-103">Passing data and messages to a dialog box from its host page (preview)</span></span>

<span data-ttu-id="79ced-104">Надстройка может отправлять сообщения с [главной страницы](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) в диалоговое окно с помощью метода [мессажечилд](/javascript/api/office/office.dialog#messagechild-message-) объекта [DIALOG](/javascript/api/office/office.dialog) .</span><span class="sxs-lookup"><span data-stu-id="79ced-104">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using the [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method of the [Dialog](/javascript/api/office/office.dialog) object.</span></span>

> [!Important]
>
> - <span data-ttu-id="79ced-105">API, описанные в этой статье, доступны в предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="79ced-105">The APIs described in this article are in preview.</span></span> <span data-ttu-id="79ced-106">Они доступны разработчикам для экспериментов; но его не следует использовать в рабочей надстройке.</span><span class="sxs-lookup"><span data-stu-id="79ced-106">They are available to developers for experimentation; but should not be used in a production add-in.</span></span> <span data-ttu-id="79ced-107">Пока этот API не будет выпущен, используйте методы, описанные в статье [Передача сведений в диалоговое окно](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) для рабочих надстроек.</span><span class="sxs-lookup"><span data-stu-id="79ced-107">Until this API is released, use the techniques described in [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) for production add-ins.</span></span>
> - <span data-ttu-id="79ced-108">Для интерфейсов API, описанных в этой статье, требуется Office 365 (версия подписки Office).</span><span class="sxs-lookup"><span data-stu-id="79ced-108">The APIs described in this article require Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="79ced-109">Следует использовать последнюю версию для текущего месяца и сборку из канала для участников программы предварительной оценки.</span><span class="sxs-lookup"><span data-stu-id="79ced-109">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="79ced-110">Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="79ced-110">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="79ced-111">Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="79ced-111">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="79ced-112">Обратите внимание на то, что при построении градуатес к производственному каналу поддержка предварительных функций для этой сборки отключена.</span><span class="sxs-lookup"><span data-stu-id="79ced-112">Please note that when a build graduates to the production semi-annual channel, support for preview features is turned off for that build.</span></span>
> - <span data-ttu-id="79ced-113">На начальном этапе предварительной версии API поддерживаются в Excel, PowerPoint и Word; но не в Outlook.</span><span class="sxs-lookup"><span data-stu-id="79ced-113">In the initial stage of the preview, the APIs are supported in Excel, PowerPoint, and Word; but not in Outlook.</span></span>
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="79ced-114">Использование `messageChild()` с главной страницы</span><span class="sxs-lookup"><span data-stu-id="79ced-114">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="79ced-115">Когда вы вызываете API диалоговых окон Office для открытия диалогового окна, возвращается объект [DIALOG](/javascript/api/office/office.dialog) .</span><span class="sxs-lookup"><span data-stu-id="79ced-115">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="79ced-116">Она должна быть назначена переменной, которая, как правило, имеет больший объем, чем метод [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) , так как на объект будут ссылаться другие методы.</span><span class="sxs-lookup"><span data-stu-id="79ced-116">It should be assigned to a variable, which typically has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="79ced-117">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="79ced-117">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="79ced-118">Этот `Dialog` объект содержит метод [мессажечилд](/javascript/api/office/office.dialog#messagechild-message-) , который отправляет любую строку или данные преобразованного в диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="79ced-118">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, or stringified data, to the dialog box.</span></span> <span data-ttu-id="79ced-119">Это вызывает `DialogParentMessageReceived` событие в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="79ced-119">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="79ced-120">Код должен обрабатывать это событие, как показано в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="79ced-120">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="79ced-121">Рассмотрим сценарий, в котором пользовательский интерфейс диалогового окна должен сопоставляться с текущим активным листом и положением листа относительно других листов.</span><span class="sxs-lookup"><span data-stu-id="79ced-121">Consider a scenario in which the UI of the dialog should correlate with the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="79ced-122">В следующем примере в диалоговое окно `sheetPropertiesChanged` отправляются свойства листа Excel.</span><span class="sxs-lookup"><span data-stu-id="79ced-122">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="79ced-123">В этом случае текущий лист называется "Мой лист" и является 2-м листом книги.</span><span class="sxs-lookup"><span data-stu-id="79ced-123">In this case the current worksheet is named "My Sheet" and it is the 2nd sheet in the workbook.</span></span> <span data-ttu-id="79ced-124">Данные инкапсулируются в объекте, который является преобразованного, чтобы его можно было передать `messageChild`.</span><span class="sxs-lookup"><span data-stu-id="79ced-124">The data is encapsulated in an object which is stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="79ced-125">Обработка Диалогпарентмессажерецеивед в диалоговом окне</span><span class="sxs-lookup"><span data-stu-id="79ced-125">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="79ced-126">В JavaScript диалогового окна Зарегистрируйте обработчик для `DialogParentMessageReceived` события с помощью метода [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="79ced-126">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="79ced-127">Как правило, это выполняется в [методах Office. onread или Office. Initialize](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="79ced-127">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md).</span></span> <span data-ttu-id="79ced-128">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="79ced-128">The following is an example:</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="79ced-129">Затем определите `onMessageFromParent` обработчик.</span><span class="sxs-lookup"><span data-stu-id="79ced-129">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="79ced-130">Приведенный ниже код продолжает пример из предыдущего раздела.</span><span class="sxs-lookup"><span data-stu-id="79ced-130">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="79ced-131">Обратите внимание, что Office передает аргумент обработчику и что `message` свойство объекта Argument содержит строку со страницы узла.</span><span class="sxs-lookup"><span data-stu-id="79ced-131">Note that Office passes an argument to the handler and that the `message` property of argument object contains the string from the host page.</span></span> <span data-ttu-id="79ced-132">В этом примере сообщение переводится в объект, а jQuery используется для установки верхнего заголовка диалогового окна в соответствующее имя нового листа.</span><span class="sxs-lookup"><span data-stu-id="79ced-132">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="79ced-133">Рекомендуется проверить правильность регистрации обработчика.</span><span class="sxs-lookup"><span data-stu-id="79ced-133">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="79ced-134">Для этого можно передать обратный вызов `addHandlerAsync` методу, который выполняется при завершении попытки регистрации обработчика.</span><span class="sxs-lookup"><span data-stu-id="79ced-134">You can do this by passing a callback to the `addHandlerAsync` method that runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="79ced-135">Используйте обработчик для записи или отображения ошибки, если обработчик не был успешно зарегистрирован.</span><span class="sxs-lookup"><span data-stu-id="79ced-135">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="79ced-136">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="79ced-136">The following is an example.</span></span> <span data-ttu-id="79ced-137">Обратите `reportError` внимание, что это функция, не определенная здесь, записывает или отображает сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="79ced-137">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## <a name="conditional-messaging"></a><span data-ttu-id="79ced-138">Условные сообщения</span><span class="sxs-lookup"><span data-stu-id="79ced-138">Conditional messaging</span></span>

<span data-ttu-id="79ced-139">Так как вы можете выполнять `messageChild` несколько вызовов со страницы узла, но у вас есть только один обработчик в диалоговом окне для `DialogParentMessageReceived` события, обработчик должен использовать условную логику для различения разных сообщений.</span><span class="sxs-lookup"><span data-stu-id="79ced-139">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="79ced-140">Это можно сделать точно так же, как при структурировании условной передачи сообщений, когда диалоговое окно отправляет сообщение на страницу узла, как описано в [условной системе обмена сообщениями](dialog-api-in-office-add-ins.md#conditional-messaging).</span><span class="sxs-lookup"><span data-stu-id="79ced-140">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](dialog-api-in-office-add-ins.md#conditional-messaging).</span></span>

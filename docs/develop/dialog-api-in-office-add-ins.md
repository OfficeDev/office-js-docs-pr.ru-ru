---
title: Использование Office Dialog API в вашей надстройках Office
description: Общие сведения о создании диалогового окна в надстройке Office.
ms.date: 06/10/2020
localization_priority: Normal
ms.openlocfilehash: 749fd6041c2ef60a4d766e865e25d53e97298d01
ms.sourcegitcommit: 449a728118db88dea22a44f83728d21604d6ee8c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2020
ms.locfileid: "44719072"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="b298a-103">Использование Office Dialog API в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="b298a-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="b298a-104">Вы можете использовать [Office dialog API](/javascript/api/office/office.ui), чтобы открывать диалоговые окна в надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="b298a-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="b298a-105">Эта статья содержит инструкции по использованию dialog API в надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="b298a-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b298a-p102">Сведения о поддержке Dialog API см. в статье [Наборы обязательных элементов Dialog API](../reference/requirement-sets/dialog-api-requirement-sets.md). В настоящее время Dialog API поддерживается для Word, Excel, PowerPoint и Outlook.</span><span class="sxs-lookup"><span data-stu-id="b298a-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](../reference/requirement-sets/dialog-api-requirement-sets.md). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="b298a-108">Основной сценарий для Dialog API - включить аутентификацию с помощью таких ресурсов, как Google, Facebook или Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b298a-108">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="b298a-109">Дополнительные сведения см. в статье [Проверка подлинности с помощью Office Dialog API](auth-with-office-dialog-api.md) *после* ознакомления с текущей статьей.</span><span class="sxs-lookup"><span data-stu-id="b298a-109">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="b298a-110">Возможность открытия диалогового окна с помощью области задач, контентной надстройки или [команды надстройки](../design/add-in-commands.md) может позволить следующее:</span><span class="sxs-lookup"><span data-stu-id="b298a-110">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="b298a-111">отобразить страницу входа, которую невозможно открыть непосредственно в области задач;</span><span class="sxs-lookup"><span data-stu-id="b298a-111">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="b298a-112">предоставить больше места на экране (или даже весь экран) для некоторых задач в надстройке;</span><span class="sxs-lookup"><span data-stu-id="b298a-112">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="b298a-113">разместить видео, которое будет слишком маленьким в области задач.</span><span class="sxs-lookup"><span data-stu-id="b298a-113">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="b298a-114">Поскольку перекрывающиеся элементы пользовательского интерфейса не приветствуются, избегайте открытия диалогового окна на панели задач, если это не требуется в сценарий.</span><span class="sxs-lookup"><span data-stu-id="b298a-114">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="b298a-115">При планировании контактной зоны помните, что в области задач можно использовать вкладки.</span><span class="sxs-lookup"><span data-stu-id="b298a-115">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="b298a-116">Например, как в [надстройке JavaScript SalesTracker для Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span><span class="sxs-lookup"><span data-stu-id="b298a-116">For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="b298a-117">На приведенном ниже изображении показан пример диалогового окна. </span><span class="sxs-lookup"><span data-stu-id="b298a-117">The following image shows an example of a dialog box.</span></span>

![Команды надстроек](../images/auth-o-dialog-open.png)

<span data-ttu-id="b298a-119">Обратите внимание, что диалоговое окно всегда открывается в центре экрана.</span><span class="sxs-lookup"><span data-stu-id="b298a-119">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="b298a-120">Пользователь может перемещать ее и изменять ее размер.</span><span class="sxs-lookup"><span data-stu-id="b298a-120">The user can move and resize it.</span></span> <span data-ttu-id="b298a-121">Окно является *не модальным*: пользователь может продолжать взаимодействовать как с документом в главном приложении Office, так и со страницей на панели задач, если она есть.</span><span class="sxs-lookup"><span data-stu-id="b298a-121">The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="b298a-122">Откройте диалоговое окно с главной страницы</span><span class="sxs-lookup"><span data-stu-id="b298a-122">Open a dialog box from a host page</span></span>

<span data-ttu-id="b298a-123">Office JavaScript API включает в себя [Диалоговый](/javascript/api/office/office.dialog) объекта и две функции в [Office.context.ui namespace](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="b298a-123">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="b298a-124">Чтобы открыть диалоговое окно, ваш код, обычно страница в панели задач, вызывает метод [displayDialogAsync](/javascript/api/office/office.ui) и передает ему URL-адрес ресурса, который вам нужно открыть.</span><span class="sxs-lookup"><span data-stu-id="b298a-124">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="b298a-125">Страница, на которой этот метод вызван, называется "главной страницей".</span><span class="sxs-lookup"><span data-stu-id="b298a-125">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="b298a-126">Например, если вы вызываете этот метод в скрипте для index.html на панели задач, то index.html - это главная страница диалогового окна, которое открывает метод.</span><span class="sxs-lookup"><span data-stu-id="b298a-126">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="b298a-127">Ресурс, который открывается в диалоговом окне, обычно представляет собой страницу, но это может быть метод контроллера в приложении MVC, маршрут, метод веб-службы или любой другой ресурс.</span><span class="sxs-lookup"><span data-stu-id="b298a-127">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="b298a-128">В этой статье "страница" или "веб-сайт" ссылается на ресурс в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="b298a-128">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="b298a-129">Ниже приведен простой пример кода.</span><span class="sxs-lookup"><span data-stu-id="b298a-129">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="b298a-130">В случае URL-адреса используется протокол HTTP**S**,</span><span class="sxs-lookup"><span data-stu-id="b298a-130">The URL uses the HTTP**S** protocol.</span></span> <span data-ttu-id="b298a-131">Обязательный для всех страниц, загружаемых в диалоговом окне, а не только для первой страницы.</span><span class="sxs-lookup"><span data-stu-id="b298a-131">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="b298a-132">Домен диалогового окна совпадает с доменом главной страницы, которая может быть страницей в панели задач или [файлом функции](../reference/manifest/functionfile.md) команды надстройки.</span><span class="sxs-lookup"><span data-stu-id="b298a-132">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](../reference/manifest/functionfile.md) of an add-in command.</span></span> <span data-ttu-id="b298a-133">Страница, метод контроллера или другой ресурс, передаваемый в метод `displayDialogAsync`, должен быть в том же домене, что и страница ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="b298a-133">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b298a-134">Главная страница и ресурс, который открывается в диалоговом окне, должны иметь один и тот же полный домен.</span><span class="sxs-lookup"><span data-stu-id="b298a-134">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="b298a-135">Если вы попробуете передать поддомен домена надстройки в `displayDialogAsync`, ничего не получится.</span><span class="sxs-lookup"><span data-stu-id="b298a-135">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="b298a-136">Полные доменные имена, включая поддомены, должны совпадать.</span><span class="sxs-lookup"><span data-stu-id="b298a-136">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="b298a-137">После загрузки первой страницы (или другого ресурса) пользователь может использовать ссылки или другой пользовательский интерфейс для перехода на любой веб-сайт (или другой ресурс), использующий HTTPS.</span><span class="sxs-lookup"><span data-stu-id="b298a-137">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="b298a-138">Первая страница также может сразу перенаправлять пользователя на другой сайт.</span><span class="sxs-lookup"><span data-stu-id="b298a-138">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="b298a-139">По умолчанию диалоговое окно занимает 80 % высоты и ширины экрана устройства, но вы можете установить другие соотношения путем передачи объекта конфигурации в метод, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="b298a-139">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="b298a-140">Подобная надстройка приведена в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="b298a-140">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="b298a-p112">Установите оба значения равными 100 %, чтобы надстройка открывалась во весь экран. (На самом деле, максимальное значение составляет 99,5 %, возможность перемещать окно и изменять его размер сохраняется.)</span><span class="sxs-lookup"><span data-stu-id="b298a-p112">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="b298a-p113">Из главного окна можно открыть только одно диалоговое окно. При попытке открыть еще одно диалоговое окно произойдет ошибка. Поэтому если пользователь, например, откроет диалоговое окно из области задач, он не сможет открыть второе диалоговое окно на другой странице в области задач. Но при открытии диалогового окна с помощью [команды надстройки](../design/add-in-commands.md) каждый раз открывается новый (невидимый) HTML-файл. При этом создается новое (невидимое) главное окно, которое может запускать собственное диалоговое окно. Дополнительные сведения см. в разделе [Ошибки метода displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="b298a-p113">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="b298a-149">Использование параметра производительности в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b298a-149">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="b298a-150">`displayInIframe` — дополнительное свойство в объекте конфигурации, которое можно передать `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="b298a-150">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="b298a-151">Когда этому свойству присвоено значение `true`, а надстройка запущена для документа в Office в Интернете, диалоговое окно будет открываться быстрее, потому что будет выступать как плавающий фрейм iframe.</span><span class="sxs-lookup"><span data-stu-id="b298a-151">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="b298a-152">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b298a-152">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="b298a-153">Значение по умолчанию: `false`. Его использование равнозначно пропуску всего свойства.</span><span class="sxs-lookup"><span data-stu-id="b298a-153">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="b298a-154">Если надстройка не работает в Office в Интернете, `displayInIframe` игнорируется.</span><span class="sxs-lookup"><span data-stu-id="b298a-154">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="b298a-155">Вам **не** следует `displayInIframe: true`использовать, если диалоговое окно будет выполнять перенаправление на страницу, которую невозможно открыть в элементе iframe.</span><span class="sxs-lookup"><span data-stu-id="b298a-155">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="b298a-156">Например, страницы входа многих популярных веб-служб, таких как Google или учетной записи Майкрософт, невозможно открыть в iframe.</span><span class="sxs-lookup"><span data-stu-id="b298a-156">For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="b298a-157">Отправка сведений из диалогового окна главной странице</span><span class="sxs-lookup"><span data-stu-id="b298a-157">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="b298a-158">Диалоговое окно может взаимодействовать с главной страницей в области задач, если:</span><span class="sxs-lookup"><span data-stu-id="b298a-158">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="b298a-159">Текущая страница в диалоговом окне не находится в том же домене, что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="b298a-159">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="b298a-p117">На странице загружается библиотека API JavaScript для Office. (Как и любая страница, использующая библиотеку API JavaScript для Office, сценарий для страницы должен назначить метод `Office.initialize` свойству, хотя это может быть пустой метод. Дополнительные сведения см. [в статье Initialize Your надстройка Office](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="b298a-p117">The Office JavaScript API library is loaded in the page. (Like any page that uses the Office JavaScript API library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initialize your Office Add-in](initialize-add-in.md).)</span></span>

<span data-ttu-id="b298a-163">Код в диалоговом окне использует функцию [messageParent](/javascript/api/office/office.ui#messageparent-message-), чтобы отправить на главную страницу логическое значение или строковое сообщение.</span><span class="sxs-lookup"><span data-stu-id="b298a-163">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="b298a-164">Строка может быть словом, предложением, большим двоичным объектом XML, преобразованными данными JSON или любыми другими объектами, которые можно сериализовать в строку.</span><span class="sxs-lookup"><span data-stu-id="b298a-164">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="b298a-165">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b298a-165">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - <span data-ttu-id="b298a-166">Функцию `messageParent` можно вызывать только на странице, которая относится к тому же домену (включая протокол и порт), что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="b298a-166">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>
> - <span data-ttu-id="b298a-167">`messageParent`Функция является одним из двух *only* API Office JS, которые можно вызывать в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="b298a-167">The `messageParent` function is one of *only* two Office JS APIs that can be called in the dialog box.</span></span> 
> - <span data-ttu-id="b298a-168">Другой API JS, который может вызываться в диалоговом окне, — это `Office.context.requirements.isSetSupported` .</span><span class="sxs-lookup"><span data-stu-id="b298a-168">The other JS API that can be called in the dialog box is `Office.context.requirements.isSetSupported`.</span></span> <span data-ttu-id="b298a-169">Сведения о нем: [Указание ведущих приложений Office и требований к API](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="b298a-169">For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span> <span data-ttu-id="b298a-170">Однако в диалоговом окне этот API не поддерживается в Outlook 2016 1-Time Purchase (версия MSI).</span><span class="sxs-lookup"><span data-stu-id="b298a-170">However, in the dialog box, this API isn't supported in Outlook 2016 one-time purchase (that is, the MSI version).</span></span>

<span data-ttu-id="b298a-171">В следующем примере `googleProfile` — это строковое представление профиля Google пользователя.</span><span class="sxs-lookup"><span data-stu-id="b298a-171">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="b298a-p120">Чтобы главная страница получила сообщение, ее необходимо настроить. Для этого добавьте параметр обратного вызова в исходный вызов `displayDialogAsync`. Обратный вызов назначает событию `DialogMessageReceived` обработчик. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b298a-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
> - <span data-ttu-id="b298a-176">Office передает объект [AsyncResult](/javascript/api/office/office.asyncresult) в функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b298a-176">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="b298a-177">Он представляет результат попытки открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="b298a-177">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="b298a-178">Он не представляет результат событий в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="b298a-178">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="b298a-179">Подробнее об этом различии см. в [Обработка ошибок и событий](dialog-handle-errors-events.md). </span><span class="sxs-lookup"><span data-stu-id="b298a-179">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="b298a-180">Для свойства `value` объекта `asyncResult` задан объект [Dialog](/javascript/api/office/office.dialog), который существует на главной странице, а не в контексте выполнения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="b298a-180">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="b298a-p122">`processMessage` — это функция, которая обрабатывает событие. Вы можете присвоить ей любое имя.</span><span class="sxs-lookup"><span data-stu-id="b298a-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="b298a-183">Переменная `dialog` объявляется в более широком контексте, чем обратный вызов, так как на нее также ссылается `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="b298a-183">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="b298a-184">Ниже приведен простой пример обработчика для события `DialogMessageReceived`.</span><span class="sxs-lookup"><span data-stu-id="b298a-184">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="b298a-185">Office передает объект `arg` в обработчик.</span><span class="sxs-lookup"><span data-stu-id="b298a-185">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="b298a-186">Его `message` свойством является логическое значение или строка, отправляемая при вызове `messageParent` в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="b298a-186">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="b298a-187">В данном примере это профиль пользователя из учетной записи Майкрософт, Google или другой службы, представленный в виде строки, поэтому он десериализируется обратно в объект с помощью метода `JSON.parse`.</span><span class="sxs-lookup"><span data-stu-id="b298a-187">In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="b298a-p124">Функция `showUserName` не показана. Она может отображать персонализированное приветствие в области задач.</span><span class="sxs-lookup"><span data-stu-id="b298a-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="b298a-190">Когда взаимодействие пользователя с диалоговым окном закончится, обработчик сообщений должен закрыть диалоговое окно, как показано в этом примере.</span><span class="sxs-lookup"><span data-stu-id="b298a-190">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="b298a-191">Объект `dialog` должен быть таким же, как объект, который возвращается при вызове `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="b298a-191">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="b298a-192">Вызов метода `dialog.close` дает указание Office немедленно закрыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="b298a-192">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="b298a-193">Пример надстройки, в которой используются эти методы, см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="b298a-193">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="b298a-p125">Если надстройка должна открыть другую страницу области задач после получения сообщения, можно использовать метод `window.location.replace` (или `window.location.href`) в последней строке обработчика. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b298a-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="b298a-196">Пример подобной надстройки см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="b298a-196">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="b298a-197">Условные сообщения</span><span class="sxs-lookup"><span data-stu-id="b298a-197">Conditional messaging</span></span>

<span data-ttu-id="b298a-p126">Так как из диалогового окна можно отправить несколько вызовов `messageParent`, но на главной странице есть только один обработчик для события `DialogMessageReceived`, обработчику необходимо использовать условную логику, чтобы различать сообщения. Например, если диалоговое окно предлагает пользователю войти в учетную запись Майкрософт, Google или другого поставщика удостоверений, оно отправляет профиль пользователя в виде сообщения. Если выполнить аутентификацию не удается, диалоговое окно отправляет сведения об ошибке на главную страницу, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="b298a-p126">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
> - <span data-ttu-id="b298a-201">Переменная `loginSuccess` будет инициализирована после считывания отклика HTTP от поставщика удостоверений.</span><span class="sxs-lookup"><span data-stu-id="b298a-201">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="b298a-p127">Реализация функций `getProfile` и `getError` не показана. Они получают данные из параметра запроса или ответа HTTP.</span><span class="sxs-lookup"><span data-stu-id="b298a-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="b298a-p128">В зависимости от того, удалось ли выполнить вход, отправляются анонимные объекты различных типов. Оба содержат свойство `messageType`, но один содержит свойство `profile`, а другой — свойство `error`.</span><span class="sxs-lookup"><span data-stu-id="b298a-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="b298a-p129">Код обработчика на главной странице использует значение свойства `messageType` для разветвления, как показано в приведенном ниже примере. Обратите внимание на то, что здесь используется та же функция `showUserName`, что и в примере выше, а функция `showNotification` отображает сообщение об ошибке в элементе пользовательского интерфейса на главной странице.</span><span class="sxs-lookup"><span data-stu-id="b298a-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

> [!NOTE]
> <span data-ttu-id="b298a-208">Реализация функции `showNotification` не показана в примере кода, представленном в этой статье.</span><span class="sxs-lookup"><span data-stu-id="b298a-208">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="b298a-209">Пример возможного способа реализации этой функции в своей надстройке см. в статье [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="b298a-209">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="b298a-210">Передача данных диалоговому окну</span><span class="sxs-lookup"><span data-stu-id="b298a-210">Pass information to the dialog box</span></span>

<span data-ttu-id="b298a-p131">Иногда главной странице нужно передать данные в диалоговое окно. Есть два основных способа обеспечить эту возможность:</span><span class="sxs-lookup"><span data-stu-id="b298a-p131">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="b298a-213">Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="b298a-213">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="b298a-214">Храните информацию в месте, доступном как для главного, так и для диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="b298a-214">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="b298a-215">Два окна не разделяют общее хранилище сеансов, но *если они имеют один и тот же домен* (включая номер порта, если таковой имеется), они совместно используют общее [Локальное хранилище](https://www.w3schools.com/html/html5_webstorage.asp).\*</span><span class="sxs-lookup"><span data-stu-id="b298a-215">The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>

> [!NOTE]
> <span data-ttu-id="b298a-216">\* Существует ошибка, влияющая на вашу стратегию обработки маркеров.</span><span class="sxs-lookup"><span data-stu-id="b298a-216">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="b298a-217">Если надстройка работает в **Office в Интернете** с использованием браузера Safari или Microsoft Edge, у диалогового окна и области задач нет одного общего локального хранилища, поэтому его нельзя использовать для связи между ними.</span><span class="sxs-lookup"><span data-stu-id="b298a-217">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="b298a-218">Использование локального хранилища</span><span class="sxs-lookup"><span data-stu-id="b298a-218">Use local storage</span></span>

<span data-ttu-id="b298a-219">Чтобы использовать локальное хранилище, код вызывает метод `setItem` объекта `window.localStorage` на главной странице перед вызовом `displayDialogAsync`, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="b298a-219">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="b298a-220">Код в диалоговом окне считывает элемент, когда он необходим, как в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="b298a-220">Code in the dialog box reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a><span data-ttu-id="b298a-221">Использование параметров запроса</span><span class="sxs-lookup"><span data-stu-id="b298a-221">Use query parameters</span></span>

<span data-ttu-id="b298a-222">В приведенном ниже примере показано, как передавать данные с помощью параметра запроса.</span><span class="sxs-lookup"><span data-stu-id="b298a-222">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="b298a-223">Пример, в котором используется эта техника, см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="b298a-223">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="b298a-224">Код в вашем диалоговом окне может проанализировать URL-адрес и прочитать значение параметра.</span><span class="sxs-lookup"><span data-stu-id="b298a-224">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b298a-p134">Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`. (Этот параметр добавляется после пользовательских параметров запроса, если они есть. Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать. То же значение добавляется в хранилище сеанса диалогового окна. *Ваш код не должен ни считывать это значение, ни записывать в него данные*.</span><span class="sxs-lookup"><span data-stu-id="b298a-p134">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

> [!NOTE]
> <span data-ttu-id="b298a-230">Теперь вы `messageChild` можете просмотреть API, который родительская страница может использовать для отправки сообщений в диалоговое окно, так как `messageParent` описанный выше API отправляет сообщения из диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="b298a-230">There is now in preview a `messageChild` API that the parent page can use to send messages to the dialog just as the `messageParent` API described above sends messages from the dialog.</span></span> <span data-ttu-id="b298a-231">Дополнительные сведения см. в статье [Передача данных и сообщений в диалоговое окно с главной страницы](parent-to-dialog.md).</span><span class="sxs-lookup"><span data-stu-id="b298a-231">For more about it, see [Passing data and messages to a dialog box from its host page](parent-to-dialog.md).</span></span> <span data-ttu-id="b298a-232">Мы рекомендуем испытать ее, но для рабочих надстроек мы рекомендуем использовать методы, описанные в этом разделе.</span><span class="sxs-lookup"><span data-stu-id="b298a-232">We encourage you to try it out, but for production add-ins, we recommend that you use the techniques described in this section.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="b298a-233">Закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="b298a-233">Closing the dialog box</span></span>

<span data-ttu-id="b298a-p136">Вы можете добавить в диалоговое окно кнопку, которая будет его закрывать. Для этого обработчик событий кнопки должен использовать `messageParent`, чтобы сообщить главной странице, что кнопка нажата. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b298a-p136">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="b298a-237">Обработчик главной страницы для `DialogMessageReceived` вызовет `dialog.close`, как показано в этом примере.</span><span class="sxs-lookup"><span data-stu-id="b298a-237">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="b298a-238">(См. предыдущие примеры, в которых показано, как `dialog` инициализируется объект).</span><span class="sxs-lookup"><span data-stu-id="b298a-238">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="b298a-239">Даже если у вас нет собственного пользовательского интерфейса для закрытия диалогового окна, пользователь может закрыть диалоговое окно, выбрав **X** в правом верхнем углу.</span><span class="sxs-lookup"><span data-stu-id="b298a-239">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="b298a-240">Это действие запускает событие `DialogEventReceived`.</span><span class="sxs-lookup"><span data-stu-id="b298a-240">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="b298a-241">Чтобы главная область могла реагировать на это событие, для нее должен быть объявлен обработчик этого события.</span><span class="sxs-lookup"><span data-stu-id="b298a-241">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="b298a-242">Дополнительные сведения см. в разделе [Ошибок и события в диалоговом окне](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box).</span><span class="sxs-lookup"><span data-stu-id="b298a-242">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="b298a-243">Продвинутые темы и специальные сценарии</span><span class="sxs-lookup"><span data-stu-id="b298a-243">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="b298a-244">Используйте Dialog API, чтобы показать видео</span><span class="sxs-lookup"><span data-stu-id="b298a-244">Use the Dialog API to show a video</span></span>

<span data-ttu-id="b298a-245">См. статью [Использование диалогового окна «Office» для отображения видео](dialog-video.md).</span><span class="sxs-lookup"><span data-stu-id="b298a-245">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="b298a-246">Использование Dialog API в потоке аутентификации</span><span class="sxs-lookup"><span data-stu-id="b298a-246">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="b298a-247">См. статью[ Проверка подлинности с помощью Office Dialog API ](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="b298a-247">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="b298a-248">Использование Office dialog API с одностраничными приложениями и клиентской маршрутизацией</span><span class="sxs-lookup"><span data-stu-id="b298a-248">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="b298a-249">При использовании Office dialog API, SPA и маршрутизация на стороне клиента должны обрабатываться с осторожностью</span><span class="sxs-lookup"><span data-stu-id="b298a-249">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="b298a-250">См. статью[Рекомендации по использованию Office dialog API в SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span><span class="sxs-lookup"><span data-stu-id="b298a-250">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="b298a-251">Обработка ошибок и событий</span><span class="sxs-lookup"><span data-stu-id="b298a-251">Error and event handling</span></span>

<span data-ttu-id="b298a-252">См. статью об ошибках и событиях [Обработка ошибок и событий в Office dialog box](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="b298a-252">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="b298a-253">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="b298a-253">Next steps</span></span>

<span data-ttu-id="b298a-254">Узнайте о том, как использовать Office dialog API, в [Рекомендации по использованию Office dialog API](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="b298a-254">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

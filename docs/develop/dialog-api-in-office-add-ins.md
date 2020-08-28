---
title: Использование Office Dialog API в вашей надстройках Office
description: Изучите основы создания диалогового окна в надстройке Office
ms.date: 08/20/2020
localization_priority: Normal
ms.openlocfilehash: 9d333c12d629232ece39bc30948318fbcafa3aa0
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292793"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="abf05-103">Использование Office Dialog API в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="abf05-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="abf05-104">Вы можете использовать [Office dialog API](/javascript/api/office/office.ui), чтобы открывать диалоговые окна в надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="abf05-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="abf05-105">Эта статья содержит инструкции по использованию dialog API в надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="abf05-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="abf05-p102">Сведения о поддержке Dialog API см. в статье [Наборы обязательных элементов Dialog API](../reference/requirement-sets/dialog-api-requirement-sets.md). В настоящее время Dialog API поддерживается для Word, Excel, PowerPoint и Outlook.</span><span class="sxs-lookup"><span data-stu-id="abf05-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](../reference/requirement-sets/dialog-api-requirement-sets.md). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="abf05-108">Основной сценарий для Dialog API - включить аутентификацию с помощью таких ресурсов, как Google, Facebook или Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="abf05-108">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="abf05-109">Дополнительные сведения см. в статье [Проверка подлинности с помощью Office Dialog API](auth-with-office-dialog-api.md) *после* ознакомления с текущей статьей.</span><span class="sxs-lookup"><span data-stu-id="abf05-109">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="abf05-110">Возможность открытия диалогового окна с помощью области задач, контентной надстройки или [команды надстройки](../design/add-in-commands.md) может позволить следующее:</span><span class="sxs-lookup"><span data-stu-id="abf05-110">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="abf05-111">отобразить страницу входа, которую невозможно открыть непосредственно в области задач;</span><span class="sxs-lookup"><span data-stu-id="abf05-111">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="abf05-112">предоставить больше места на экране (или даже весь экран) для некоторых задач в надстройке;</span><span class="sxs-lookup"><span data-stu-id="abf05-112">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="abf05-113">разместить видео, которое будет слишком маленьким в области задач.</span><span class="sxs-lookup"><span data-stu-id="abf05-113">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="abf05-114">Поскольку перекрывающиеся элементы пользовательского интерфейса не приветствуются, избегайте открытия диалогового окна на панели задач, если это не требуется в сценарий.</span><span class="sxs-lookup"><span data-stu-id="abf05-114">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="abf05-115">При планировании контактной зоны помните, что в области задач можно использовать вкладки.</span><span class="sxs-lookup"><span data-stu-id="abf05-115">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="abf05-116">Например, как в [надстройке JavaScript SalesTracker для Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span><span class="sxs-lookup"><span data-stu-id="abf05-116">For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="abf05-117">На приведенном ниже изображении показан пример диалогового окна. </span><span class="sxs-lookup"><span data-stu-id="abf05-117">The following image shows an example of a dialog box.</span></span>

![Команды надстроек](../images/auth-o-dialog-open.png)

<span data-ttu-id="abf05-119">Обратите внимание, что диалоговое окно всегда открывается в центре экрана.</span><span class="sxs-lookup"><span data-stu-id="abf05-119">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="abf05-120">Пользователь может перемещать ее и изменять ее размер.</span><span class="sxs-lookup"><span data-stu-id="abf05-120">The user can move and resize it.</span></span> <span data-ttu-id="abf05-121">Окно не является *модальным*— пользователь может продолжать взаимодействовать с документом в приложении Office и со страницей в области задач, если таковая имеется.</span><span class="sxs-lookup"><span data-stu-id="abf05-121">The window is *nonmodal*--a user can continue to interact with both the document in the Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="abf05-122">Откройте диалоговое окно с главной страницы</span><span class="sxs-lookup"><span data-stu-id="abf05-122">Open a dialog box from a host page</span></span>

<span data-ttu-id="abf05-123">Office JavaScript API включает в себя [Диалоговый](/javascript/api/office/office.dialog) объекта и две функции в [Office.context.ui namespace](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="abf05-123">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="abf05-124">Чтобы открыть диалоговое окно, ваш код, обычно страница в панели задач, вызывает метод [displayDialogAsync](/javascript/api/office/office.ui) и передает ему URL-адрес ресурса, который вам нужно открыть.</span><span class="sxs-lookup"><span data-stu-id="abf05-124">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="abf05-125">Страница, на которой этот метод вызван, называется "главной страницей".</span><span class="sxs-lookup"><span data-stu-id="abf05-125">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="abf05-126">Например, если вы вызываете этот метод в скрипте для index.html на панели задач, то index.html - это главная страница диалогового окна, которое открывает метод.</span><span class="sxs-lookup"><span data-stu-id="abf05-126">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="abf05-127">Ресурс, который открывается в диалоговом окне, обычно представляет собой страницу, но это может быть метод контроллера в приложении MVC, маршрут, метод веб-службы или любой другой ресурс.</span><span class="sxs-lookup"><span data-stu-id="abf05-127">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="abf05-128">В этой статье "страница" или "веб-сайт" ссылается на ресурс в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="abf05-128">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="abf05-129">Ниже приведен простой пример кода.</span><span class="sxs-lookup"><span data-stu-id="abf05-129">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="abf05-130">В случае URL-адреса используется протокол HTTP**S**,</span><span class="sxs-lookup"><span data-stu-id="abf05-130">The URL uses the HTTP**S** protocol.</span></span> <span data-ttu-id="abf05-131">Обязательный для всех страниц, загружаемых в диалоговом окне, а не только для первой страницы.</span><span class="sxs-lookup"><span data-stu-id="abf05-131">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="abf05-132">Домен диалогового окна совпадает с доменом главной страницы, которая может быть страницей в панели задач или [файлом функции](../reference/manifest/functionfile.md) команды надстройки.</span><span class="sxs-lookup"><span data-stu-id="abf05-132">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](../reference/manifest/functionfile.md) of an add-in command.</span></span> <span data-ttu-id="abf05-133">Страница, метод контроллера или другой ресурс, передаваемый в метод `displayDialogAsync`, должен быть в том же домене, что и страница ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="abf05-133">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="abf05-134">Главная страница и ресурс, который открывается в диалоговом окне, должны иметь один и тот же полный домен.</span><span class="sxs-lookup"><span data-stu-id="abf05-134">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="abf05-135">Если вы попробуете передать поддомен домена надстройки в `displayDialogAsync`, ничего не получится.</span><span class="sxs-lookup"><span data-stu-id="abf05-135">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="abf05-136">Полные доменные имена, включая поддомены, должны совпадать.</span><span class="sxs-lookup"><span data-stu-id="abf05-136">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="abf05-137">После загрузки первой страницы (или другого ресурса) пользователь может использовать ссылки или другой пользовательский интерфейс для перехода на любой веб-сайт (или другой ресурс), использующий HTTPS.</span><span class="sxs-lookup"><span data-stu-id="abf05-137">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="abf05-138">Первая страница также может сразу перенаправлять пользователя на другой сайт.</span><span class="sxs-lookup"><span data-stu-id="abf05-138">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="abf05-139">По умолчанию диалоговое окно занимает 80 % высоты и ширины экрана устройства, но вы можете установить другие соотношения путем передачи объекта конфигурации в метод, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="abf05-139">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="abf05-140">Подобная надстройка приведена в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="abf05-140">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="abf05-p112">Установите оба значения равными 100 %, чтобы надстройка открывалась во весь экран. (На самом деле, максимальное значение составляет 99,5 %, возможность перемещать окно и изменять его размер сохраняется.)</span><span class="sxs-lookup"><span data-stu-id="abf05-p112">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="abf05-p113">Из главного окна можно открыть только одно диалоговое окно. При попытке открыть еще одно диалоговое окно произойдет ошибка. Поэтому если пользователь, например, откроет диалоговое окно из области задач, он не сможет открыть второе диалоговое окно на другой странице в области задач. Но при открытии диалогового окна с помощью [команды надстройки](../design/add-in-commands.md) каждый раз открывается новый (невидимый) HTML-файл. При этом создается новое (невидимое) главное окно, которое может запускать собственное диалоговое окно. Дополнительные сведения см. в разделе [Ошибки метода displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="abf05-p113">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="abf05-149">Использование параметра производительности в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="abf05-149">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="abf05-150">`displayInIframe` — дополнительное свойство в объекте конфигурации, которое можно передать `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="abf05-150">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="abf05-151">Когда этому свойству присвоено значение `true`, а надстройка запущена для документа в Office в Интернете, диалоговое окно будет открываться быстрее, потому что будет выступать как плавающий фрейм iframe.</span><span class="sxs-lookup"><span data-stu-id="abf05-151">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="abf05-152">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="abf05-152">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="abf05-153">Значение по умолчанию: `false`. Его использование равнозначно пропуску всего свойства.</span><span class="sxs-lookup"><span data-stu-id="abf05-153">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="abf05-154">Если надстройка не работает в Office в Интернете, `displayInIframe` игнорируется.</span><span class="sxs-lookup"><span data-stu-id="abf05-154">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="abf05-155">Вам **не** следует `displayInIframe: true`использовать, если диалоговое окно будет выполнять перенаправление на страницу, которую невозможно открыть в элементе iframe.</span><span class="sxs-lookup"><span data-stu-id="abf05-155">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="abf05-156">Например, страницы входа многих популярных веб-служб, таких как учетные записи Google и Майкрософт, не могут быть открыты в IFRAME.</span><span class="sxs-lookup"><span data-stu-id="abf05-156">For example, the sign in pages of many popular web services, such as Google and Microsoft account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="abf05-157">Отправка сведений из диалогового окна главной странице</span><span class="sxs-lookup"><span data-stu-id="abf05-157">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="abf05-158">Диалоговое окно может взаимодействовать с главной страницей в области задач, если:</span><span class="sxs-lookup"><span data-stu-id="abf05-158">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="abf05-159">Текущая страница в диалоговом окне не находится в том же домене, что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="abf05-159">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="abf05-p117">На странице загружается библиотека API JavaScript для Office. (Как и любая страница, использующая библиотеку API JavaScript для Office, сценарий для страницы должен назначить метод `Office.initialize` свойству, хотя это может быть пустой метод. Дополнительные сведения см. [в статье Initialize Your надстройка Office](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="abf05-p117">The Office JavaScript API library is loaded in the page. (Like any page that uses the Office JavaScript API library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initialize your Office Add-in](initialize-add-in.md).)</span></span>

<span data-ttu-id="abf05-163">Код в диалоговом окне использует функцию [messageParent](/javascript/api/office/office.ui#messageparent-message-), чтобы отправить на главную страницу логическое значение или строковое сообщение.</span><span class="sxs-lookup"><span data-stu-id="abf05-163">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="abf05-164">Строка может быть словом, предложением, большим двоичным объектом XML, преобразованными данными JSON или любыми другими объектами, которые можно сериализовать в строку.</span><span class="sxs-lookup"><span data-stu-id="abf05-164">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="abf05-165">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="abf05-165">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - <span data-ttu-id="abf05-166">Функцию `messageParent` можно вызывать только на странице, которая относится к тому же домену (включая протокол и порт), что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="abf05-166">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>
> - <span data-ttu-id="abf05-167">`messageParent`Функция является одним из двух *only* API Office JS, которые можно вызывать в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="abf05-167">The `messageParent` function is one of *only* two Office JS APIs that can be called in the dialog box.</span></span> 
> - <span data-ttu-id="abf05-168">Другой API JS, который может вызываться в диалоговом окне, — это `Office.context.requirements.isSetSupported` .</span><span class="sxs-lookup"><span data-stu-id="abf05-168">The other JS API that can be called in the dialog box is `Office.context.requirements.isSetSupported`.</span></span> <span data-ttu-id="abf05-169">Сведения о том, как [указать приложения Office и требования к API](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="abf05-169">For information about it, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).</span></span> <span data-ttu-id="abf05-170">Однако в диалоговом окне этот API не поддерживается в Outlook 2016 1-Time Purchase (версия MSI).</span><span class="sxs-lookup"><span data-stu-id="abf05-170">However, in the dialog box, this API isn't supported in Outlook 2016 one-time purchase (that is, the MSI version).</span></span>


<span data-ttu-id="abf05-171">В следующем примере `googleProfile` — это строковое представление профиля Google пользователя.</span><span class="sxs-lookup"><span data-stu-id="abf05-171">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="abf05-p120">Чтобы главная страница получила сообщение, ее необходимо настроить. Для этого добавьте параметр обратного вызова в исходный вызов `displayDialogAsync`. Обратный вызов назначает событию `DialogMessageReceived` обработчик. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="abf05-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="abf05-176">Office передает объект [AsyncResult](/javascript/api/office/office.asyncresult) в функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="abf05-176">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="abf05-177">Он представляет результат попытки открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="abf05-177">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="abf05-178">Он не представляет результат событий в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="abf05-178">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="abf05-179">Подробнее об этом различии см. в [Обработка ошибок и событий](dialog-handle-errors-events.md). </span><span class="sxs-lookup"><span data-stu-id="abf05-179">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="abf05-180">Для свойства `value` объекта `asyncResult` задан объект [Dialog](/javascript/api/office/office.dialog), который существует на главной странице, а не в контексте выполнения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="abf05-180">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="abf05-p122">`processMessage` — это функция, которая обрабатывает событие. Вы можете присвоить ей любое имя.</span><span class="sxs-lookup"><span data-stu-id="abf05-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="abf05-183">Переменная `dialog` объявляется в более широком контексте, чем обратный вызов, так как на нее также ссылается `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="abf05-183">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="abf05-184">Ниже приведен простой пример обработчика для события `DialogMessageReceived`.</span><span class="sxs-lookup"><span data-stu-id="abf05-184">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="abf05-185">Office передает объект `arg` в обработчик.</span><span class="sxs-lookup"><span data-stu-id="abf05-185">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="abf05-186">Его `message` свойством является логическое значение или строка, отправляемая при вызове `messageParent` в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="abf05-186">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="abf05-187">В этом примере это преобразованногоное представление профиля пользователя из службы, например учетной записи Майкрософт или Google, поэтому она возвращается в объект с `JSON.parse` .</span><span class="sxs-lookup"><span data-stu-id="abf05-187">In this example, it is a stringified representation of a user's profile from a service such as Microsoft account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="abf05-p124">Функция `showUserName` не показана. Она может отображать персонализированное приветствие в области задач.</span><span class="sxs-lookup"><span data-stu-id="abf05-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="abf05-190">Когда взаимодействие пользователя с диалоговым окном закончится, обработчик сообщений должен закрыть диалоговое окно, как показано в этом примере.</span><span class="sxs-lookup"><span data-stu-id="abf05-190">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="abf05-191">Объект `dialog` должен быть таким же, как объект, который возвращается при вызове `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="abf05-191">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="abf05-192">Вызов метода `dialog.close` дает указание Office немедленно закрыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="abf05-192">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="abf05-193">Пример надстройки, в которой используются эти методы, см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="abf05-193">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="abf05-p125">Если надстройка должна открыть другую страницу области задач после получения сообщения, можно использовать метод `window.location.replace` (или `window.location.href`) в последней строке обработчика. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="abf05-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="abf05-196">Пример подобной надстройки см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="abf05-196">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="abf05-197">Условные сообщения</span><span class="sxs-lookup"><span data-stu-id="abf05-197">Conditional messaging</span></span>

<span data-ttu-id="abf05-198">Так как из диалогового окна можно отправить несколько вызовов `messageParent`, но на главной странице есть только один обработчик для события `DialogMessageReceived`, обработчику необходимо использовать условную логику, чтобы различать сообщения.</span><span class="sxs-lookup"><span data-stu-id="abf05-198">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="abf05-199">Например, если диалоговое окно предлагает пользователю выполнить вход в поставщика удостоверений, например в учетной записи Майкрософт или Google, он отправляет профиль пользователя в виде сообщения.</span><span class="sxs-lookup"><span data-stu-id="abf05-199">For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft account or Google, it sends the user's profile as a message.</span></span> <span data-ttu-id="abf05-200">Если выполнить аутентификацию не удается, диалоговое окно отправляет сведения об ошибке на главную страницу, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="abf05-200">If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="abf05-201">Переменная `loginSuccess` будет инициализирована после считывания отклика HTTP от поставщика удостоверений.</span><span class="sxs-lookup"><span data-stu-id="abf05-201">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="abf05-p127">Реализация функций `getProfile` и `getError` не показана. Они получают данные из параметра запроса или ответа HTTP.</span><span class="sxs-lookup"><span data-stu-id="abf05-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="abf05-p128">В зависимости от того, удалось ли выполнить вход, отправляются анонимные объекты различных типов. Оба содержат свойство `messageType`, но один содержит свойство `profile`, а другой — свойство `error`.</span><span class="sxs-lookup"><span data-stu-id="abf05-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="abf05-p129">Код обработчика на главной странице использует значение свойства `messageType` для разветвления, как показано в приведенном ниже примере. Обратите внимание на то, что здесь используется та же функция `showUserName`, что и в примере выше, а функция `showNotification` отображает сообщение об ошибке в элементе пользовательского интерфейса на главной странице.</span><span class="sxs-lookup"><span data-stu-id="abf05-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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
> <span data-ttu-id="abf05-208">Реализация функции `showNotification` не показана в примере кода, представленном в этой статье.</span><span class="sxs-lookup"><span data-stu-id="abf05-208">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="abf05-209">Пример возможного способа реализации этой функции в своей надстройке см. в статье [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="abf05-209">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="abf05-210">Передача данных диалоговому окну</span><span class="sxs-lookup"><span data-stu-id="abf05-210">Pass information to the dialog box</span></span>

<span data-ttu-id="abf05-211">Надстройка может отправлять сообщения с [главной страницы](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) в диалоговое окно с помощью [диалогового окна Dialog. мессажечилд](/javascript/api/office/office.dialog#messagechild-message-).</span><span class="sxs-lookup"><span data-stu-id="abf05-211">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using [Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-).</span></span>

> [!NOTE]
> <span data-ttu-id="abf05-212">API диалоговых окон поддерживаются только в Excel, PowerPoint и Word.</span><span class="sxs-lookup"><span data-stu-id="abf05-212">These dialog APIs are supported in only Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="abf05-213">Поддержка Outlook находится на стадии разработки.</span><span class="sxs-lookup"><span data-stu-id="abf05-213">Support for Outlook is under development.</span></span>

### <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="abf05-214">Использование `messageChild()` с главной страницы</span><span class="sxs-lookup"><span data-stu-id="abf05-214">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="abf05-215">Когда вы вызываете API диалоговых окон Office для открытия диалогового окна, возвращается объект [DIALOG](/javascript/api/office/office.dialog) .</span><span class="sxs-lookup"><span data-stu-id="abf05-215">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="abf05-216">Она должна быть назначена переменной с большей областью действия, чем метод [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) , так как на объект будут ссылаться другие методы.</span><span class="sxs-lookup"><span data-stu-id="abf05-216">It should be assigned to a variable that has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="abf05-217">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="abf05-217">The following is an example:</span></span>

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

<span data-ttu-id="abf05-218">Этот `Dialog` объект содержит метод [мессажечилд](/javascript/api/office/office.dialog#messagechild-message-) , который отправляет любую строку, в том числе данные преобразованного, в диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="abf05-218">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, including stringified data, to the dialog box.</span></span> <span data-ttu-id="abf05-219">Это вызывает `DialogParentMessageReceived` событие в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="abf05-219">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="abf05-220">Код должен обрабатывать это событие, как показано в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="abf05-220">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="abf05-221">Рассмотрим сценарий, в котором пользовательский интерфейс диалогового окна связан с текущим активным листом и положением листа относительно других листов.</span><span class="sxs-lookup"><span data-stu-id="abf05-221">Consider a scenario in which the UI of the dialog is related to the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="abf05-222">В следующем примере в `sheetPropertiesChanged` диалоговое окно отправляются свойства листа Excel.</span><span class="sxs-lookup"><span data-stu-id="abf05-222">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="abf05-223">В этом случае текущий лист называется "Мой лист", а второй лист книги.</span><span class="sxs-lookup"><span data-stu-id="abf05-223">In this case, the current worksheet is named "My Sheet" and it's the second sheet in the workbook.</span></span> <span data-ttu-id="abf05-224">Данные инкапсулируются в объекте и преобразованного, чтобы их можно было передать `messageChild` .</span><span class="sxs-lookup"><span data-stu-id="abf05-224">The data is encapsulated in an object and stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="abf05-225">Обработка Диалогпарентмессажерецеивед в диалоговом окне</span><span class="sxs-lookup"><span data-stu-id="abf05-225">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="abf05-226">В JavaScript диалогового окна Зарегистрируйте обработчик для `DialogParentMessageReceived` события с помощью метода [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="abf05-226">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="abf05-227">Это обычно делается в [методах Office. onread или Office.iniтиализе](initialize-add-in.md), как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="abf05-227">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md), as shown in the following.</span></span> <span data-ttu-id="abf05-228">(Ниже приведен пример более надежного примера.)</span><span class="sxs-lookup"><span data-stu-id="abf05-228">(A more robust example is below.)</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="abf05-229">Затем определите `onMessageFromParent` обработчик.</span><span class="sxs-lookup"><span data-stu-id="abf05-229">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="abf05-230">Приведенный ниже код продолжает пример из предыдущего раздела.</span><span class="sxs-lookup"><span data-stu-id="abf05-230">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="abf05-231">Обратите внимание, что Office передает аргумент обработчику и что `message` свойство объекта Argument содержит строку со страницы узла.</span><span class="sxs-lookup"><span data-stu-id="abf05-231">Note that Office passes an argument to the handler and that the `message` property of the argument object contains the string from the host page.</span></span> <span data-ttu-id="abf05-232">В этом примере сообщение переводится в объект, а jQuery используется для установки верхнего заголовка диалогового окна в соответствующее имя нового листа.</span><span class="sxs-lookup"><span data-stu-id="abf05-232">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="abf05-233">Рекомендуется проверить правильность регистрации обработчика.</span><span class="sxs-lookup"><span data-stu-id="abf05-233">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="abf05-234">Для этого можно передать обратный вызов `addHandlerAsync` методу.</span><span class="sxs-lookup"><span data-stu-id="abf05-234">You can do this by passing a callback to the `addHandlerAsync` method.</span></span> <span data-ttu-id="abf05-235">Это выполняется при завершении попытки регистрации обработчика.</span><span class="sxs-lookup"><span data-stu-id="abf05-235">This runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="abf05-236">Используйте обработчик для записи или отображения ошибки, если обработчик не был успешно зарегистрирован.</span><span class="sxs-lookup"><span data-stu-id="abf05-236">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="abf05-237">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="abf05-237">The following is an example.</span></span> <span data-ttu-id="abf05-238">Обратите внимание, что `reportError` это функция, не определенная здесь, записывает или отображает сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="abf05-238">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

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

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a><span data-ttu-id="abf05-239">Диалоговое окно "Условная передача сообщений из родительской страницы"</span><span class="sxs-lookup"><span data-stu-id="abf05-239">Conditional messaging from parent page to dialog box</span></span>

<span data-ttu-id="abf05-240">Так как вы можете выполнять несколько `messageChild` вызовов со страницы узла, но у вас есть только один обработчик в диалоговом окне для `DialogParentMessageReceived` события, обработчик должен использовать условную логику для различения разных сообщений.</span><span class="sxs-lookup"><span data-stu-id="abf05-240">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="abf05-241">Это можно сделать точно так же, как при структурировании условной передачи сообщений, когда диалоговое окно отправляет сообщение на страницу узла, как описано в [условной системе обмена сообщениями](#conditional-messaging).</span><span class="sxs-lookup"><span data-stu-id="abf05-241">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](#conditional-messaging).</span></span>

> [!NOTE]
> <span data-ttu-id="abf05-242">В некоторых случаях `messageChild` API, который является частью [набора требований DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md), может не поддерживаться.</span><span class="sxs-lookup"><span data-stu-id="abf05-242">In some situations, the `messageChild` API, which is a part of the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md),  may not be supported.</span></span> <span data-ttu-id="abf05-243">Некоторые альтернативные способы обмена сообщениями с родительским диалоговым окном описаны в разделе [альтернативные способы передачи сообщений в диалоговое окно со страницы узла](parent-to-dialog.md).</span><span class="sxs-lookup"><span data-stu-id="abf05-243">Some alternative ways for parent-to-dialog-box messaging are described in [Alternative ways of passing messages to a dialog box from its host page](parent-to-dialog.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="abf05-244">[Набор требований DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md) не может быть указан в `<Requirements>` разделе манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="abf05-244">The [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md) cannot be specified in the `<Requirements>` section of an add-in manifest.</span></span> <span data-ttu-id="abf05-245">Вам потребуется проверить поддержку DialogApi 1,2 во время выполнения с помощью метода [метод issetsupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) .</span><span class="sxs-lookup"><span data-stu-id="abf05-245">You will have to check for support for DialogApi 1.2 at runtime using the [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) method.</span></span> <span data-ttu-id="abf05-246">Поддержка требований к манифесту находится на стадии разработки.</span><span class="sxs-lookup"><span data-stu-id="abf05-246">Support for manifest requirements is under development.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="abf05-247">Закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="abf05-247">Closing the dialog box</span></span>

<span data-ttu-id="abf05-p141">Вы можете добавить в диалоговое окно кнопку, которая будет его закрывать. Для этого обработчик событий кнопки должен использовать `messageParent`, чтобы сообщить главной странице, что кнопка нажата. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="abf05-p141">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="abf05-251">Обработчик главной страницы для `DialogMessageReceived` вызовет `dialog.close`, как показано в этом примере.</span><span class="sxs-lookup"><span data-stu-id="abf05-251">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="abf05-252">(См. предыдущие примеры, в которых показано, как `dialog` инициализируется объект).</span><span class="sxs-lookup"><span data-stu-id="abf05-252">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="abf05-253">Даже если у вас нет собственного пользовательского интерфейса для закрытия диалогового окна, пользователь может закрыть диалоговое окно, выбрав **X** в правом верхнем углу.</span><span class="sxs-lookup"><span data-stu-id="abf05-253">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="abf05-254">Это действие запускает событие `DialogEventReceived`.</span><span class="sxs-lookup"><span data-stu-id="abf05-254">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="abf05-255">Чтобы главная область могла реагировать на это событие, для нее должен быть объявлен обработчик этого события.</span><span class="sxs-lookup"><span data-stu-id="abf05-255">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="abf05-256">Дополнительные сведения см. в разделе [Ошибок и события в диалоговом окне](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box).</span><span class="sxs-lookup"><span data-stu-id="abf05-256">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="abf05-257">Продвинутые темы и специальные сценарии</span><span class="sxs-lookup"><span data-stu-id="abf05-257">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="abf05-258">Используйте Dialog API, чтобы показать видео</span><span class="sxs-lookup"><span data-stu-id="abf05-258">Use the Dialog API to show a video</span></span>

<span data-ttu-id="abf05-259">См. статью [Использование диалогового окна «Office» для отображения видео](dialog-video.md).</span><span class="sxs-lookup"><span data-stu-id="abf05-259">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="abf05-260">Использование Dialog API в потоке аутентификации</span><span class="sxs-lookup"><span data-stu-id="abf05-260">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="abf05-261">См. статью[ Проверка подлинности с помощью Office Dialog API ](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="abf05-261">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="abf05-262">Использование Office dialog API с одностраничными приложениями и клиентской маршрутизацией</span><span class="sxs-lookup"><span data-stu-id="abf05-262">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="abf05-263">При использовании Office dialog API, SPA и маршрутизация на стороне клиента должны обрабатываться с осторожностью</span><span class="sxs-lookup"><span data-stu-id="abf05-263">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="abf05-264">См. статью[Рекомендации по использованию Office dialog API в SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span><span class="sxs-lookup"><span data-stu-id="abf05-264">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="abf05-265">Обработка ошибок и событий</span><span class="sxs-lookup"><span data-stu-id="abf05-265">Error and event handling</span></span>

<span data-ttu-id="abf05-266">См. статью об ошибках и событиях [Обработка ошибок и событий в Office dialog box](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="abf05-266">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="abf05-267">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="abf05-267">Next steps</span></span>

<span data-ttu-id="abf05-268">Узнайте о том, как использовать Office dialog API, в [Рекомендации по использованию Office dialog API](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="abf05-268">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

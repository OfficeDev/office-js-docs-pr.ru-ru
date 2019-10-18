---
title: Использование Dialog API в надстройках Office
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 5cafb2396c92576bd5ac6d6d52105e0bb5ee579d
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302583"
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a><span data-ttu-id="543f7-102">Использование Dialog API в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="543f7-102">Use the Dialog API in your Office Add-ins</span></span>

<span data-ttu-id="543f7-p101">Вы можете использовать [Dialog API](/javascript/api/office/office.ui), чтобы открывать диалоговые окна в надстройке Office. Эта статья содержит рекомендации по использованию Dialog API в надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="543f7-p101">You can use the [Dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in. This article provides guidance for using the Dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="543f7-p102">Сведения о поддержке Dialog API см. в статье [Наборы обязательных элементов Dialog API](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). В настоящее время Dialog API поддерживается для Word, Excel, PowerPoint и Outlook.</span><span class="sxs-lookup"><span data-stu-id="543f7-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="543f7-107">Основной сценарий применения Dialog API — обеспечение проверки подлинности с использованием таких ресурсов, как Google, Facebook или Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="543f7-107">A primary scenario for the Dialog APIs is to enable authentication with a resource such as Google or Facebook.</span></span> <span data-ttu-id="543f7-108">Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md) *после* прочтения текущей статьи.</span><span class="sxs-lookup"><span data-stu-id="543f7-108">For more information, see [Authenticate with the Office Dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="543f7-109">Возможность открытия диалогового окна с помощью области задач, контентной надстройки или [команды надстройки](../design/add-in-commands.md) может позволить следующее:</span><span class="sxs-lookup"><span data-stu-id="543f7-109">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="543f7-110">отобразить страницу входа, которую невозможно открыть непосредственно в области задач;</span><span class="sxs-lookup"><span data-stu-id="543f7-110">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="543f7-111">предоставить больше места на экране (или даже весь экран) для некоторых задач в надстройке;</span><span class="sxs-lookup"><span data-stu-id="543f7-111">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="543f7-112">разместить видео, которое будет слишком маленьким в области задач.</span><span class="sxs-lookup"><span data-stu-id="543f7-112">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="543f7-p104">Так как пользователей раздражают элементы интерфейса, перекрывающие основное содержимое, не допускайте открытия диалогового окна из области задач, если этого не требует сценарий. При планировании контактной зоны помните, что в области задач можно использовать вкладки. Например, как в [надстройке SalesTracker на JavaScript для Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span><span class="sxs-lookup"><span data-stu-id="543f7-p104">Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="543f7-116">На приведенном ниже изображении показан пример диалогового окна. </span><span class="sxs-lookup"><span data-stu-id="543f7-116">The following image shows an example of a dialog box.</span></span>

![Команды надстроек](../images/auth-o-dialog-open.png)

<span data-ttu-id="543f7-p105">Обратите внимание на то, что диалоговое окно всегда открывается в центре экрана. Пользователь может перемещать его и изменять его размер. Окно *не модальное*: пользователь может продолжать работать и с документом в ведущем приложении Office, и с главной страницей в области задач.</span><span class="sxs-lookup"><span data-stu-id="543f7-p105">Note that the dialog box always opens in the center of the screen. The user can move and resize it. The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the host page in the task pane, if there is one.</span></span>

## <a name="dialog-api-scenarios"></a><span data-ttu-id="543f7-121">Сценарии с Dialog API</span><span class="sxs-lookup"><span data-stu-id="543f7-121">Dialog API scenarios</span></span>

<span data-ttu-id="543f7-122">Интерфейсы API JavaScript для Office поддерживают указанные ниже сценарии с объектом [Dialog](/javascript/api/office/office.dialog) и две функции в [пространстве имен Office.context.ui](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="543f7-122">The Office JavaScript APIs support the following scenarios with a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

### <a name="open-a-dialog-box"></a><span data-ttu-id="543f7-123">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="543f7-123">Open a dialog box</span></span>

<span data-ttu-id="543f7-p106">Чтобы открыть диалоговое окно, код в области задач вызывает метод [displayDialogAsync](/javascript/api/office/office.ui) и передает ему URL-адрес ресурса, который нужно открыть. Таким ресурсом обычно является страница, но может быть и метод контроллера в приложении MVC, метод веб-службы, маршрута или любой другой ресурс. В этой статье термин "страница" или "веб-сайт" означает ресурс в диалоговом окне. Ниже приведен простой пример кода.</span><span class="sxs-lookup"><span data-stu-id="543f7-p106">To open a dialog box, your code in the task pane calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open. This is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource. In this article, 'page' or 'website' refers to the resource in the dialog. The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="543f7-p107">В случае URL-адреса используется протокол HTTP**S**, обязательный для всех страниц, загружаемых в диалоговом окне, а не только для первой страницы.</span><span class="sxs-lookup"><span data-stu-id="543f7-p107">The URL uses the HTTP**S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="543f7-130">Домен ресурса диалоговых окон совпадает с доменом страницы ведущего приложения, которая может быть страницей в области задач или [файле функций](/office/dev/add-ins/reference/manifest/functionfile) для команды надстройки.</span><span class="sxs-lookup"><span data-stu-id="543f7-130">The dialog resource's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](/office/dev/add-ins/reference/manifest/functionfile) of an add-in command.</span></span> <span data-ttu-id="543f7-131">Страница, метод контроллера или другой ресурс, передаваемый в метод `displayDialogAsync`, должен быть в том же домене, что и страница ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="543f7-131">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="543f7-132">Страница ведущего приложения и ресурсы диалоговых окон должны иметь одинаковые полные доменные имена.</span><span class="sxs-lookup"><span data-stu-id="543f7-132">The host page and the resources of the dialog must have the same full domain.</span></span> <span data-ttu-id="543f7-133">Если вы попробуете передать поддомен домена надстройки в `displayDialogAsync`, ничего не получится.</span><span class="sxs-lookup"><span data-stu-id="543f7-133">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="543f7-134">Полные доменные имена, включая поддомены, должны совпадать.</span><span class="sxs-lookup"><span data-stu-id="543f7-134">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="543f7-p110">После загрузки первой страницы (или другого ресурса) пользователь может перейти к любому веб-сайту (или другому ресурсу), который использует HTTPS. Первая страница также может сразу перенаправлять пользователя на другой сайт.</span><span class="sxs-lookup"><span data-stu-id="543f7-p110">After the first page (or other resource) is loaded, a user can go to any website (or other resource) that uses HTTPS. You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="543f7-137">По умолчанию диалоговое окно занимает 80 % высоты и ширины экрана устройства, но вы можете установить другие соотношения путем передачи объекта конфигурации в метод, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="543f7-137">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="543f7-138">Подобная надстройка приведена в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="543f7-138">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="543f7-p111">Установите оба значения равными 100 %, чтобы надстройка открывалась во весь экран. (На самом деле, максимальное значение составляет 99,5 %, возможность перемещать окно и изменять его размер сохраняется.)</span><span class="sxs-lookup"><span data-stu-id="543f7-p111">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="543f7-p112">Из главного окна можно открыть только одно диалоговое окно. При попытке открыть еще одно диалоговое окно произойдет ошибка. Поэтому если пользователь, например, откроет диалоговое окно из области задач, он не сможет открыть второе диалоговое окно на другой странице в области задач. Но при открытии диалогового окна с помощью [команды надстройки](../design/add-in-commands.md) каждый раз открывается новый (невидимый) HTML-файл. При этом создается новое (невидимое) главное окно, которое может запускать собственное диалоговое окно. Дополнительные сведения см. в разделе [Ошибки метода displayDialogAsync](#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="543f7-p112">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="543f7-147">Использование параметра производительности в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="543f7-147">Take advantage of a performance option in Office Online</span></span>

<span data-ttu-id="543f7-148">`displayInIframe` — дополнительное свойство в объекте конфигурации, которое можно передать `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="543f7-148">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="543f7-149">Когда этому свойству присвоено значение `true`, а надстройка запущена для документа в Office в Интернете, диалоговое окно будет открываться быстрее, потому что будет выступать как плавающий фрейм iframe.</span><span class="sxs-lookup"><span data-stu-id="543f7-149">When this property is set to `true`, and the add-in is running in a document opened in Office Online, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="543f7-150">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="543f7-150">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="543f7-151">Значение по умолчанию: `false`. Его использование равнозначно пропуску всего свойства.</span><span class="sxs-lookup"><span data-stu-id="543f7-151">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="543f7-152">Если надстройка не работает в Office в Интернете, `displayInIframe` игнорируется.</span><span class="sxs-lookup"><span data-stu-id="543f7-152">If the add-in is not running in Office Online, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="543f7-p115">**Не** следует использовать `displayInIframe: true`, если диалоговое окно будет выполнять перенаправление на страницу, которую невозможно открыть в элементе iframe. Например, страницы входа многих популярных веб-служб (который выполняется, например, с помощью учетной записи Майкрософт или Google), невозможно открыть в элементе iframe.</span><span class="sxs-lookup"><span data-stu-id="543f7-p115">You should **not** use `displayInIframe: true` if the dialog will at any point redirect to a page that cannot be opened in an iframe. For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a><span data-ttu-id="543f7-155">Обработка блокировщиков всплывающих окон с помощью Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="543f7-155">Handling pop-up blockers with Office on the web</span></span>

<span data-ttu-id="543f7-156">Попытка отобразить диалоговое окно при использовании Office в Интернете может вызвать блокировку этого диалогового окна блокировщиком всплывающих окон браузера.</span><span class="sxs-lookup"><span data-stu-id="543f7-156">Attempting to display a dialog while using Office Online may cause the browser's pop-up blocker to block the dialog.</span></span> <span data-ttu-id="543f7-157">Блокировщик всплывающих окон браузера можно обойти, если пользователь надстройки сначала предоставит согласие на запрос от надстройки.</span><span class="sxs-lookup"><span data-stu-id="543f7-157">The browser's pop-up blocker can be circumvented if the user of your add-in first agrees to a prompt from the add-in.</span></span> <span data-ttu-id="543f7-158">У объекта [DialogOptions](/javascript/api/office/office.dialogoptions) метода `displayDialogAsync` есть свойство `promptBeforeOpen`, вызывающее такое всплывающее окно.</span><span class="sxs-lookup"><span data-stu-id="543f7-158">`displayDialogAsync`'s [DialogOptions](/javascript/api/office/office.dialogoptions) has the `promptBeforeOpen` property to trigger such a pop-up.</span></span> <span data-ttu-id="543f7-159">`promptBeforeOpen` предоставляет собой логическое значение, обеспечивающее указанное ниже поведение.</span><span class="sxs-lookup"><span data-stu-id="543f7-159">`promptBeforeOpen` is a boolean value which provides the following behavior:</span></span>

 - <span data-ttu-id="543f7-160">`true` — платформа отображает всплывающее окно, чтобы запустить навигацию и обойти блокировщик всплывающих окон браузера.</span><span class="sxs-lookup"><span data-stu-id="543f7-160">`true` - The framework displays a pop-up to trigger the navigation and avoid the browser's pop-up blocker.</span></span> 
 - <span data-ttu-id="543f7-161">`false` — диалоговое окно не отображается. Всплывающие окна должны обрабатываться разработчиком (путем предоставления артефакта пользовательского интерфейса для запуска навигации).</span><span class="sxs-lookup"><span data-stu-id="543f7-161">`false` - The dialog will not be shown and the developer must handle pop-ups (by providing a user interface artifact to trigger the navigation).</span></span> 
 
<span data-ttu-id="543f7-162">Всплывающее окно аналогично представленному на снимке экрана ниже.</span><span class="sxs-lookup"><span data-stu-id="543f7-162">The pop-up looks similiar to that in the following screenshot:</span></span>

![Запрос, который может создавать диалоговое окно надстройки, чтобы обойти блокировщик всплывающих окон браузера.](../images/dialog-prompt-before-open.png)
 
### <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="543f7-164">Отправка сведений из диалогового окна главной странице</span><span class="sxs-lookup"><span data-stu-id="543f7-164">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="543f7-165">Диалоговое окно может взаимодействовать с главной страницей в области задач, если:</span><span class="sxs-lookup"><span data-stu-id="543f7-165">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="543f7-166">Текущая страница в диалоговом окне не находится в том же домене, что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="543f7-166">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="543f7-p117">Библиотека JavaScript для Office не загружена на странице. (Как и любая страница, которая использует библиотеку JavaScript для Office, сценарий для страницы должен назначить метод свойству `Office.initialize`. Метод может быть пустой. Дополнительные сведения см. в разделе [Инициализация надстройки](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span><span class="sxs-lookup"><span data-stu-id="543f7-p117">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="543f7-p118">Код в диалоговом окне использует функцию `messageParent` для отправки логического значения или строки на главную страницу. Строка может быть словом, предложением, большим двоичным объектом XML, строковым представлением JSON или любым другим объектом, который можно сериализовать, представив в виде строки. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="543f7-p118">Code in the dialog page uses the `messageParent` function to send either a Boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="543f7-p119">Функция `messageParent` — это один из *двух* API Office, которые можно вызывать в диалоговом окне. Другой — `Office.context.requirements.isSetSupported`. Дополнительные сведения см. в статье [Указание ведущих приложений Office и требований к API](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="543f7-p119">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="543f7-176">Функцию `messageParent` можно вызывать только на странице, которая относится к тому же домену (включая протокол и порт), что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="543f7-176">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="543f7-177">В следующем примере `googleProfile` — это строковое представление профиля Google пользователя.</span><span class="sxs-lookup"><span data-stu-id="543f7-177">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="543f7-p120">Чтобы главная страница получила сообщение, ее необходимо настроить. Для этого добавьте параметр обратного вызова в исходный вызов `displayDialogAsync`. Обратный вызов назначает событию `DialogMessageReceived` обработчик. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="543f7-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="543f7-p121">Office передает объект [AsyncResult](/javascript/api/office/office.asyncresult) в функцию обратного вызова. Он представляет собой результат попытки открыть диалоговое окно, но не результат событий в диалоговом окне. Дополнительные сведения об этой особенности см. в разделе [Обработка ошибок и событий](#handle-errors-and-events).</span><span class="sxs-lookup"><span data-stu-id="543f7-p121">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback. It represents the result of the attempt to open the dialog box. It does not represent the outcome of any events in the dialog box. For more on this distinction, see the section [Handle errors and events](#handle-errors-and-events).</span></span>
> - <span data-ttu-id="543f7-186">Для свойства `value` объекта `asyncResult` задан объект [Dialog](/javascript/api/office/office.dialog), который существует на главной странице, а не в контексте выполнения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="543f7-186">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="543f7-p122">`processMessage` — это функция, которая обрабатывает событие. Вы можете присвоить ей любое имя.</span><span class="sxs-lookup"><span data-stu-id="543f7-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="543f7-189">Переменная `dialog` объявляется в более широком контексте, чем обратный вызов, так как на нее также ссылается `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="543f7-189">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="543f7-190">Ниже приведен простой пример обработчика для события `DialogMessageReceived`.</span><span class="sxs-lookup"><span data-stu-id="543f7-190">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="543f7-p123">Office передает объект `arg` в обработчик. Его свойство `message` — это логическое значение или строка, отправляемая при вызове `messageParent` в диалоговом окне. В данном примере это профиль пользователя из учетной записи Майкрософт, Google или другой службы, представленный в виде строки, поэтому он десериализируется обратно в объект с помощью метода `JSON.parse`.</span><span class="sxs-lookup"><span data-stu-id="543f7-p123">Office passes the `arg` object to the handler. Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="543f7-p124">Функция `showUserName` не показана. Она может отображать персонализированное приветствие в области задач.</span><span class="sxs-lookup"><span data-stu-id="543f7-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="543f7-196">Когда взаимодействие пользователя с диалоговым окном закончится, обработчик сообщений должен закрыть диалоговое окно, как показано в этом примере.</span><span class="sxs-lookup"><span data-stu-id="543f7-196">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="543f7-197">Объект `dialog` должен быть таким же, как объект, который возвращается при вызове `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="543f7-197">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="543f7-198">Вызов метода `dialog.close` дает указание Office немедленно закрыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="543f7-198">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="543f7-199">Пример надстройки, в которой используются эти методы, см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="543f7-199">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="543f7-p125">Если надстройка должна открыть другую страницу области задач после получения сообщения, можно использовать метод `window.location.replace` (или `window.location.href`) в последней строке обработчика. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="543f7-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="543f7-202">Пример подобной надстройки см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="543f7-202">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

#### <a name="conditional-messaging"></a><span data-ttu-id="543f7-203">Условные сообщения</span><span class="sxs-lookup"><span data-stu-id="543f7-203">Conditional messaging</span></span>

<span data-ttu-id="543f7-p126">Так как из диалогового окна можно отправить несколько вызовов `messageParent`, но на главной странице есть только один обработчик для события `DialogMessageReceived`, обработчику необходимо использовать условную логику, чтобы различать сообщения. Например, если диалоговое окно предлагает пользователю войти в учетную запись Майкрософт, Google или другого поставщика удостоверений, оно отправляет профиль пользователя в виде сообщения. Если выполнить аутентификацию не удается, диалоговое окно отправляет сведения об ошибке на главную страницу, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="543f7-p126">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="543f7-207">Переменная `loginSuccess` будет инициализирована после считывания отклика HTTP от поставщика удостоверений.</span><span class="sxs-lookup"><span data-stu-id="543f7-207">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="543f7-p127">Реализация функций `getProfile` и `getError` не показана. Они получают данные из параметра запроса или ответа HTTP.</span><span class="sxs-lookup"><span data-stu-id="543f7-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="543f7-p128">В зависимости от того, удалось ли выполнить вход, отправляются анонимные объекты различных типов. Оба содержат свойство `messageType`, но один содержит свойство `profile`, а другой — свойство `error`.</span><span class="sxs-lookup"><span data-stu-id="543f7-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="543f7-p129">Код обработчика на главной странице использует значение свойства `messageType` для разветвления, как показано в приведенном ниже примере. Обратите внимание на то, что здесь используется та же функция `showUserName`, что и в примере выше, а функция `showNotification` отображает сообщение об ошибке в элементе пользовательского интерфейса на главной странице.</span><span class="sxs-lookup"><span data-stu-id="543f7-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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
> <span data-ttu-id="543f7-214">Реализация функции `showNotification` не показана в примере кода, представленном в этой статье.</span><span class="sxs-lookup"><span data-stu-id="543f7-214">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="543f7-215">Пример возможного способа реализации этой функции в своей надстройке см. в статье [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="543f7-215">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

### <a name="closing-the-dialog-box"></a><span data-ttu-id="543f7-216">Закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="543f7-216">Closing the dialog box</span></span>

<span data-ttu-id="543f7-p131">Вы можете добавить в диалоговое окно кнопку, которая будет его закрывать. Для этого обработчик событий кнопки должен использовать `messageParent`, чтобы сообщить главной странице, что кнопка нажата. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="543f7-p131">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="543f7-p132">Обработчик главной страницы для `DialogMessageReceived` вызовет `dialog.close`, как показано в этом примере. (Примеры инициализации объекта dialog см. выше в этой статье.)</span><span class="sxs-lookup"><span data-stu-id="543f7-p132">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example. (See previous examples that show how the dialog object is initialized.)</span></span>


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="543f7-p133">Даже если у вас нет собственной кнопки для закрытия диалогового окна, пользователь сможет закрыть его, нажав кнопку **X** в правом верхнем углу. Это действие запускает событие `DialogEventReceived`. Чтобы главная область могла реагировать на это событие, для нее должен быть объявлен обработчик этого события. Подробнее: [Ошибки и события в диалоговом окне](#errors-and-events-in-the-dialog-window).</span><span class="sxs-lookup"><span data-stu-id="543f7-p133">Even when you don't have your own close dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens, it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.</span></span>

## <a name="handle-errors-and-events"></a><span data-ttu-id="543f7-226">Обработка ошибок и событий</span><span class="sxs-lookup"><span data-stu-id="543f7-226">Handle errors and events</span></span>

<span data-ttu-id="543f7-227">Код должен обрабатывать две категории событий:</span><span class="sxs-lookup"><span data-stu-id="543f7-227">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="543f7-228">Ошибки, возвращаемые при вызове метода `displayDialogAsync`, так как не удается создать диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="543f7-228">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="543f7-229">Ошибки и другие события в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="543f7-229">Errors, and other events, in the dialog window.</span></span>

### <a name="errors-from-displaydialogasync"></a><span data-ttu-id="543f7-230">Ошибки метода displayDialogAsync</span><span class="sxs-lookup"><span data-stu-id="543f7-230">Errors from displayDialogAsync</span></span>

<span data-ttu-id="543f7-231">Кроме общих ошибок платформы и системы, при вызове метода `displayDialogAsync` возникают указанные ниже ошибки.</span><span class="sxs-lookup"><span data-stu-id="543f7-231">In addition to general platform and system errors, three errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="543f7-232">Цифровой код</span><span class="sxs-lookup"><span data-stu-id="543f7-232">Code number</span></span>|<span data-ttu-id="543f7-233">Значение</span><span class="sxs-lookup"><span data-stu-id="543f7-233">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="543f7-234">12004</span><span class="sxs-lookup"><span data-stu-id="543f7-234">12004</span></span>|<span data-ttu-id="543f7-p134">Домен URL-адреса, передаваемого в метод `displayDialogAsync`, не является доверенным. Домен должен быть таким же, как и для главной страницы (а также протокол и номер порта).</span><span class="sxs-lookup"><span data-stu-id="543f7-p134">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="543f7-237">12005</span><span class="sxs-lookup"><span data-stu-id="543f7-237">12005</span></span>|<span data-ttu-id="543f7-p135">URL-адрес, передаваемый в метод `displayDialogAsync`, использует протокол HTTP. Необходим протокол HTTPS. (В некоторых версиях Office сообщение об ошибке 12005 совпадает с сообщением 12004.)</span><span class="sxs-lookup"><span data-stu-id="543f7-p135">The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="543f7-241"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="543f7-241"><span id="12007">12007</span></span></span>|<span data-ttu-id="543f7-p136">Диалоговое окно уже открыто из этого главного окна. Для главного окна, например области задач, невозможно открыть сразу несколько диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="543f7-p136">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="543f7-244">12009</span><span class="sxs-lookup"><span data-stu-id="543f7-244">12009</span></span>|<span data-ttu-id="543f7-245">Пользователь проигнорировал диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="543f7-245">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="543f7-246">Эта ошибка может возникнуть в веб-версиях Office, где пользователи могут не разрешить надстройке открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="543f7-246">This error can occur in online versions of Office, where users may choose not to allow an add-in to present a dialog.</span></span>|

<span data-ttu-id="543f7-247">При вызове `displayDialogAsync` он всегда передает объект [AsyncResult](/javascript/api/office/office.asyncresult) в функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="543f7-247">When `displayDialogAsync` is called, it always passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="543f7-248">Если вызов выполнен, т. е. диалоговое окно открыто, свойство `value` объекта `AsyncResult` представляет собой объект [Dialog](/javascript/api/office/office.dialog).</span><span class="sxs-lookup"><span data-stu-id="543f7-248">When the call is successful - that is, the dialog window is opened - the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="543f7-249">См. пример в разделе [Отправка данных из диалогового окна на страницу ведущего приложения](#send-information-from-the-dialog-box-to-the-host-page).</span><span class="sxs-lookup"><span data-stu-id="543f7-249">An example of this is in the section [Send information from the dialog box to the host page](#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="543f7-250">Если вызвать `displayDialogAsync` не удается, то окно не создается, свойству `status` объекта `AsyncResult` присваивается значение `Office.AsyncResultStatus.Failed`, а также заполняется свойство `error` объекта.</span><span class="sxs-lookup"><span data-stu-id="543f7-250">When the call to `displayDialogAsync` fails, the window is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="543f7-251">У вас всегда должна быть функция обратного вызова, которая проверяет `status` и сообщает об ошибке.</span><span class="sxs-lookup"><span data-stu-id="543f7-251">You should always have a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="543f7-252">Ниже приведен пример кода, сообщающий об ошибке, независимо от ее кода.</span><span class="sxs-lookup"><span data-stu-id="543f7-252">For an example that simply reports the error message regardless of its code number, see the following code:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a><span data-ttu-id="543f7-253">Ошибки и события в диалоговом окне</span><span class="sxs-lookup"><span data-stu-id="543f7-253">Errors and events in the dialog window</span></span>

<span data-ttu-id="543f7-254">Три ошибки и события, известных по цифровым кодам, в диалоговом окне вызывают событие `DialogEventReceived` на главной странице.</span><span class="sxs-lookup"><span data-stu-id="543f7-254">Three errors and events, known by their code numbers, in the dialog box will trigger a `DialogEventReceived` event in the host page.</span></span>

|<span data-ttu-id="543f7-255">Цифровой код</span><span class="sxs-lookup"><span data-stu-id="543f7-255">Code number</span></span>|<span data-ttu-id="543f7-256">Значение</span><span class="sxs-lookup"><span data-stu-id="543f7-256">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="543f7-257">12002</span><span class="sxs-lookup"><span data-stu-id="543f7-257">12002</span></span>|<span data-ttu-id="543f7-258">Одно из следующих:</span><span class="sxs-lookup"><span data-stu-id="543f7-258">One of the following:</span></span><br> <span data-ttu-id="543f7-259">– По URL-адресу, переданному в `displayDialogAsync`, не существует страницы.</span><span class="sxs-lookup"><span data-stu-id="543f7-259">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="543f7-260">– Страница, переданная в метод `displayDialogAsync`, загружена, но выполнена попытка открыть из диалогового окна страницу, которую не удается найти или загрузить, или для которой указан URL-адрес с недопустимым синтаксисом.</span><span class="sxs-lookup"><span data-stu-id="543f7-260">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="543f7-261">12003</span><span class="sxs-lookup"><span data-stu-id="543f7-261">12003</span></span>|<span data-ttu-id="543f7-p139">Выполнена попытка открыть из диалогового окна страницу, для URL-адреса которой используется протокол HTTP. Необходим протокол HTTPS.</span><span class="sxs-lookup"><span data-stu-id="543f7-p139">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="543f7-264">12006</span><span class="sxs-lookup"><span data-stu-id="543f7-264">12006</span></span>|<span data-ttu-id="543f7-265">Диалоговое окно закрыто. Скорее всего, пользователь нажал кнопку **X**.</span><span class="sxs-lookup"><span data-stu-id="543f7-265">The dialog box was closed, usually because the user chooses the **X** button.</span></span>|

<span data-ttu-id="543f7-p140">Код может назначить обработчик для события `DialogEventReceived` при вызове `displayDialogAsync`. Ниже приведен простой пример.</span><span class="sxs-lookup"><span data-stu-id="543f7-p140">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="543f7-268">Ниже приведен пример обработчика для события `DialogEventReceived`, который создает особые сообщения об ошибках для каждого кода ошибки.</span><span class="sxs-lookup"><span data-stu-id="543f7-268">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

<span data-ttu-id="543f7-269">Надстройку с такой обработкой ошибок см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="543f7-269">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>


## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="543f7-270">Передача данных диалоговому окну</span><span class="sxs-lookup"><span data-stu-id="543f7-270">Pass information to the dialog box</span></span>

<span data-ttu-id="543f7-p141">Иногда главной странице нужно передать данные в диалоговое окно. Есть два основных способа обеспечить эту возможность:</span><span class="sxs-lookup"><span data-stu-id="543f7-p141">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="543f7-273">Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="543f7-273">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="543f7-p142">Храните информацию в месте, доступном как для главного, так и для диалогового окна. У всех окон есть отдельное хранилище сеанса, но *если для них используется один домен* (включая номер порта), у них общее [локальное хранилище](https://www.w3schools.com/html/html5_webstorage.asp).</span><span class="sxs-lookup"><span data-stu-id="543f7-p142">Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any),  they share a common [local storage](https://www.w3schools.com/html/html5_webstorage.asp).</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="543f7-276">Использование локального хранилища</span><span class="sxs-lookup"><span data-stu-id="543f7-276">Use local storage</span></span>

<span data-ttu-id="543f7-277">Чтобы использовать локальное хранилище, код вызывает метод `setItem` объекта `window.localStorage` на главной странице перед вызовом `displayDialogAsync`, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="543f7-277">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="543f7-278">Код в диалоговом окне считывает элемент, когда это необходимо, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="543f7-278">Code in the dialog window reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a><span data-ttu-id="543f7-279">Использование параметров запроса</span><span class="sxs-lookup"><span data-stu-id="543f7-279">Use query parameters</span></span>

<span data-ttu-id="543f7-280">В приведенном ниже примере показано, как передавать данные с помощью параметра запроса.</span><span class="sxs-lookup"><span data-stu-id="543f7-280">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="543f7-281">Пример, в котором используется эта техника, см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="543f7-281">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="543f7-282">Код в диалоговом окне может проанализировать URL-адрес и считать значение параметра.</span><span class="sxs-lookup"><span data-stu-id="543f7-282">Code in your dialog window can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="543f7-p143">Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`. (Этот параметр добавляется после пользовательских параметров запроса, если они есть. Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать. То же значение добавляется в хранилище сеанса диалогового окна. *Ваш код не должен ни считывать это значение, ни записывать в него данные*.</span><span class="sxs-lookup"><span data-stu-id="543f7-p143">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="use-the-dialog-apis-to-show-a-video"></a><span data-ttu-id="543f7-288">Использование Dialog API для показа видео</span><span class="sxs-lookup"><span data-stu-id="543f7-288">Use the Dialog APIs to show a video</span></span>

<span data-ttu-id="543f7-289">Чтобы показать видео в диалоговом окне:</span><span class="sxs-lookup"><span data-stu-id="543f7-289">To show a video in a dialog box:</span></span>

1.  <span data-ttu-id="543f7-p144">Создайте страницу с единственным содержимым — элементом iframe. Атрибут `src` этого элемента указывает на видео из Интернета. В URL-адресе видео должен быть указан протокол HTTP**S**. В этой статье мы назовем эту страницу "video.dialogbox.html". Ниже приведен пример кода.</span><span class="sxs-lookup"><span data-stu-id="543f7-p144">Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  <span data-ttu-id="543f7-295">Страница video.dialogbox.html должна находиться в том же домене, что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="543f7-295">The video.dialogbox.html page must be in the same domain as the host page.</span></span>
3.  <span data-ttu-id="543f7-296">Используйте вызов `displayDialogAsync` на главной странице, чтобы открыть страницу video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="543f7-296">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
4.  <span data-ttu-id="543f7-p145">Если надстройке необходимо знать, когда пользователь закрывает диалоговое окно, зарегистрируйте обработчик для события `DialogEventReceived` и обработайте событие 12006. Дополнительные сведения см. в разделе [Ошибки и события в диалоговом окне](#errors-and-events-in-the-dialog-window).</span><span class="sxs-lookup"><span data-stu-id="543f7-p145">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).</span></span>

<span data-ttu-id="543f7-299">Пример видео в диалоговом окне см. в статье [Конструктивный шаблон размещения видео](/office/dev/add-ins/design/first-run-experience-patterns#video-placemat).</span><span class="sxs-lookup"><span data-stu-id="543f7-299">For a sample that shows a video in a dialog box, see the [video placemat design pattern](/office/dev/add-ins/design/first-run-experience-patterns#video-placemat).</span></span>

![Снимок экрана: видео в диалоговом окне надстройки](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="543f7-301">Использование Dialog API в потоке аутентификации</span><span class="sxs-lookup"><span data-stu-id="543f7-301">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="543f7-302">См. статью [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="543f7-302">See [Authenticate with the Office Dialog API](auth-with-office-dialog-api.md).</span></span>

## <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="543f7-303">Использование Dialog API для Office с одностраничными приложениями и клиентской маршрутизацией</span><span class="sxs-lookup"><span data-stu-id="543f7-303">Using the Office Dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="543f7-304">Если надстройка использует клиентскую маршрутизацию подобно тому, как это делает одностраничное приложение (SPA), вы можете передавать в метод [displayDialogAsync](/javascript/api/office/office.ui) (*который мы не рекомендуем использовать*) не URL-адрес отдельной HTML-страницы, а URL-адрес маршрута.</span><span class="sxs-lookup"><span data-stu-id="543f7-304">If your add-in uses client-side routing, as single-page applications typically do, you have the option to pass the URL of a route to the [displayDialogAsync](/javascript/api/office/office.ui) method, instead of the URL of a complete and separate HTML page.</span></span>

<span data-ttu-id="543f7-305">Диалоговое окно — это новое окно с собственным контекстом выполнения.</span><span class="sxs-lookup"><span data-stu-id="543f7-305">The dialog box is in a new window with its own execution context.</span></span> <span data-ttu-id="543f7-306">Если вы передаете маршрут, базовая страница со всем ее кодом инициализации и начальной загрузки запускается снова в этом новом контексте, а возможным переменным присваиваются первоначальные значения в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="543f7-306">If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window.</span></span> <span data-ttu-id="543f7-307">Такой способ приводит к скачиванию и запуску второго экземпляра приложения в диалоговом окне, что частично противоречит смыслу одностраничного приложения.</span><span class="sxs-lookup"><span data-stu-id="543f7-307">So this technique downloads and launches a second instance of your application in the dialog window, which partially defeats the purpose of an SPA.</span></span> <span data-ttu-id="543f7-308">Кроме того, код, меняющий переменные в диалоговом окне, не меняет версию области задач этих переменных.</span><span class="sxs-lookup"><span data-stu-id="543f7-308">Code that changes variables in the dialog window does not change the task pane version of the same variables.</span></span> <span data-ttu-id="543f7-309">Для диалогового окна предусмотрено отдельное хранилище сеанса, недоступное из кода в области задач.</span><span class="sxs-lookup"><span data-stu-id="543f7-309">Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.</span></span>

<span data-ttu-id="543f7-310">Поэтому если вы передавали маршрут методу `displayDialogAsync`, вы в действительности запускали не одностраничное предложение, а использовали два экземпляра одного одностраничного приложения.</span><span class="sxs-lookup"><span data-stu-id="543f7-310">So, if you passed a route to the `displayDialogAsync` method, you wouldn't really have an SPA; you'd have two instances of the same SPA.</span></span> <span data-ttu-id="543f7-311">Кроме того, большая часть кода в экземпляре области задач и большая часть кода в экземпляре диалогового окна никогда не применялись в соответствующих экземплярах.</span><span class="sxs-lookup"><span data-stu-id="543f7-311">Moreover, much of the code in the task pane instance would never be used in that instance and much of the code in the dialog instance would never be used in that instance.</span></span> <span data-ttu-id="543f7-312">Это соответствует применению двух одностраничных приложений в одном пакете.</span><span class="sxs-lookup"><span data-stu-id="543f7-312">It would be like having two SPAs in the same bundle.</span></span> <span data-ttu-id="543f7-313">Если код, который нужно выполнить в диалоговом окне достаточно сложен, рекомендуется выполнить его явным образом, то есть разместить два одностраничных приложения в разных папках одного домена.</span><span class="sxs-lookup"><span data-stu-id="543f7-313">If the code that you want to run in the dialog is sufficiently complex, you might want to do this explicitly; that is, have two SPAs in different folders of the same domain.</span></span> <span data-ttu-id="543f7-314">Но в большинстве случаев в диалоговом окне требуется только простая логика.</span><span class="sxs-lookup"><span data-stu-id="543f7-314">But in most scenarios, only simple logic is needed in the dialog.</span></span> <span data-ttu-id="543f7-315">В таких случаях проект значительно упрощается благодаря размещению простой HTML-страницы с внедренным или упомянутым кодом JavaScript в домене вашего одностраничного приложения.</span><span class="sxs-lookup"><span data-stu-id="543f7-315">In such cases, your project will be greatly simplified by simply hosting a simple HTML page, with embedded or referenced JavaScript, in the domain of your SPA.</span></span> <span data-ttu-id="543f7-316">Передайте URL-адрес страницы в метод `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="543f7-316">Pass the URL of the page to the `displayDialogAsync` method.</span></span> <span data-ttu-id="543f7-317">Это может означать, что вы отклоняетесь от буквальной идеи одностраничного приложения. Но как указано выше, при использовании диалогового окна вы в любом случае применяете не один экземпляр одностраничного приложения.</span><span class="sxs-lookup"><span data-stu-id="543f7-317">This might mean that you are deviating from the literal idea of a single-page app; but as noted above you don't really have a single instance of an SPA anyway when you are using the dialog.</span></span>

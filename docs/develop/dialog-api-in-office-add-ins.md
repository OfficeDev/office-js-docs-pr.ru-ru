---
title: Использование Office Dialog API в вашей надстройках Office
description: Узнайте, как создать диалоговое окно в надстройке Office.
ms.date: 01/28/2021
localization_priority: Normal
ms.openlocfilehash: 9061b4c048a133572e615152d61df611e5f15068
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237868"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="4b7fd-103">Использование Office Dialog API в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="4b7fd-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="4b7fd-104">Вы можете использовать [Office dialog API](/javascript/api/office/office.ui), чтобы открывать диалоговые окна в надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="4b7fd-105">Эта статья содержит инструкции по использованию dialog API в надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="4b7fd-106">Сведения о том, где в настоящее время поддерживается Dialog API, см. в наборах требований [Dialog API.](../reference/requirement-sets/dialog-api-requirement-sets.md)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-106">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](../reference/requirement-sets/dialog-api-requirement-sets.md).</span></span> <span data-ttu-id="4b7fd-107">Dialog API в настоящее время поддерживается для Excel, PowerPoint и Word.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-107">The Dialog API is currently supported for Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="4b7fd-108">Поддержка Outlook включена в различные наборы требований почтовых ящиков, дополнительные сведения см. в справочнике &mdash; по API.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-108">Outlook support is included across various Mailbox requirement sets&mdash;see the API reference for more details.</span></span>

<span data-ttu-id="4b7fd-109">Основной сценарий для Dialog API - включить аутентификацию с помощью таких ресурсов, как Google, Facebook или Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-109">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="4b7fd-110">Дополнительные сведения см. в статье [Проверка подлинности с помощью Office Dialog API](auth-with-office-dialog-api.md) *после* ознакомления с текущей статьей.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-110">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="4b7fd-111">Возможность открытия диалогового окна с помощью области задач, контентной надстройки или [команды надстройки](../design/add-in-commands.md) может позволить следующее:</span><span class="sxs-lookup"><span data-stu-id="4b7fd-111">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="4b7fd-112">отобразить страницу входа, которую невозможно открыть непосредственно в области задач;</span><span class="sxs-lookup"><span data-stu-id="4b7fd-112">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="4b7fd-113">предоставить больше места на экране (или даже весь экран) для некоторых задач в надстройке;</span><span class="sxs-lookup"><span data-stu-id="4b7fd-113">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="4b7fd-114">разместить видео, которое будет слишком маленьким в области задач.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-114">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="4b7fd-115">Поскольку перекрывающиеся элементы пользовательского интерфейса не приветствуются, избегайте открытия диалогового окна на панели задач, если это не требуется в сценарий.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-115">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="4b7fd-116">При планировании контактной зоны помните, что в области задач можно использовать вкладки.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-116">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="4b7fd-117">Пример области задач с вкладками см. в примере Надстройка [Excel JavaScript SalesTracker.](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-117">For an example of a tabbed task pane, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="4b7fd-118">На приведенном ниже изображении показан пример диалогового окна. </span><span class="sxs-lookup"><span data-stu-id="4b7fd-118">The following image shows an example of a dialog box.</span></span>

![Screenshot showing dialog with 3 sign-in options displayed in front of Word](../images/auth-o-dialog-open.png)

<span data-ttu-id="4b7fd-120">Обратите внимание, что диалоговое окно всегда открывается в центре экрана.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-120">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="4b7fd-121">Пользователь может перемещать ее и изменять ее размер.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-121">The user can move and resize it.</span></span> <span data-ttu-id="4b7fd-122">Это немодульное окно— пользователь может продолжать взаимодействовать как с документом в приложении Office, так и со страницей в области задач, если она существует. </span><span class="sxs-lookup"><span data-stu-id="4b7fd-122">The window is *nonmodal*--a user can continue to interact with both the document in the Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="4b7fd-123">Откройте диалоговое окно с главной страницы</span><span class="sxs-lookup"><span data-stu-id="4b7fd-123">Open a dialog box from a host page</span></span>

<span data-ttu-id="4b7fd-124">Office JavaScript API включает в себя [Диалоговый](/javascript/api/office/office.dialog) объекта и две функции в [Office.context.ui namespace](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-124">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="4b7fd-125">Чтобы открыть диалоговое окно, ваш код, обычно страница в панели задач, вызывает метод [displayDialogAsync](/javascript/api/office/office.ui) и передает ему URL-адрес ресурса, который вам нужно открыть.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-125">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="4b7fd-126">Страница, на которой этот метод вызван, называется "главной страницей".</span><span class="sxs-lookup"><span data-stu-id="4b7fd-126">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="4b7fd-127">Например, если вы вызываете этот метод в скрипте для index.html на панели задач, то index.html - это главная страница диалогового окна, которое открывает метод.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-127">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="4b7fd-128">Ресурс, который открывается в диалоговом окне, обычно представляет собой страницу, но это может быть метод контроллера в приложении MVC, маршрут, метод веб-службы или любой другой ресурс.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-128">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="4b7fd-129">В этой статье "страница" или "веб-сайт" ссылается на ресурс в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-129">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="4b7fd-130">Ниже приведен простой пример кода.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-130">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="4b7fd-131">В случае URL-адреса используется протокол HTTP **S**,</span><span class="sxs-lookup"><span data-stu-id="4b7fd-131">The URL uses the HTTP **S** protocol.</span></span> <span data-ttu-id="4b7fd-132">Обязательный для всех страниц, загружаемых в диалоговом окне, а не только для первой страницы.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-132">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="4b7fd-133">Домен диалогового окна совпадает с доменом главной страницы, которая может быть страницей в панели задач или [файлом функции](../reference/manifest/functionfile.md) команды надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-133">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](../reference/manifest/functionfile.md) of an add-in command.</span></span> <span data-ttu-id="4b7fd-134">Страница, метод контроллера или другой ресурс, передаваемый в метод `displayDialogAsync`, должен быть в том же домене, что и страница ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-134">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4b7fd-135">Главная страница и ресурс, который открывается в диалоговом окне, должны иметь один и тот же полный домен.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-135">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="4b7fd-136">Если вы попробуете передать поддомен домена надстройки в `displayDialogAsync`, ничего не получится.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-136">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="4b7fd-137">Полные доменные имена, включая поддомены, должны совпадать.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-137">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="4b7fd-138">После загрузки первой страницы (или другого ресурса) пользователь может использовать ссылки или другой пользовательский интерфейс для перехода на любой веб-сайт (или другой ресурс), использующий HTTPS.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-138">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="4b7fd-139">Первая страница также может сразу перенаправлять пользователя на другой сайт.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-139">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="4b7fd-140">По умолчанию диалоговое окно занимает 80 % высоты и ширины экрана устройства, но вы можете установить другие соотношения путем передачи объекта конфигурации в метод, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-140">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="4b7fd-141">Подобная надстройка приведена в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-141">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span> <span data-ttu-id="4b7fd-142">Дополнительные примеры см. в `displayDialogAsync` [примерах.](#samples)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-142">For more samples that use `displayDialogAsync`, see [Samples](#samples).</span></span>

<span data-ttu-id="4b7fd-p113">Установите оба значения равными 100 %, чтобы надстройка открывалась во весь экран. (На самом деле, максимальное значение составляет 99,5 %, возможность перемещать окно и изменять его размер сохраняется.)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p113">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="4b7fd-p114">Из главного окна можно открыть только одно диалоговое окно. При попытке открыть еще одно диалоговое окно произойдет ошибка. Поэтому если пользователь, например, откроет диалоговое окно из области задач, он не сможет открыть второе диалоговое окно на другой странице в области задач. Но при открытии диалогового окна с помощью [команды надстройки](../design/add-in-commands.md) каждый раз открывается новый (невидимый) HTML-файл. При этом создается новое (невидимое) главное окно, которое может запускать собственное диалоговое окно. Дополнительные сведения см. в разделе [Ошибки метода displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p114">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="4b7fd-151">Использование параметра производительности в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="4b7fd-151">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="4b7fd-152">`displayInIframe` — дополнительное свойство в объекте конфигурации, которое можно передать `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-152">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="4b7fd-153">Когда этому свойству присвоено значение `true`, а надстройка запущена для документа в Office в Интернете, диалоговое окно будет открываться быстрее, потому что будет выступать как плавающий фрейм iframe.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-153">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="4b7fd-154">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-154">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="4b7fd-155">Значение по умолчанию: `false`. Его использование равнозначно пропуску всего свойства.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-155">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="4b7fd-156">Если надстройка не работает в Office в Интернете, `displayInIframe` игнорируется.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-156">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="4b7fd-157">Вам **не** следует `displayInIframe: true`использовать, если диалоговое окно будет выполнять перенаправление на страницу, которую невозможно открыть в элементе iframe.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-157">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="4b7fd-158">Например, страницы для входов многих популярных веб-служб, таких как учетная запись Google и Майкрософт, невозможно открыть в iframe.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-158">For example, the sign in pages of many popular web services, such as Google and Microsoft account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="4b7fd-159">Отправка сведений из диалогового окна главной странице</span><span class="sxs-lookup"><span data-stu-id="4b7fd-159">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="4b7fd-160">Диалоговое окно может взаимодействовать с главной страницей в области задач, если:</span><span class="sxs-lookup"><span data-stu-id="4b7fd-160">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="4b7fd-161">Текущая страница в диалоговом окне не находится в том же домене, что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-161">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="4b7fd-162">Библиотека API JavaScript для Office загружается на странице.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-162">The Office JavaScript API library is loaded in the page.</span></span> <span data-ttu-id="4b7fd-163">(Как и любая страница, использующая библиотеку API JavaScript для Office, сценарий для страницы должен назначить свойству метод, хотя это может быть `Office.initialize` пустой метод.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-163">(Like any page that uses the Office JavaScript API library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method.</span></span> <span data-ttu-id="4b7fd-164">Дополнительные сведения см. в [инициализации надстройки Office.)](initialize-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-164">For details, see [Initialize your Office Add-in](initialize-add-in.md).)</span></span>

<span data-ttu-id="4b7fd-165">Код в диалоговом окне использует функцию [messageParent](/javascript/api/office/office.ui#messageparent-message-), чтобы отправить на главную страницу логическое значение или строковое сообщение.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-165">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="4b7fd-166">Строка может быть словом, предложением, большим двоичным объектом XML, преобразованными данными JSON или любыми другими объектами, которые можно сериализовать в строку.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-166">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="4b7fd-167">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-167">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - <span data-ttu-id="4b7fd-168">Функцию `messageParent` можно вызывать только на странице, которая относится к тому же домену (включая протокол и порт), что и главная страница.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-168">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>
> - <span data-ttu-id="4b7fd-169">Эта функция является одним из двух API JS Office, которые можно назвать `messageParent` в диалоговом  окне.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-169">The `messageParent` function is one of *only* two Office JS APIs that can be called in the dialog box.</span></span>
> - <span data-ttu-id="4b7fd-170">Другой API JS, который можно назвать в диалоговом окне: `Office.context.requirements.isSetSupported` .</span><span class="sxs-lookup"><span data-stu-id="4b7fd-170">The other JS API that can be called in the dialog box is `Office.context.requirements.isSetSupported`.</span></span> <span data-ttu-id="4b7fd-171">Сведения об этом см. в [подразделе "Указание приложений Office и требований к API".](specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-171">For information about it, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).</span></span> <span data-ttu-id="4b7fd-172">Однако в диалоговом окне этот API не поддерживается в Outlook 2016 одноразовая покупка (то есть версия MSI).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-172">However, in the dialog box, this API isn't supported in Outlook 2016 one-time purchase (that is, the MSI version).</span></span>

<span data-ttu-id="4b7fd-173">В следующем примере `googleProfile` — это строковое представление профиля Google пользователя.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-173">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="4b7fd-p121">Чтобы главная страница получила сообщение, ее необходимо настроить. Для этого добавьте параметр обратного вызова в исходный вызов `displayDialogAsync`. Обратный вызов назначает событию `DialogMessageReceived` обработчик. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p121">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="4b7fd-178">Office передает объект [AsyncResult](/javascript/api/office/office.asyncresult) в функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-178">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="4b7fd-179">Он представляет результат попытки открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-179">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="4b7fd-180">Он не представляет результат событий в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-180">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="4b7fd-181">Подробнее об этом различии см. в [Обработка ошибок и событий](dialog-handle-errors-events.md). </span><span class="sxs-lookup"><span data-stu-id="4b7fd-181">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="4b7fd-182">Для свойства `value` объекта `asyncResult` задан объект [Dialog](/javascript/api/office/office.dialog), который существует на главной странице, а не в контексте выполнения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-182">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="4b7fd-p123">`processMessage` — это функция, которая обрабатывает событие. Вы можете присвоить ей любое имя.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p123">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="4b7fd-185">Переменная `dialog` объявляется в более широком контексте, чем обратный вызов, так как на нее также ссылается `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-185">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="4b7fd-186">Ниже приведен простой пример обработчика для события `DialogMessageReceived`.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-186">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="4b7fd-187">Office передает объект `arg` в обработчик.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-187">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="4b7fd-188">Его `message` свойством является логическое значение или строка, отправляемая при вызове `messageParent` в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-188">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="4b7fd-189">В этом примере это строковые представления профиля пользователя из службы, такой как учетная запись Майкрософт или Google, поэтому она десериализирована обратно в объект с `JSON.parse` .</span><span class="sxs-lookup"><span data-stu-id="4b7fd-189">In this example, it is a stringified representation of a user's profile from a service such as Microsoft account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="4b7fd-p125">Функция `showUserName` не показана. Она может отображать персонализированное приветствие в области задач.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p125">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="4b7fd-192">Когда взаимодействие пользователя с диалоговым окном закончится, обработчик сообщений должен закрыть диалоговое окно, как показано в этом примере.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-192">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="4b7fd-193">Объект `dialog` должен быть таким же, как объект, который возвращается при вызове `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-193">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="4b7fd-194">Вызов метода `dialog.close` дает указание Office немедленно закрыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-194">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="4b7fd-195">Пример надстройки, в которой используются эти методы, см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-195">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="4b7fd-p126">Если надстройка должна открыть другую страницу области задач после получения сообщения, можно использовать метод `window.location.replace` (или `window.location.href`) в последней строке обработчика. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p126">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="4b7fd-198">Пример подобной надстройки см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-198">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="4b7fd-199">Условные сообщения</span><span class="sxs-lookup"><span data-stu-id="4b7fd-199">Conditional messaging</span></span>

<span data-ttu-id="4b7fd-200">Так как из диалогового окна можно отправить несколько вызовов `messageParent`, но на главной странице есть только один обработчик для события `DialogMessageReceived`, обработчику необходимо использовать условную логику, чтобы различать сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-200">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="4b7fd-201">Например, если в диалоговом окне пользователю будет предложено войти к поставщику удостоверений, например учетной записи Майкрософт или Google, профиль пользователя отправляется в качестве сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-201">For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft account or Google, it sends the user's profile as a message.</span></span> <span data-ttu-id="4b7fd-202">Если выполнить аутентификацию не удается, диалоговое окно отправляет сведения об ошибке на главную страницу, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-202">If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="4b7fd-203">Переменная `loginSuccess` будет инициализирована после считывания отклика HTTP от поставщика удостоверений.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-203">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="4b7fd-p128">Реализация функций `getProfile` и `getError` не показана. Они получают данные из параметра запроса или ответа HTTP.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p128">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="4b7fd-p129">В зависимости от того, удалось ли выполнить вход, отправляются анонимные объекты различных типов. Оба содержат свойство `messageType`, но один содержит свойство `profile`, а другой — свойство `error`.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p129">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="4b7fd-p130">Код обработчика на главной странице использует значение свойства `messageType` для разветвления, как показано в приведенном ниже примере. Обратите внимание на то, что здесь используется та же функция `showUserName`, что и в примере выше, а функция `showNotification` отображает сообщение об ошибке в элементе пользовательского интерфейса на главной странице.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p130">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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
> <span data-ttu-id="4b7fd-210">Реализация функции `showNotification` не показана в примере кода, представленном в этой статье.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-210">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="4b7fd-211">Пример возможного способа реализации этой функции в своей надстройке см. в статье [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-211">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="4b7fd-212">Передача данных диалоговому окну</span><span class="sxs-lookup"><span data-stu-id="4b7fd-212">Pass information to the dialog box</span></span>

<span data-ttu-id="4b7fd-213">Надстройка может отправлять сообщения [](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) с хост-страницы в диалоговое окно с помощью [Dialog.messageChild.](/javascript/api/office/office.dialog#messagechild-message-)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-213">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using [Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-).</span></span>

### <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="4b7fd-214">Использование `messageChild()` со страницы ведущего сайта</span><span class="sxs-lookup"><span data-stu-id="4b7fd-214">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="4b7fd-215">При вызове dialog API Office для открытия диалоговых окно возвращается объект [Dialog.](/javascript/api/office/office.dialog)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-215">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="4b7fd-216">Он должен быть назначен переменной с большей областью действия, чем метод [displayDialogAsync,](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) так как на объект будут ссылаться другие методы.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-216">It should be assigned to a variable that has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="4b7fd-217">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-217">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="4b7fd-218">Этот `Dialog` объект имеет метод [messageChild,](/javascript/api/office/office.dialog#messagechild-message-) который отправляет любую строку, включая строковые данные, в диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-218">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, including stringified data, to the dialog box.</span></span> <span data-ttu-id="4b7fd-219">Это вызывает `DialogParentMessageReceived` событие в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-219">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="4b7fd-220">Код должен обрабатывать это событие, как показано в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-220">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="4b7fd-221">Рассмотрим сценарий, в котором пользовательский интерфейс диалога связан с текущим активным на данный момент и положением этого таблицы относительно других таблиц.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-221">Consider a scenario in which the UI of the dialog is related to the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="4b7fd-222">В следующем примере свойства листа Excel отправляются `sheetPropertiesChanged` в диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-222">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="4b7fd-223">В этом случае текущий лист называется "Мой лист" и является вторым листом в книге.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-223">In this case, the current worksheet is named "My Sheet" and it's the second sheet in the workbook.</span></span> <span data-ttu-id="4b7fd-224">Данные инкапсулированы в объект и строку, чтобы их можно было передать `messageChild` в .</span><span class="sxs-lookup"><span data-stu-id="4b7fd-224">The data is encapsulated in an object and stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="4b7fd-225">Обработка DialogParentMessageReceived в диалоговом окне</span><span class="sxs-lookup"><span data-stu-id="4b7fd-225">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="4b7fd-226">В JavaScript диалогового окна зарегистрируйте обработатель события с помощью метода `DialogParentMessageReceived` [UI.addHandlerAsync.](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-226">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="4b7fd-227">Обычно это делается в [методах Office.onReady Office.initialize,](initialize-add-in.md)как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-227">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md), as shown in the following.</span></span> <span data-ttu-id="4b7fd-228">(Более надежный пример приведен ниже.)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-228">(A more robust example is below.)</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="4b7fd-229">Затем определите `onMessageFromParent` обработок.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-229">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="4b7fd-230">Следующий код продолжает пример из предыдущего раздела.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-230">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="4b7fd-231">Обратите внимание, что Office передает аргумент обработработеку и свойство объекта аргумента содержит строку `message` с хост-страницы.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-231">Note that Office passes an argument to the handler and that the `message` property of the argument object contains the string from the host page.</span></span> <span data-ttu-id="4b7fd-232">В этом примере сообщение перенаправляется в объект, и jQuery используется для того, чтобы установить верхний заголовок диалоговых окно в соответствие с новым именем таблицы.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-232">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="4b7fd-233">Лучше всего проверить правильность регистрации обработера.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-233">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="4b7fd-234">Это можно сделать, передав методу `addHandlerAsync` вызов.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-234">You can do this by passing a callback to the `addHandlerAsync` method.</span></span> <span data-ttu-id="4b7fd-235">Этот запуск выполняется после завершения попытки регистрации обработера.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-235">This runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="4b7fd-236">Используйте обработок для регистрации в журнале или для показа ошибки, если обработка не была успешно зарегистрирована.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-236">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="4b7fd-237">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-237">The following is an example.</span></span> <span data-ttu-id="4b7fd-238">Обратите внимание, что это функция, которая не определена здесь, которая регистрет `reportError` или отображает ошибку.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-238">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a><span data-ttu-id="4b7fd-239">Условный обмен сообщениями с родительской страницы на диалоговое окно</span><span class="sxs-lookup"><span data-stu-id="4b7fd-239">Conditional messaging from parent page to dialog box</span></span>

<span data-ttu-id="4b7fd-240">Так как на хост-странице можно сделать несколько вызовов, но в диалоговом окне для события есть только один обработок, обработчивая должна использовать условную логику, чтобы различать различные `messageChild` `DialogParentMessageReceived` сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-240">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="4b7fd-241">Это можно сделать точно так, как вы бы структурировали условные сообщения, когда диалоговое окно отправляет сообщение на страницу ведущего приложения, как описано в условных [сообщениях.](#conditional-messaging)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-241">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](#conditional-messaging).</span></span>

> [!NOTE]
> <span data-ttu-id="4b7fd-242">В некоторых ситуациях API, который входит в набор требований `messageChild` [DialogApi 1.2,](../reference/requirement-sets/dialog-api-requirement-sets.md)может не поддерживаться.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-242">In some situations, the `messageChild` API, which is a part of the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md),  may not be supported.</span></span> <span data-ttu-id="4b7fd-243">Некоторые альтернативные способы передачи сообщений из родительского диалогового окна описаны в альтернативных способах передачи сообщений в диалоговое окно с его [хост-страницы.](parent-to-dialog.md)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-243">Some alternative ways for parent-to-dialog-box messaging are described in [Alternative ways of passing messages to a dialog box from its host page](parent-to-dialog.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4b7fd-244">Набор [требований DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md) не может быть указан в разделе `<Requirements>` манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-244">The [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md) cannot be specified in the `<Requirements>` section of an add-in manifest.</span></span> <span data-ttu-id="4b7fd-245">Вам придется проверить поддержку DialogApi 1.2 во время работы с помощью [метода isSetSupported.](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-245">You will have to check for support for DialogApi 1.2 at runtime using the [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) method.</span></span> <span data-ttu-id="4b7fd-246">Поддержка требований манифеста находится на стадии разработки.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-246">Support for manifest requirements is under development.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="4b7fd-247">Закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="4b7fd-247">Closing the dialog box</span></span>

<span data-ttu-id="4b7fd-p141">Вы можете добавить в диалоговое окно кнопку, которая будет его закрывать. Для этого обработчик событий кнопки должен использовать `messageParent`, чтобы сообщить главной странице, что кнопка нажата. Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-p141">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="4b7fd-251">Обработчик главной страницы для `DialogMessageReceived` вызовет `dialog.close`, как показано в этом примере.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-251">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="4b7fd-252">(См. предыдущие примеры, в которых показано, как `dialog` инициализируется объект).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-252">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="4b7fd-253">Даже если у вас нет собственного пользовательского интерфейса для закрытия диалогового окна, пользователь может закрыть диалоговое окно, выбрав **X** в правом верхнем углу.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-253">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="4b7fd-254">Это действие запускает событие `DialogEventReceived`.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-254">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="4b7fd-255">Чтобы главная область могла реагировать на это событие, для нее должен быть объявлен обработчик этого события.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-255">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="4b7fd-256">Дополнительные сведения см. в разделе [Ошибок и события в диалоговом окне](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-256">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="4b7fd-257">Продвинутые темы и специальные сценарии</span><span class="sxs-lookup"><span data-stu-id="4b7fd-257">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="4b7fd-258">Используйте Dialog API, чтобы показать видео</span><span class="sxs-lookup"><span data-stu-id="4b7fd-258">Use the Dialog API to show a video</span></span>

<span data-ttu-id="4b7fd-259">См. статью [Использование диалогового окна «Office» для отображения видео](dialog-video.md).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-259">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="4b7fd-260">Использование Dialog API в потоке аутентификации</span><span class="sxs-lookup"><span data-stu-id="4b7fd-260">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="4b7fd-261">См. статью[ Проверка подлинности с помощью Office Dialog API ](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-261">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="4b7fd-262">Использование Office dialog API с одностраничными приложениями и клиентской маршрутизацией</span><span class="sxs-lookup"><span data-stu-id="4b7fd-262">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="4b7fd-263">При использовании Office dialog API, SPA и маршрутизация на стороне клиента должны обрабатываться с осторожностью</span><span class="sxs-lookup"><span data-stu-id="4b7fd-263">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="4b7fd-264">См. статью[Рекомендации по использованию Office dialog API в SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-264">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="4b7fd-265">Обработка ошибок и событий</span><span class="sxs-lookup"><span data-stu-id="4b7fd-265">Error and event handling</span></span>

<span data-ttu-id="4b7fd-266">См. статью об ошибках и событиях [Обработка ошибок и событий в Office dialog box](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-266">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="4b7fd-267">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="4b7fd-267">Next steps</span></span>

<span data-ttu-id="4b7fd-268">Узнайте о том, как использовать Office dialog API, в [Рекомендации по использованию Office dialog API](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="4b7fd-268">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

## <a name="samples"></a><span data-ttu-id="4b7fd-269">Примеры</span><span class="sxs-lookup"><span data-stu-id="4b7fd-269">Samples</span></span>

<span data-ttu-id="4b7fd-270">Все следующие примеры используют `displayDialogAsync` .</span><span class="sxs-lookup"><span data-stu-id="4b7fd-270">All of the following samples use `displayDialogAsync`.</span></span> <span data-ttu-id="4b7fd-271">Некоторые из них используют серверы на основе NodeJS, а другие ASP.NET/IIS-based, но логика использования метода не зависит от того, как реализована надстройка на стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="4b7fd-271">Some have NodeJS-based servers and others have ASP.NET/IIS-based servers, but the logic of using the method is the same regardless of how the server-side of the add-in is implemented.</span></span>

<span data-ttu-id="4b7fd-272">**Основные принципы:**</span><span class="sxs-lookup"><span data-stu-id="4b7fd-272">**Basics:**</span></span>

- [<span data-ttu-id="4b7fd-273">Пример использования API диалоговых окон в надстройке Office</span><span class="sxs-lookup"><span data-stu-id="4b7fd-273">Office Add-in Dialog API Example</span></span>](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [<span data-ttu-id="4b7fd-274">Обучающий контент/ создание надстроек (несколько примеров)</span><span class="sxs-lookup"><span data-stu-id="4b7fd-274">Training Content / Building Add-ins (several samples)</span></span>](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

<span data-ttu-id="4b7fd-275">**Более сложные примеры:**</span><span class="sxs-lookup"><span data-stu-id="4b7fd-275">**More complex samples:**</span></span>

- [<span data-ttu-id="4b7fd-276">ASPNET надстройки Office для Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="4b7fd-276">Office Add-in Microsoft Graph ASPNET</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [<span data-ttu-id="4b7fd-277">Надстройка Office в Microsoft Graph React</span><span class="sxs-lookup"><span data-stu-id="4b7fd-277">Office Add-in Microsoft Graph React</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [<span data-ttu-id="4b7fd-278">Единый вход с использованием NodeJS для надстройки Office</span><span class="sxs-lookup"><span data-stu-id="4b7fd-278">Office Add-in NodeJS SSO</span></span>](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [<span data-ttu-id="4b7fd-279">Служба SSO для надстройки Office ASPNET</span><span class="sxs-lookup"><span data-stu-id="4b7fd-279">Office Add-in ASPNET SSO</span></span>](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [<span data-ttu-id="4b7fd-280">Пример монетизации SAAS для надстройки Office</span><span class="sxs-lookup"><span data-stu-id="4b7fd-280">Office Add-in SAAS Monetization Sample</span></span>](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [<span data-ttu-id="4b7fd-281">AsPNET надстройки Outlook для Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="4b7fd-281">Outlook Add-in Microsoft Graph ASPNET</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [<span data-ttu-id="4b7fd-282">SSO для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="4b7fd-282">Outlook Add-in SSO</span></span>](https://github.com/OfficeDev/Outlook-Add-in-SSO)
- [<span data-ttu-id="4b7fd-283">Просмотр маркеров надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="4b7fd-283">Outlook Add-in Token Viewer</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [<span data-ttu-id="4b7fd-284">Сообщение с действиями надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="4b7fd-284">Outlook Add-in Actionable Message</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [<span data-ttu-id="4b7fd-285">Общий доступ к надстройки Outlook в OneDrive</span><span class="sxs-lookup"><span data-stu-id="4b7fd-285">Outlook Add-in Sharing to OneDrive</span></span>](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [<span data-ttu-id="4b7fd-286">Вставка диаграммы ASPNET надстройки PowerPoint в Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="4b7fd-286">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [<span data-ttu-id="4b7fd-287">Сценарий общей времени работы Excel</span><span class="sxs-lookup"><span data-stu-id="4b7fd-287">Excel Shared Runtime Scenario</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/900b5769bca9bbcff79d6cd6106d9fcc55c70d5a/Samples/excel-shared-runtime-scenario)
- [<span data-ttu-id="4b7fd-288">Краткие книги надстройки Excel в ASPNET</span><span class="sxs-lookup"><span data-stu-id="4b7fd-288">Excel Add-in ASPNET QuickBooks</span></span>](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [<span data-ttu-id="4b7fd-289">Надстройка Word JS Redact</span><span class="sxs-lookup"><span data-stu-id="4b7fd-289">Word Add-in JS Redact</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [<span data-ttu-id="4b7fd-290">Надстройка Word JS SpecKit</span><span class="sxs-lookup"><span data-stu-id="4b7fd-290">Word Add-in JS SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [<span data-ttu-id="4b7fd-291">OAuth клиента надстройки Word в AngularJS</span><span class="sxs-lookup"><span data-stu-id="4b7fd-291">Word Add-in AngularJS Client OAuth</span></span>](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [<span data-ttu-id="4b7fd-292">Надстройка Office Auth0</span><span class="sxs-lookup"><span data-stu-id="4b7fd-292">Office Add-in Auth0</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="4b7fd-293">Надстройки Office OAuth.io</span><span class="sxs-lookup"><span data-stu-id="4b7fd-293">Office Add-in OAuth.io</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [<span data-ttu-id="4b7fd-294">Код шаблонов дизайна для надстройки Office</span><span class="sxs-lookup"><span data-stu-id="4b7fd-294">Office Add-in UX Design Patterns Code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

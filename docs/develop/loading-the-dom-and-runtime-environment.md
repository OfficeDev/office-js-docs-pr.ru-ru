---
title: Загрузка модели DOM и среды выполнения
description: Загрузка среды выполнения надстроек DOM и надстроек Office
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 02f950ca23d52b333f704c7d8aed431cb426a6f0
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293277"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="7010a-103">Загрузка модели DOM и среды выполнения</span><span class="sxs-lookup"><span data-stu-id="7010a-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="7010a-104">Перед запуском собственной логики надстройка должна проверить, что загружены модель DOM и среда выполнения Надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="7010a-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="7010a-105">Запуск контентной надстройки или надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="7010a-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="7010a-106">На рисунке ниже приведен поток событий, происходящих при запуске контентной надстройки или надстройки области задач в Excel, PowerPoint, Project или Word.</span><span class="sxs-lookup"><span data-stu-id="7010a-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![Поток событий при запуске контентной надстройки или надстройки области задач](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="7010a-108">При запуске контентной надстройки или надстройки области задач возникают указанные ниже события.</span><span class="sxs-lookup"><span data-stu-id="7010a-108">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="7010a-109">Пользователь открывает документ, который уже содержит надстройку, или вставляет надстройку в документ.</span><span class="sxs-lookup"><span data-stu-id="7010a-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="7010a-110">Клиентское приложение Office считывает XML-манифест надстройки из AppSource, каталога приложений в SharePoint или каталога общих папок, от которого он создан.</span><span class="sxs-lookup"><span data-stu-id="7010a-110">The Office client application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="7010a-111">Клиентское приложение Office открывает HTML-страницу надстройки в элементе управления браузера.</span><span class="sxs-lookup"><span data-stu-id="7010a-111">The Office client application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="7010a-p101">Следующие два действия, 4 и 5, выполняются одновременно и параллельно. Поэтому код надстройки перед обработкой должен убедиться, что и модель DOM, и среда выполнения надстройки полностью загрузились.</span><span class="sxs-lookup"><span data-stu-id="7010a-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="7010a-114">Элемент управления браузером загружает DOM и основной текст HTML и вызывает обработчик события для `window.onload` события.</span><span class="sxs-lookup"><span data-stu-id="7010a-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="7010a-115">Клиентское приложение Office загружает среду выполнения, которая загружает и кэширует файлы библиотеки API JavaScript для Office из сервера сети распространения содержимого (CDN), а затем вызывает обработчик событий надстройки для события [Initialize](/javascript/api/office#office-initialize-reason-) объекта [Office](/javascript/api/office) , если ему назначен обработчик.</span><span class="sxs-lookup"><span data-stu-id="7010a-115">The Office client application loads the runtime environment, which downloads and caches the Office JavaScript API library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="7010a-116">В это время также проверяется, выполнялась ли передача (или связывание) любых обратных вызовов (или связанных функций `then()`) обработчику `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="7010a-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="7010a-117">Для получения дополнительных сведений о различии между `Office.initialize` и `Office.onReady` см [Initialize your add-in](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="7010a-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="7010a-118">После завершения загрузки DOM и основного текста HTML и инициализации надстройки запускается основная функция надстройки.</span><span class="sxs-lookup"><span data-stu-id="7010a-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="7010a-119">Запуск надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="7010a-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="7010a-120">На рисунке ниже приведен поток событий при запуске надстройки Outlook на настольном компьютере, планшетном ПК или смартфоне.</span><span class="sxs-lookup"><span data-stu-id="7010a-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Поток событий при запуске надстройки Outlook](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="7010a-122">При запуске надстройки Outlook происходят указанные ниже события.</span><span class="sxs-lookup"><span data-stu-id="7010a-122">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="7010a-123">При запуске Outlook считывает XML-манифесты надстроек Outlook, установленных для учетной записи пользователя.</span><span class="sxs-lookup"><span data-stu-id="7010a-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="7010a-124">Пользователь выбирает элемент в Outlook.</span><span class="sxs-lookup"><span data-stu-id="7010a-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="7010a-125">Если выбранный элемент удовлетворяет условиям активации надстройки Outlook, то Outlook активирует надстройку и делает соответствующую кнопку видимой в пользовательском интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="7010a-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="7010a-p103">Если пользователь нажимает кнопку для запуска надстройки Outlook, то ведущее приложение открывает HTML-страницу в элементе управления браузером. Следующие два шага, шаг 5 и шаг 6, выполняются одновременно.</span><span class="sxs-lookup"><span data-stu-id="7010a-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="7010a-128">Элемент управления браузером загружает DOM и основной текст HTML и вызывает обработчик события для `onload` события.</span><span class="sxs-lookup"><span data-stu-id="7010a-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="7010a-129">Outlook загружает среду выполнения, которая загружает и кэширует API JavaScript для файлов библиотеки JavaScript с сервера сети доставки содержимого, а затем вызывает обработчик события [инициализации](/javascript/api/office#office-initialize-reason-) объекта [Office](/javascript/api/office) надстройки, если ему назначен обработчик.</span><span class="sxs-lookup"><span data-stu-id="7010a-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="7010a-130">В это время также проверяется, выполнялась ли передача (или связывание) любых обратных вызовов (или связанных функций `then()`) обработчику `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="7010a-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="7010a-131">Для получения дополнительных сведений о различии между `Office.initialize` и `Office.onReady` см [Initialize your add-in](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="7010a-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="7010a-132">После завершения загрузки DOM и основного текста HTML и инициализации надстройки запускается основная функция надстройки.</span><span class="sxs-lookup"><span data-stu-id="7010a-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="checking-the-load-status"></a><span data-ttu-id="7010a-133">Проверка состояния загрузки</span><span class="sxs-lookup"><span data-stu-id="7010a-133">Checking the load status</span></span>

<span data-ttu-id="7010a-134">Одним из способов проверки завершения загрузки DOM и среды выполнения надстроек является использование функции [.ready()](https://api.jquery.com/ready/) jQuery — `$(document).ready()`.</span><span class="sxs-lookup"><span data-stu-id="7010a-134">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="7010a-135">Например, следующий `onReady` обработчик событий гарантирует, что модель DOM сначала загружается, прежде чем будет запущен код, предназначенный для инициализации надстройки.</span><span class="sxs-lookup"><span data-stu-id="7010a-135">For example, the following `onReady` event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="7010a-136">Затем `onReady` обработчик продолжает использовать свойство [Mailbox. Item](/javascript/api/outlook/office.mailbox#item) для получения выбранного в данный момент элемента в Outlook и вызывает основную функцию надстройки `initDialer` .</span><span class="sxs-lookup"><span data-stu-id="7010a-136">Subsequently, the `onReady` handler proceeds to use the [mailbox.item](/javascript/api/outlook/office.mailbox#item) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

```js
Office.onReady()
    .then(
        // Checks for the DOM to load.
        $(document).ready(function () {
            // After the DOM is loaded, add-in-specific code can run.
            var mailbox = Office.context.mailbox;
            _Item = mailbox.item;
            initDialer();
        });
);
```

<span data-ttu-id="7010a-137">Кроме того, вы можете использовать один и тот же код в `initialize` обработчике событий, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="7010a-137">Alternatively, you can use the same code in an `initialize` event handler as shown in the following example.</span></span>

```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

<span data-ttu-id="7010a-138">Этот же метод можно использовать в `onReady` `initialize` обработчиках или в обработчиках любой надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="7010a-138">This same technique can be used in the `onReady` or `initialize` handlers of any Office Add-in.</span></span>

<span data-ttu-id="7010a-139">В примере надстройки Outlook "Телефон" показан несколько другой подход, использующий только JavaScript для проверки тех же условий.</span><span class="sxs-lookup"><span data-stu-id="7010a-139">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7010a-140">Даже если у надстройки нет задач инициализации, необходимо включить по крайней мере вызов `Office.onReady` или назначить минимальную `Office.initialize` функцию обработчика событий, как показано в следующих примерах.</span><span class="sxs-lookup"><span data-stu-id="7010a-140">Even if your add-in has no initialization tasks to perform, you must include at least a call of `Office.onReady` or assign minimal `Office.initialize` event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="7010a-141">Если не вызвать `Office.onReady` или назначить `Office.initialize` обработчик события, надстройка может вызвать ошибку при запуске.</span><span class="sxs-lookup"><span data-stu-id="7010a-141">If you do not call `Office.onReady` or assign an `Office.initialize` event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="7010a-142">Кроме того, если пользователь попробует использовать надстройку с веб-клиентом Office, например Excel, PowerPoint или Outlook, произойдет сбой.</span><span class="sxs-lookup"><span data-stu-id="7010a-142">Also, if a user attempts to use your add-in with an Office web client, such as Excel, PowerPoint, or Outlook, it will fail to run.</span></span>
>
> <span data-ttu-id="7010a-143">Если надстройка содержит несколько страниц, при загрузке новой страницы, которая должна вызываться `Office.onReady` или назначить `Office.initialize` обработчик событий.</span><span class="sxs-lookup"><span data-stu-id="7010a-143">If your add-in includes more than one page, whenever it loads a new page that page must either call `Office.onReady` or assign an `Office.initialize` event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="7010a-144">См. также</span><span class="sxs-lookup"><span data-stu-id="7010a-144">See also</span></span>

- [<span data-ttu-id="7010a-145">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="7010a-145">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="7010a-146">Инициализация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="7010a-146">Initialize your Office Add-in</span></span>](initialize-add-in.md)

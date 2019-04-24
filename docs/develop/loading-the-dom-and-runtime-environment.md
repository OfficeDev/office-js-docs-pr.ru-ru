---
title: Загрузка модели DOM и среды выполнения
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: b1f63d9fe012ed8c8a5cf4a0f7de862ddabcd4d3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449848"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="10228-102">Загрузка модели DOM и среды выполнения</span><span class="sxs-lookup"><span data-stu-id="10228-102">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="10228-103">Перед запуском собственной логики надстройка должна проверить, что загружены модель DOM и среда выполнения Надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="10228-103">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span> 

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="10228-104">Запуск контентной надстройки или надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="10228-104">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="10228-105">На рисунке ниже приведен поток событий, происходящих при запуске контентной надстройки или надстройки области задач в Excel, PowerPoint, Project, Word или Access.</span><span class="sxs-lookup"><span data-stu-id="10228-105">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, Word, or Access.</span></span>

![Поток событий при запуске контентной надстройки или надстройки области задач](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="10228-107">При запуске контентной надстройки или надстройки области задач возникают указанные ниже события.</span><span class="sxs-lookup"><span data-stu-id="10228-107">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="10228-108">Пользователь открывает документ, который уже содержит надстройку, или вставляет надстройку в документ.</span><span class="sxs-lookup"><span data-stu-id="10228-108">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="10228-109">Ведущее приложение Office читает XML-манифест надстройки из AppSource, каталога надстроек в SharePoint или каталога общей папки, в зависимости от того, откуда берется надстройка.</span><span class="sxs-lookup"><span data-stu-id="10228-109">The Office host application reads the add-in's XML manifest from AppSource, an add-in catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="10228-110">Ведущее приложение Office открывает HTML-страницу надстройки в элементе управления браузера.</span><span class="sxs-lookup"><span data-stu-id="10228-110">The Office host application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="10228-p101">Следующие два действия, 4 и 5, выполняются одновременно и параллельно. Поэтому код надстройки перед обработкой должен убедиться, что и модель DOM, и среда выполнения надстройки полностью загрузились.</span><span class="sxs-lookup"><span data-stu-id="10228-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="10228-113">Элемент управления браузера загружает модель DOM и основной текст HTML, а также вызывает обработчик для события **window.onload**.</span><span class="sxs-lookup"><span data-stu-id="10228-113">The browser control loads the DOM and HTML body, and calls the event handler for the  **window.onload** event.</span></span>

5. <span data-ttu-id="10228-114">Ведущее приложение Office загружает среду выполнения, которая загружает и кэширует API JavaScript для файлов библиотеки JavaScript с сервера сети доставки содержимого, а затем вызывает обработчик события [инициализации](/javascript/api/office#office-initialize) объекта [Office](/javascript/api/office), если ему назначен обработчик.</span><span class="sxs-lookup"><span data-stu-id="10228-114">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="10228-115">В это время также проверяется, выполнялась ли передача (или связывание) любых обратных вызовов (или связанных функций `then()`) обработчику `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="10228-115">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="10228-116">Дополнительные сведения о различии между `Office.initialize` и `Office.onReady` см. в статье [Инициализация надстройки](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="10228-116">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initializing your add-in](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).</span></span>

6. <span data-ttu-id="10228-117">После завершения загрузки DOM и основного текста HTML и инициализации надстройки запускается основная функция надстройки.</span><span class="sxs-lookup"><span data-stu-id="10228-117">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="10228-118">Запуск надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="10228-118">Startup of an Outlook add-in</span></span>

<span data-ttu-id="10228-119">На рисунке ниже приведен поток событий при запуске надстройки Outlook на настольном компьютере, планшетном ПК или смартфоне.</span><span class="sxs-lookup"><span data-stu-id="10228-119">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Поток событий при запуске надстройки Outlook](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="10228-121">При запуске надстройки Outlook происходят указанные ниже события.</span><span class="sxs-lookup"><span data-stu-id="10228-121">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="10228-122">При запуске Outlook считывает XML-манифесты надстроек Outlook, установленных для учетной записи пользователя.</span><span class="sxs-lookup"><span data-stu-id="10228-122">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="10228-123">Пользователь выбирает элемент в Outlook.</span><span class="sxs-lookup"><span data-stu-id="10228-123">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="10228-124">Если выбранный элемент удовлетворяет условиям активации надстройки Outlook, то Outlook активирует надстройку и делает соответствующую кнопку видимой в пользовательском интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="10228-124">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="10228-p103">Если пользователь нажимает кнопку для запуска надстройки Outlook, то ведущее приложение открывает HTML-страницу в элементе управления браузером. Следующие два шага, шаг 5 и шаг 6, выполняются одновременно.</span><span class="sxs-lookup"><span data-stu-id="10228-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="10228-127">Элемент управления браузером загружает DOM и основной текст HTML и вызывает обработчик события **onload**.</span><span class="sxs-lookup"><span data-stu-id="10228-127">The browser control loads the DOM and HTML body, and calls the event handler for the  **onload** event.</span></span>

6. <span data-ttu-id="10228-128">Outlook загружает среду выполнения, которая загружает и кэширует API JavaScript для файлов библиотеки JavaScript с сервера сети доставки содержимого, а затем вызывает обработчик события [инициализации](/javascript/api/office#office-initialize) объекта [Office](/javascript/api/office) надстройки, если ему назначен обработчик.</span><span class="sxs-lookup"><span data-stu-id="10228-128">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="10228-129">В это время также проверяется, выполнялась ли передача (или связывание) любых обратных вызовов (или связанных функций `then()`) обработчику `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="10228-129">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="10228-130">Дополнительные сведения о различии между `Office.initialize` и `Office.onReady` см. в статье [Инициализация надстройки](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="10228-130">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initializing your add-in](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).</span></span>

7. <span data-ttu-id="10228-131">После завершения загрузки DOM и основного текста HTML и инициализации надстройки запускается основная функция надстройки.</span><span class="sxs-lookup"><span data-stu-id="10228-131">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="checking-the-load-status"></a><span data-ttu-id="10228-132">Проверка состояния загрузки</span><span class="sxs-lookup"><span data-stu-id="10228-132">Checking the load status</span></span>

<span data-ttu-id="10228-133">Одним из способов проверки завершения загрузки DOM и среды выполнения надстроек является использование функции [.ready()](https://api.jquery.com/ready/) jQuery — `$(document).ready()`.</span><span class="sxs-lookup"><span data-stu-id="10228-133">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="10228-134">Например, указанный ниже обработчик событий **onReady** убеждается в полной загрузке DOM, прежде чем выполняется код, относящийся к инициализации надстройки.</span><span class="sxs-lookup"><span data-stu-id="10228-134">For example, the following **onReady** event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="10228-135">После этого обработчик **onReady** переходит на использование свойства [mailbox.item](/javascript/api/outlook/office.mailbox) для получения выбранного в настоящий момент элемента Outlook и вызывает основную функцию надстройки `initDialer`.</span><span class="sxs-lookup"><span data-stu-id="10228-135">Subsequently, the **onReady** handler proceeds to use the [mailbox.item](/javascript/api/outlook/office.mailbox) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

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

<span data-ttu-id="10228-136">Кроме того, можно использовать такой же код в обработчике событий **инициализации**, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="10228-136">Alternatively, you can use the same code in an  **initialize** event handler as shown in the following example.</span></span>

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

<span data-ttu-id="10228-137">Этот же способ можно использовать в обработчиках **инициализации** или **onReady** любой надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="10228-137">This same technique can be used in the **onReady** or **initialize** handlers of any Office Add-in.</span></span>

<span data-ttu-id="10228-138">В примере надстройки Outlook "Телефон" показан несколько другой подход, использующий только JavaScript для проверки тех же условий.</span><span class="sxs-lookup"><span data-stu-id="10228-138">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="10228-139">Даже если у надстройки нет задач инициализации, необходимо включить по крайней мере вызов обработчика **Office.onReady** или назначить минимальную функцию обработчика событий **Office.initialize**, как показано в примере ниже.</span><span class="sxs-lookup"><span data-stu-id="10228-139">Even if your add-in has no initialization tasks to perform, you must include at least a call of **Office.onReady** or assign minimal **Office.initialize** event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="10228-140">Если не вызвать **Office.onReady** или не назначить обработчик событий **Office.initialize**, надстройка может выдать ошибку при запуске.</span><span class="sxs-lookup"><span data-stu-id="10228-140">If you do not call **Office.onReady** or assign an  **Office.initialize** event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="10228-141">Кроме того, если пользователь попробует использовать надстройку с веб-клиентом Office Online, например Excel Online, PowerPoint Online или Outlook Web App, произойдет сбой.</span><span class="sxs-lookup"><span data-stu-id="10228-141">Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run.</span></span>
>
> <span data-ttu-id="10228-142">Если надстройка содержит несколько страниц, каждая загружаемая страница должна вызывать  **Office.onReady** или назначать обработчик событий **Office.initialize**.</span><span class="sxs-lookup"><span data-stu-id="10228-142">If your add-in includes more than one page, whenever it loads a new page that page must either call **Office.onReady** or assign an  **Office.initialize** event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="10228-143">См. также</span><span class="sxs-lookup"><span data-stu-id="10228-143">See also</span></span>

- [<span data-ttu-id="10228-144">Общие сведения об интерфейсе API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="10228-144">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)

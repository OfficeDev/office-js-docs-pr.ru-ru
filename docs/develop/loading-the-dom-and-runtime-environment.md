---
title: Загрузка модели DOM и среды выполнения
description: Загрузите среду запуска надстройки DOM и Office.
ms.date: 04/20/2021
localization_priority: Normal
ms.openlocfilehash: 5a215bf5a81dd291e72ed9e396c156d9ea7c6db0
ms.sourcegitcommit: 691fa338029c9cbd9a7194d163f390c3321a0cd8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/23/2021
ms.locfileid: "51959168"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="bb67c-103">Загрузка модели DOM и среды выполнения</span><span class="sxs-lookup"><span data-stu-id="bb67c-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="bb67c-104">Перед запуском собственной логики надстройка должна проверить, что загружены модель DOM и среда выполнения Надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="bb67c-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="bb67c-105">Запуск контентной надстройки или надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="bb67c-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="bb67c-106">На рисунке ниже приведен поток событий, происходящих при запуске контентной надстройки или надстройки области задач в Excel, PowerPoint, Project или Word.</span><span class="sxs-lookup"><span data-stu-id="bb67c-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![Поток событий при запуске контентной надстройки или надстройки области задач](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="bb67c-108">При запуске контентной надстройки или надстройки области задач возникают указанные ниже события.</span><span class="sxs-lookup"><span data-stu-id="bb67c-108">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="bb67c-109">Пользователь открывает документ, который уже содержит надстройку, или вставляет надстройку в документ.</span><span class="sxs-lookup"><span data-stu-id="bb67c-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="bb67c-110">Клиентское приложение Office читает XML-манифест надстройки из AppSource, каталога приложений в SharePoint или из общего каталога папок, из него исходят.</span><span class="sxs-lookup"><span data-stu-id="bb67c-110">The Office client application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="bb67c-111">Клиентский приложение Office открывает HTML-страницу надстройки в элементе управления браузером.</span><span class="sxs-lookup"><span data-stu-id="bb67c-111">The Office client application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="bb67c-p101">Следующие два действия, 4 и 5, выполняются одновременно и параллельно. Поэтому код надстройки перед обработкой должен убедиться, что и модель DOM, и среда выполнения надстройки полностью загрузились.</span><span class="sxs-lookup"><span data-stu-id="bb67c-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="bb67c-114">Элемент управления браузером загружает тело DOM и HTML и вызывает обработчик событий для `window.onload` события.</span><span class="sxs-lookup"><span data-stu-id="bb67c-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="bb67c-115">Клиентское приложение Office загружает среду времени запуска, которая загружает и кэшизирует файлы библиотеки API API Office с сервера сети распространения контента (CDN), а затем вызывает обработчик событий надстройки для события инициализации объекта [Office,](/javascript/api/office) если ему назначен обработчик. [](/javascript/api/office#office-initialize-reason-)</span><span class="sxs-lookup"><span data-stu-id="bb67c-115">The Office client application loads the runtime environment, which downloads and caches the Office JavaScript API library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="bb67c-116">В это время также проверяется, выполнялась ли передача (или связывание) любых обратных вызовов (или связанных функций `then()`) обработчику `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="bb67c-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="bb67c-117">Дополнительные сведения о различиях между `Office.initialize` и `Office.onReady` , см. в [инициализации надстройки](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="bb67c-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="bb67c-118">После завершения загрузки DOM и основного текста HTML и инициализации надстройки запускается основная функция надстройки.</span><span class="sxs-lookup"><span data-stu-id="bb67c-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="bb67c-119">Запуск надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="bb67c-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="bb67c-120">На рисунке ниже приведен поток событий при запуске надстройки Outlook на настольном компьютере, планшетном ПК или смартфоне.</span><span class="sxs-lookup"><span data-stu-id="bb67c-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Поток событий при запуске надстройки Outlook](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="bb67c-122">При запуске надстройки Outlook происходят указанные ниже события.</span><span class="sxs-lookup"><span data-stu-id="bb67c-122">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="bb67c-123">При запуске Outlook считывает XML-манифесты надстроек Outlook, установленных для учетной записи пользователя.</span><span class="sxs-lookup"><span data-stu-id="bb67c-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="bb67c-124">Пользователь выбирает элемент в Outlook.</span><span class="sxs-lookup"><span data-stu-id="bb67c-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="bb67c-125">Если выбранный элемент удовлетворяет условиям активации надстройки Outlook, то Outlook активирует надстройку и делает соответствующую кнопку видимой в пользовательском интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="bb67c-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="bb67c-p103">Если пользователь нажимает кнопку для запуска надстройки Outlook, то ведущее приложение открывает HTML-страницу в элементе управления браузером. Следующие два шага, шаг 5 и шаг 6, выполняются одновременно.</span><span class="sxs-lookup"><span data-stu-id="bb67c-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="bb67c-128">Элемент управления браузером загружает тело DOM и HTML и вызывает обработчик событий для `onload` события.</span><span class="sxs-lookup"><span data-stu-id="bb67c-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="bb67c-129">Outlook загружает среду выполнения, которая загружает и кэширует API JavaScript для файлов библиотеки JavaScript с сервера сети доставки содержимого, а затем вызывает обработчик события [инициализации](/javascript/api/office#office-initialize-reason-) объекта [Office](/javascript/api/office) надстройки, если ему назначен обработчик.</span><span class="sxs-lookup"><span data-stu-id="bb67c-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="bb67c-130">В это время также проверяется, выполнялась ли передача (или связывание) любых обратных вызовов (или связанных функций `then()`) обработчику `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="bb67c-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="bb67c-131">Дополнительные сведения о различиях между `Office.initialize` и `Office.onReady` , см. в [инициализации надстройки](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="bb67c-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="bb67c-132">После завершения загрузки DOM и основного текста HTML и инициализации надстройки запускается основная функция надстройки.</span><span class="sxs-lookup"><span data-stu-id="bb67c-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>

## <a name="see-also"></a><span data-ttu-id="bb67c-133">См. также</span><span class="sxs-lookup"><span data-stu-id="bb67c-133">See also</span></span>

- [<span data-ttu-id="bb67c-134">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="bb67c-134">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="bb67c-135">Инициализация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="bb67c-135">Initialize your Office Add-in</span></span>](initialize-add-in.md)

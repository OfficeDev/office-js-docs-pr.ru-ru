---
title: Загрузка модели DOM и среды выполнения
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 0de9acf842d28002ec134998c36224f98574cf8f
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457693"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="48bdf-102">Загрузка модели DOM и среды выполнения</span><span class="sxs-lookup"><span data-stu-id="48bdf-102">Loading the DOM and runtime environment</span></span>



<span data-ttu-id="48bdf-103">Перед запуском собственной логики надстройка должна проверить, что загружены модель DOM и среда выполнения Надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="48bdf-103">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span> 

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="48bdf-104">Запуск контентной надстройки или надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="48bdf-104">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="48bdf-105">На рисунке ниже приведен поток событий, происходящих при запуске контентной надстройки или надстройки области задач в Excel, PowerPoint, Project, Word или Access.</span><span class="sxs-lookup"><span data-stu-id="48bdf-105">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, Word, or Access.</span></span>

![Поток событий при запуске контентной надстройки или надстройки области задач](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="48bdf-107">При запуске контентной надстройки или надстройки области задач возникают указанные ниже события.</span><span class="sxs-lookup"><span data-stu-id="48bdf-107">The following events occur when a content or task pane add-in starts:</span></span> 



1. <span data-ttu-id="48bdf-108">Пользователь открывает документ, который уже содержит надстройку, или вставляет надстройку в документ.</span><span class="sxs-lookup"><span data-stu-id="48bdf-108">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>
    
2. <span data-ttu-id="48bdf-109">Ведущее приложение Office читает XML-манифест надстройки из AppSource, каталога надстроек в SharePoint или каталога общей папки, в зависимости от того, откуда берется надстройка.</span><span class="sxs-lookup"><span data-stu-id="48bdf-109">The Office host application reads the add-in's XML manifest from AppSource, an add-in catalog on SharePoint, or the shared folder catalog it originates from.</span></span>
    
3. <span data-ttu-id="48bdf-110">Ведущее приложение Office открывает HTML-страницу надстройки в элементе управления браузера.</span><span class="sxs-lookup"><span data-stu-id="48bdf-110">The Office host application opens the add-in's HTML page in a browser control.</span></span>
    
    <span data-ttu-id="48bdf-p101">Следующие два действия, 4 и 5, выполняются одновременно и параллельно. Поэтому код надстройки перед обработкой должен убедиться, что и модель DOM, и среда выполнения надстройки полностью загрузились.</span><span class="sxs-lookup"><span data-stu-id="48bdf-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>
    
4. <span data-ttu-id="48bdf-113">Элемент управления браузера загружает модель DOM и основной текст HTML, а также вызывает обработчик для события  **window.onload**.</span><span class="sxs-lookup"><span data-stu-id="48bdf-113">The browser control loads the DOM and HTML body, and calls the event handler for the  **window.onload** event.</span></span>
    
5. <span data-ttu-id="48bdf-114">Ведущее приложение Office загружает среду выполнения, которая загружает и кэширует API JavaScript для файлов библиотеки JavaScript с сервера сети доставки содержимого, а затем вызывает обработчик события [инициализации](https://docs.microsoft.com/javascript/api/office) объекта [Office](https://docs.microsoft.com/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="48bdf-114">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](https://docs.microsoft.com/javascript/api/office) event of the [Office](https://docs.microsoft.com/javascript/api/office) object.</span></span>
    
6. <span data-ttu-id="48bdf-115">После завершения загрузки DOM и основного текста HTML и инициализации надстройки запускается основная функция надстройки.</span><span class="sxs-lookup"><span data-stu-id="48bdf-115">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>
    

## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="48bdf-116">Запуск надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="48bdf-116">Startup of an Outlook add-in</span></span>



<span data-ttu-id="48bdf-117">На рисунке ниже приведен поток событий при запуске надстройки Outlook на настольном компьютере, планшетном ПК или смартфоне.</span><span class="sxs-lookup"><span data-stu-id="48bdf-117">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Поток событий при запуске надстройки Outlook](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="48bdf-119">При запуске надстройки Outlook происходят указанные ниже события.</span><span class="sxs-lookup"><span data-stu-id="48bdf-119">The following events occur when an Outlook add-in starts:</span></span> 



1. <span data-ttu-id="48bdf-120">При запуске Outlook считывает XML-манифесты надстроек Outlook, установленных для учетной записи пользователя.</span><span class="sxs-lookup"><span data-stu-id="48bdf-120">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>
    
2. <span data-ttu-id="48bdf-121">Пользователь выбирает элемент в Outlook.</span><span class="sxs-lookup"><span data-stu-id="48bdf-121">The user selects an item in Outlook.</span></span>
    
3. <span data-ttu-id="48bdf-122">Если выбранный элемент удовлетворяет условиям активации надстройки Outlook, то Outlook активирует надстройку и делает соответствующую кнопку видимой в пользовательском интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="48bdf-122">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>
    
4. <span data-ttu-id="48bdf-p102">Если пользователь нажимает кнопку для запуска надстройки Outlook, то ведущее приложение открывает HTML-страницу в элементе управления браузером. Следующие два шага, шаг 5 и шаг 6, выполняются одновременно.</span><span class="sxs-lookup"><span data-stu-id="48bdf-p102">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>
    
5. <span data-ttu-id="48bdf-125">Элемент управления браузером загружает DOM и основной текст HTML и вызывает обработчик события  **onload**.</span><span class="sxs-lookup"><span data-stu-id="48bdf-125">The browser control loads the DOM and HTML body, and calls the event handler for the  **onload** event.</span></span>
    
6. <span data-ttu-id="48bdf-126">Outlook вызывает обработчик события [initialize](https://docs.microsoft.com/javascript/api/office) объекта [Office](https://docs.microsoft.com/javascript/api/office) надстройки.</span><span class="sxs-lookup"><span data-stu-id="48bdf-126">Outlook calls the event handler for the [initialize](https://docs.microsoft.com/javascript/api/office) event of the [Office](https://docs.microsoft.com/javascript/api/office) object of the add-in.</span></span>
    
7. <span data-ttu-id="48bdf-127">После завершения загрузки DOM и основного текста HTML и инициализации надстройки запускается основная функция надстройки.</span><span class="sxs-lookup"><span data-stu-id="48bdf-127">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>
    

## <a name="checking-the-load-status"></a><span data-ttu-id="48bdf-128">Проверка состояния загрузки</span><span class="sxs-lookup"><span data-stu-id="48bdf-128">Checking the load status</span></span>


<span data-ttu-id="48bdf-p103">Один из способов проверки завершения загрузки DOM и среды выполнения надстроек для — это возможность использования функции [.ready()](https://api.jquery.com/ready/) jQuery — `$(document).ready()`. Например, следующая функция обработчика событий  **initialize** убеждается в полной загрузке DOM, прежде чем выполняется код, относящийся к инициализации надстроек. После этого обработчик события **initialize**переходит на использование текущего выбранного элемента в Outlook, а обработчик событий переходит на использование свойства [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox) для получения выбранного в настоящий момент элемента Outlook и вызывает основную функцию надстройки `initDialer`.</span><span class="sxs-lookup"><span data-stu-id="48bdf-p103">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`. For example, the following  **initialize** event handler function makes sure the DOM is first loaded before the code specific to initializing the add-in runs. Subsequently, the **initialize** event handler proceeds to use the [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>


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

<span data-ttu-id="48bdf-132">Эта же техника может использоваться в обработчике событий  **initialize** любого приложения Надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="48bdf-132">This same technique can be used in the  **initialize** handler of any Office Add-in.</span></span>

<span data-ttu-id="48bdf-133">В примере надстройки Outlook "Телефон" показан несколько другой подход, использующий только JavaScript для проверки тех же условий.</span><span class="sxs-lookup"><span data-stu-id="48bdf-133">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="48bdf-134">Даже если у надстройки нет задач инициализации, необходимо включить по крайней мере минимальную функцию обработчика событий **Office.initialize**, как в примере ниже.</span><span class="sxs-lookup"><span data-stu-id="48bdf-134">Even if your add-in has no initialization tasks to perform, you must include at least a minimal **Office.initialize** event handler function like the following example.</span></span>

```js
Office.initialize = function () {
};
```

<span data-ttu-id="48bdf-p104">Если вы не включите обработчик событий  **Office.initialize**, надстройка может выдать ошибку при запуске. Кроме того, если пользователь попытается применить надстройку с веб-клиентом Office Online, например Excel Online, PowerPoint Online или Outlook Web App, она не будет работать.</span><span class="sxs-lookup"><span data-stu-id="48bdf-p104">If you fail to include an  **Office.initialize** event handler, your add-in may raise an error when it starts. Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run.</span></span>

<span data-ttu-id="48bdf-137">Если надстройка содержит несколько страниц, каждая загружаемая страница должна содержать обработчик событий  **Office.initialize** или вызывать его.</span><span class="sxs-lookup"><span data-stu-id="48bdf-137">If your add-in includes more than one page, whenever it loads a new page that page must include or call an  **Office.initialize** event handler.</span></span>


## <a name="see-also"></a><span data-ttu-id="48bdf-138">См. также</span><span class="sxs-lookup"><span data-stu-id="48bdf-138">See also</span></span>

- [<span data-ttu-id="48bdf-139">Общие сведения об интерфейсе API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="48bdf-139">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
    

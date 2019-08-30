---
title: Обзор создания кода с помощью API JavaScript для OneNote
description: ''
ms.date: 07/05/2019
ms.topic: overview
scenarios: getting-started
localization_priority: Priority
ms.openlocfilehash: 214a74c29d748b0354c4dc75f85c3ada442214d5
ms.sourcegitcommit: 49af31060aa56c1e1ec1e08682914d3cbefc3f1c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/29/2019
ms.locfileid: "36672805"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="a66a1-102">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="a66a1-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="a66a1-103">В OneNote представлен API JavaScript для надстроек OneNote в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a66a1-103">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="a66a1-104">Вы можете создавать надстройки области задач, контентные надстройки и команды надстроек, которые взаимодействуют с объектами OneNote и подключаются к веб-службам или другим веб-ресурсам.</span><span class="sxs-lookup"><span data-stu-id="a66a1-104">OneNote introduces a JavaScript API for OneNote Online add-ins. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="a66a1-p102">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="a66a1-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="a66a1-107">Компоненты надстройки Office</span><span class="sxs-lookup"><span data-stu-id="a66a1-107">Components of an Office Add-in</span></span>

<span data-ttu-id="a66a1-108">Надстройки состоят из двух указанных ниже основных компонентов.</span><span class="sxs-lookup"><span data-stu-id="a66a1-108">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="a66a1-109">**Веб-приложение**, состоящее из веб-страницы и необходимых JavaScript-, CSS- или других файлов.</span><span class="sxs-lookup"><span data-stu-id="a66a1-109">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="a66a1-110">Эти файлы можно разместить на веб-сервере или в службе веб-хостинга, например в Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="a66a1-110">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="a66a1-111">В OneNote в Интернете веб-приложение отображается в элементе управления браузера или в iFrame.</span><span class="sxs-lookup"><span data-stu-id="a66a1-111">In OneNote Online, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="a66a1-p104">**Манифест в формате XML**, в котором указан URL-адрес веб-страницы надстройки и все требования, необходимые для получения доступа, параметры и возможности для надстройки. Этот файл хранится на клиентском компьютере. Для надстроек OneNote используется такой же формат [манифеста](../develop/add-in-manifests.md), что и для других надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="a66a1-p104">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="a66a1-115">**Надстройка Office = манифест + веб-страница**</span><span class="sxs-lookup"><span data-stu-id="a66a1-115">**Office Add-in = Manifest + Webpage**</span></span>

![Надстройка Office состоит из манифеста и веб-страницы](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="a66a1-117">Использование API JavaScript</span><span class="sxs-lookup"><span data-stu-id="a66a1-117">Using the JavaScript API</span></span>

<span data-ttu-id="a66a1-p105">Для доступа к API JavaScript надстройки используют контекст среды выполнения ведущего приложения. API состоит из двух указанных ниже уровней.</span><span class="sxs-lookup"><span data-stu-id="a66a1-p105">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="a66a1-120">**API для определенных ведущих приложений** для связанных с OneNote операций, доступ к которому осуществляется с помощью объекта **Application**.</span><span class="sxs-lookup"><span data-stu-id="a66a1-120">A **host-specific API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="a66a1-121">**Общий API**, используемый приложениями Office, доступ к которому осуществляется с помощью объекта **Document**.</span><span class="sxs-lookup"><span data-stu-id="a66a1-121">A **Common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-host-specific-api-through-the-application-object"></a><span data-ttu-id="a66a1-122">Доступ к API для определенных ведущих приложений с помощью объекта *Application*</span><span class="sxs-lookup"><span data-stu-id="a66a1-122">Accessing the host-specific API through the *Application* object</span></span>

<span data-ttu-id="a66a1-123">Для доступа к объектам OneNote, например к объектам **Notebook**, **Section** и **Page**, используйте объект **Application**.</span><span class="sxs-lookup"><span data-stu-id="a66a1-123">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="a66a1-124">С помощью API для определенных ведущих приложений вы можете запустить пакетные операции на прокси-объектах.</span><span class="sxs-lookup"><span data-stu-id="a66a1-124">With host-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="a66a1-125">Основной процесс выглядит примерно так, как указано ниже.</span><span class="sxs-lookup"><span data-stu-id="a66a1-125">The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="a66a1-126">Получение экземпляра приложения из контекста.</span><span class="sxs-lookup"><span data-stu-id="a66a1-126">Get the application instance from the context.</span></span>

2. <span data-ttu-id="a66a1-p107">Создание прокси-объекта, представляющего объект OneNote, с которым вам необходимо работать. Для синхронного взаимодействия с прокси-объектами можно считывать и записывать их свойства и вызывать имеющиеся в них методы.</span><span class="sxs-lookup"><span data-stu-id="a66a1-p107">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="a66a1-p108">Вызовите метод **load** прокси-объекта, чтобы указать для него значения свойств, указанные в параметре. Этот вызов будет добавлен в очередь команд.</span><span class="sxs-lookup"><span data-stu-id="a66a1-p108">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="a66a1-131">Вызовы, которые методы совершают к API (например, `context.application.getActiveSection().pages;`), также добавляются в очередь.</span><span class="sxs-lookup"><span data-stu-id="a66a1-131">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="a66a1-p109">Чтобы запустить все поставленные в очередь команды в том порядке, в котором они находятся в очереди, вызовите метод **context.sync**. Этот метод синхронизирует состояния выполняющихся сценариев и реальных объектов, а также получает свойства загруженных объектов OneNote, которые необходимо использовать в сценарии. Вы можете использовать возвращенный объект обещания для связывания дополнительных действий в цепочку.</span><span class="sxs-lookup"><span data-stu-id="a66a1-p109">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="a66a1-135">Примеры:</span><span class="sxs-lookup"><span data-stu-id="a66a1-135">For example:</span></span>

```js
function getPagesInSection() {
    OneNote.run(function (context) {

        // Get the pages in the current section.
        var pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function () {

                // Read the id and title of each page.
                $.each(pages.items, function(index, page) {
                    var pageId = page.id;
                    var pageTitle = page.title;
                    console.log(pageTitle + ': ' + pageId);
                });
            })
            .catch(function (error) {
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    });
}
```

<span data-ttu-id="a66a1-136">Сведения о поддерживаемых объектах и операциях OneNote см. в [справочнике по API](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference).</span><span class="sxs-lookup"><span data-stu-id="a66a1-136">You can find supported OneNote objects and operations in the [API reference](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference).</span></span>

#### <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="a66a1-137">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="a66a1-137">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="a66a1-138">Наборы обязательных элементов — именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="a66a1-138">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="a66a1-139">Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API.</span><span class="sxs-lookup"><span data-stu-id="a66a1-139">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="a66a1-140">Дополнительные сведения о наборах обязательных элементов API JavaScript для OneNote см. в статье [Наборы обязательных элементов API JavaScript для OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="a66a1-140">For detailed information about OneNote JavaScript API requirement sets, see the [OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md) article.</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="a66a1-141">Получение доступа к общему API с помощью объекта *Document*</span><span class="sxs-lookup"><span data-stu-id="a66a1-141">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="a66a1-142">Для доступа к общему API, например к методам [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) и [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-), используйте объект **Document**.</span><span class="sxs-lookup"><span data-stu-id="a66a1-142">Use the **Document** object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span> 


<span data-ttu-id="a66a1-143">Пример:</span><span class="sxs-lookup"><span data-stu-id="a66a1-143">For example:</span></span>  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

<span data-ttu-id="a66a1-144">Надстройки OneNote поддерживают только указанные ниже общие API.</span><span class="sxs-lookup"><span data-stu-id="a66a1-144">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="a66a1-145">API</span><span class="sxs-lookup"><span data-stu-id="a66a1-145">API</span></span> | <span data-ttu-id="a66a1-146">Примечания</span><span class="sxs-lookup"><span data-stu-id="a66a1-146">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="a66a1-147">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a66a1-147">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="a66a1-148">Только **Office.CoercionType.Text** и **Office.CoercionType.Matrix**</span><span class="sxs-lookup"><span data-stu-id="a66a1-148">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="a66a1-149">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a66a1-149">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="a66a1-150">Только **Office.CoercionType.Text**, **Office.CoercionType.Image** и **Office.CoercionType.Html**</span><span class="sxs-lookup"><span data-stu-id="a66a1-150">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| [<span data-ttu-id="a66a1-151">var mySetting = Office.context.document.settings.get(имя);</span><span class="sxs-lookup"><span data-stu-id="a66a1-151">var mySetting = Office.context.document.settings.get(name);</span></span>](/javascript/api/office/office.settings#get-name-) | <span data-ttu-id="a66a1-152">Параметры поддерживаются только контентными надстройками</span><span class="sxs-lookup"><span data-stu-id="a66a1-152">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="a66a1-153">Office.context.document.settings.set(имя, значение);</span><span class="sxs-lookup"><span data-stu-id="a66a1-153">Office.context.document.settings.set(name, value);</span></span>](/javascript/api/office/office.settings#set-name--value-) | <span data-ttu-id="a66a1-154">Параметры поддерживаются только контентными надстройками</span><span class="sxs-lookup"><span data-stu-id="a66a1-154">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="a66a1-155">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="a66a1-155">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="a66a1-156">Обычно общий API следует использовать только тогда, когда необходимые возможности не поддерживаются в API для определенных ведущих приложений.</span><span class="sxs-lookup"><span data-stu-id="a66a1-156">In general, you only use the Common API to do something that isn't supported in the host-specific API.</span></span> <span data-ttu-id="a66a1-157">Дополнительные сведения об использовании общего API см. в [документации](../overview/office-add-ins.md) и [справочнике](../reference/javascript-api-for-office.md) по надстройкам Office.</span><span class="sxs-lookup"><span data-stu-id="a66a1-157">To learn more about using the Common API, see the Office Add-ins [documentation](../overview/office-add-ins.md) and [reference](../reference/javascript-api-for-office.md).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="a66a1-158">Схема объектной модели OneNote</span><span class="sxs-lookup"><span data-stu-id="a66a1-158">OneNote object model diagram</span></span> 
<span data-ttu-id="a66a1-159">На схеме ниже показаны возможности, которые на данный момент доступны в API JavaScript для OneNote .</span><span class="sxs-lookup"><span data-stu-id="a66a1-159">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Схема объектной модели OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="a66a1-161">См. также</span><span class="sxs-lookup"><span data-stu-id="a66a1-161">See also</span></span>

- [<span data-ttu-id="a66a1-162">Документация по надстройкам OneNote</span><span class="sxs-lookup"><span data-stu-id="a66a1-162">OneNote add-ins documentation</span></span>](index.md)
- [<span data-ttu-id="a66a1-163">Создание первой надстройки OneNote</span><span class="sxs-lookup"><span data-stu-id="a66a1-163">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="a66a1-164">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="a66a1-164">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="a66a1-165">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="a66a1-165">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="a66a1-166">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a66a1-166">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

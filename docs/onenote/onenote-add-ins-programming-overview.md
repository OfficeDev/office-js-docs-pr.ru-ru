---
title: Обзор создания кода с помощью API JavaScript для OneNote
description: Узнайте об API OneNote JavaScript для надстроек OneNote в Интернете.
ms.date: 03/18/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: ae88c2bba6c23a2c3ec3358db121a2ca3630f09d
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891056"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="61195-103">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="61195-103">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="61195-104">В OneNote представлен API JavaScript для надстроек OneNote в Интернете.</span><span class="sxs-lookup"><span data-stu-id="61195-104">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="61195-105">Вы можете создавать надстройки области задач, контентные надстройки и команды надстроек, которые взаимодействуют с объектами OneNote и подключаются к веб-службам или другим веб-ресурсам.</span><span class="sxs-lookup"><span data-stu-id="61195-105">You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="61195-106">Компоненты надстройки Office</span><span class="sxs-lookup"><span data-stu-id="61195-106">Components of an Office Add-in</span></span>

<span data-ttu-id="61195-107">Надстройки состоят из двух указанных ниже основных компонентов.</span><span class="sxs-lookup"><span data-stu-id="61195-107">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="61195-108">**Веб-приложение**, состоящее из веб-страницы и необходимых JavaScript-, CSS- или других файлов.</span><span class="sxs-lookup"><span data-stu-id="61195-108">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="61195-109">Эти файлы можно разместить на веб-сервере или в службе веб-хостинга, например в Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="61195-109">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="61195-110">В OneNote в Интернете веб-приложение отображается в элементе управления браузера или в iFrame.</span><span class="sxs-lookup"><span data-stu-id="61195-110">In OneNote on the web, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="61195-p103">**Манифест в формате XML**, в котором указан URL-адрес веб-страницы надстройки и все требования, необходимые для получения доступа, параметры и возможности для надстройки. Этот файл хранится на клиентском компьютере. Для надстроек OneNote используется такой же формат [манифеста](../develop/add-in-manifests.md), что и для других надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="61195-p103">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="61195-114">**Надстройка Office = манифест + веб-страница**</span><span class="sxs-lookup"><span data-stu-id="61195-114">**Office Add-in = Manifest + Webpage**</span></span>

![Надстройка Office состоит из манифеста и веб-страницы](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="61195-116">Использование API JavaScript</span><span class="sxs-lookup"><span data-stu-id="61195-116">Using the JavaScript API</span></span>

<span data-ttu-id="61195-p104">Для доступа к API JavaScript надстройки используют контекст среды выполнения ведущего приложения. API состоит из двух указанных ниже уровней.</span><span class="sxs-lookup"><span data-stu-id="61195-p104">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span>

- <span data-ttu-id="61195-119">**API для определенных ведущих приложений** для связанных с OneNote операций, доступ к которому осуществляется с помощью объекта `Application`.</span><span class="sxs-lookup"><span data-stu-id="61195-119">A **host-specific API** for OneNote-specific operations, accessed through the `Application` object.</span></span>
- <span data-ttu-id="61195-120">**Общий API**, используемый приложениями Office, доступ к которому осуществляется с помощью объекта `Document`.</span><span class="sxs-lookup"><span data-stu-id="61195-120">A **Common API** that's shared across Office applications, accessed through the `Document` object.</span></span>

### <a name="accessing-the-host-specific-api-through-the-application-object"></a><span data-ttu-id="61195-121">Доступ к API для определенных ведущих приложений с помощью объекта *Application*</span><span class="sxs-lookup"><span data-stu-id="61195-121">Accessing the host-specific API through the *Application* object</span></span>

<span data-ttu-id="61195-122">Для доступа к объектам OneNote, например к объектам **Notebook**, **Section** и **Page**, используйте объект `Application`.</span><span class="sxs-lookup"><span data-stu-id="61195-122">Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="61195-123">С помощью API для определенных ведущих приложений вы можете запустить пакетные операции на прокси-объектах.</span><span class="sxs-lookup"><span data-stu-id="61195-123">With host-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="61195-124">Основной процесс выглядит примерно так, как указано ниже.</span><span class="sxs-lookup"><span data-stu-id="61195-124">The basic flow goes something like this:</span></span>

1. <span data-ttu-id="61195-125">Получение экземпляра приложения из контекста.</span><span class="sxs-lookup"><span data-stu-id="61195-125">Get the application instance from the context.</span></span>

2. <span data-ttu-id="61195-p106">Создание прокси-объекта, представляющего объект OneNote, с которым вам необходимо работать. Для синхронного взаимодействия с прокси-объектами можно считывать и записывать их свойства и вызывать имеющиеся в них методы.</span><span class="sxs-lookup"><span data-stu-id="61195-p106">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="61195-p107">Вызовите метод `load` прокси-объекта, чтобы указать для него значения свойств, указанные в параметре. Этот вызов будет добавлен в очередь команд.</span><span class="sxs-lookup"><span data-stu-id="61195-p107">Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="61195-130">Вызовы, которые методы совершают к API (например, `context.application.getActiveSection().pages;`), также добавляются в очередь.</span><span class="sxs-lookup"><span data-stu-id="61195-130">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="61195-p108">Чтобы запустить все поставленные в очередь команды в том порядке, в котором они находятся в очереди, вызовите метод `context.sync`. Этот метод синхронизирует состояния выполняющихся сценариев и реальных объектов, а также получает свойства загруженных объектов OneNote, которые необходимо использовать в сценарии. Вы можете использовать возвращенный объект обещания для связывания дополнительных действий в цепочку.</span><span class="sxs-lookup"><span data-stu-id="61195-p108">Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="61195-134">Например:</span><span class="sxs-lookup"><span data-stu-id="61195-134">For example:</span></span>

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

<span data-ttu-id="61195-135">Сведения о поддерживаемых объектах и операциях OneNote см. в [справочнике по API](../reference/overview/onenote-add-ins-javascript-reference.md).</span><span class="sxs-lookup"><span data-stu-id="61195-135">You can find supported OneNote objects and operations in the [API reference](../reference/overview/onenote-add-ins-javascript-reference.md).</span></span>

#### <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="61195-136">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="61195-136">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="61195-137">Наборы обязательных элементов — именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="61195-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="61195-138">Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API.</span><span class="sxs-lookup"><span data-stu-id="61195-138">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="61195-139">Дополнительные сведения о наборах обязательных элементов API JavaScript для OneNote см. в статье [Наборы обязательных элементов API JavaScript для OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="61195-139">For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="61195-140">Получение доступа к общему API с помощью объекта *Document*</span><span class="sxs-lookup"><span data-stu-id="61195-140">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="61195-141">Для доступа к общему API, например к методам [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) и [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-), используйте объект `Document`.</span><span class="sxs-lookup"><span data-stu-id="61195-141">Use the `Document` object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span>


<span data-ttu-id="61195-142">Например:</span><span class="sxs-lookup"><span data-stu-id="61195-142">For example:</span></span>  

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

<span data-ttu-id="61195-143">Надстройки OneNote поддерживают только указанные ниже общие API.</span><span class="sxs-lookup"><span data-stu-id="61195-143">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="61195-144">API</span><span class="sxs-lookup"><span data-stu-id="61195-144">API</span></span> | <span data-ttu-id="61195-145">Примечания</span><span class="sxs-lookup"><span data-stu-id="61195-145">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="61195-146">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="61195-146">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="61195-147">Только `Office.CoercionType.Text` и `Office.CoercionType.Matrix`</span><span class="sxs-lookup"><span data-stu-id="61195-147">`Office.CoercionType.Text` and `Office.CoercionType.Matrix` only</span></span> |
| [<span data-ttu-id="61195-148">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="61195-148">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="61195-149">Только `Office.CoercionType.Text`, `Office.CoercionType.Image` и `Office.CoercionType.Html`</span><span class="sxs-lookup"><span data-stu-id="61195-149">`Office.CoercionType.Text`, `Office.CoercionType.Image`, and `Office.CoercionType.Html` only</span></span> | 
| [<span data-ttu-id="61195-150">var mySetting = Office.context.document.settings.get(имя);</span><span class="sxs-lookup"><span data-stu-id="61195-150">var mySetting = Office.context.document.settings.get(name);</span></span>](/javascript/api/office/office.settings#get-name-) | <span data-ttu-id="61195-151">Параметры поддерживаются только контентными надстройками</span><span class="sxs-lookup"><span data-stu-id="61195-151">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="61195-152">Office.context.document.settings.set(имя, значение);</span><span class="sxs-lookup"><span data-stu-id="61195-152">Office.context.document.settings.set(name, value);</span></span>](/javascript/api/office/office.settings#set-name--value-) | <span data-ttu-id="61195-153">Параметры поддерживаются только контентными надстройками</span><span class="sxs-lookup"><span data-stu-id="61195-153">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="61195-154">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="61195-154">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="61195-155">Обычно общий API следует использовать, когда необходимые возможности не поддерживаются в API для определенных ведущих приложений.</span><span class="sxs-lookup"><span data-stu-id="61195-155">In general, you use the Common API to do something that isn't supported in the host-specific API.</span></span> <span data-ttu-id="61195-156">Дополнительные сведения об использовании общего API см. в статье [Общая объектная модель API JavaScript](../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="61195-156">To learn more about using the Common API, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="61195-157">Схема объектной модели OneNote</span><span class="sxs-lookup"><span data-stu-id="61195-157">OneNote object model diagram</span></span> 
<span data-ttu-id="61195-158">На схеме ниже показаны возможности, которые на данный момент доступны в API JavaScript для OneNote .</span><span class="sxs-lookup"><span data-stu-id="61195-158">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Схема объектной модели OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="61195-160">См. также</span><span class="sxs-lookup"><span data-stu-id="61195-160">See also</span></span>

- [<span data-ttu-id="61195-161">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="61195-161">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="61195-162">Создание первой надстройки OneNote</span><span class="sxs-lookup"><span data-stu-id="61195-162">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="61195-163">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="61195-163">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="61195-164">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="61195-164">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="61195-165">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="61195-165">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

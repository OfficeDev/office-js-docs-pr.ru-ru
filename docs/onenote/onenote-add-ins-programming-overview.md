---
title: Обзор создания кода с помощью API JavaScript для OneNote
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: a2161484991888a6a8c834ba61398cc8c0afb955
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925173"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="c4f20-102">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="c4f20-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="c4f20-103">В OneNote представлен API JavaScript для надстроек OneNote Online. Вы можете создавать надстройки области задач, контентные надстройки и команды надстроек, которые взаимодействуют с объектами OneNote и подключаются к веб-службам или другим веб-ресурсам.</span><span class="sxs-lookup"><span data-stu-id="c4f20-103">OneNote introduces a JavaScript API for OneNote Online add-ins. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="c4f20-p101">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="c4f20-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="c4f20-106">Компоненты надстройки Office</span><span class="sxs-lookup"><span data-stu-id="c4f20-106">Components of an Office Add-in</span></span>

<span data-ttu-id="c4f20-107">Надстройки состоят из двух указанных ниже основных компонентов.</span><span class="sxs-lookup"><span data-stu-id="c4f20-107">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="c4f20-p102">**Веб-приложение**, состоящее из веб-страницы и необходимых JavaScript-, CSS- или других файлов. Эти файлы можно разместить на веб-сервере или в службе веб-хостинга, например в Microsoft Azure. В OneNote Online веб-приложение отображается в элементе управления браузера или в iFrame.</span><span class="sxs-lookup"><span data-stu-id="c4f20-p102">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote Online, the web application displays in a browser control or iframe.</span></span>
    
- <span data-ttu-id="c4f20-p103">**Манифест в формате XML**, в котором указан URL-адрес веб-страницы надстройки и все требования, необходимые для получения доступа, параметры и возможности для надстройки. Этот файл хранится на клиентском компьютере. Для надстроек OneNote используется такой же формат [манифеста](../develop/add-in-manifests.md), что и для других надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="c4f20-p103">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="c4f20-114">**Надстройка Office = манифест + веб-страница**</span><span class="sxs-lookup"><span data-stu-id="c4f20-114">**Office Add-in = Manifest + Webpage**</span></span>

![Надстройка Office состоит из манифеста и веб-страницы](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="c4f20-116">Использование API JavaScript</span><span class="sxs-lookup"><span data-stu-id="c4f20-116">Using the JavaScript API</span></span>

<span data-ttu-id="c4f20-p104">Для доступа к API JavaScript надстройки используют контекст среды выполнения ведущего приложения. API состоит из двух указанных ниже уровней.</span><span class="sxs-lookup"><span data-stu-id="c4f20-p104">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="c4f20-119">**Многофункциональный API** для связанных с OneNote операций, доступ к которому осуществляется с помощью объекта **Application**.</span><span class="sxs-lookup"><span data-stu-id="c4f20-119">A **rich API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="c4f20-120">**Стандартный API**, используемый приложениями Office, доступ к которому осуществляется с помощью объекта **Document**.</span><span class="sxs-lookup"><span data-stu-id="c4f20-120">A **common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-rich-api-through-the-application-object"></a><span data-ttu-id="c4f20-121">Доступ к многофункциональному API с помощью объекта *Application*</span><span class="sxs-lookup"><span data-stu-id="c4f20-121">Accessing the rich API through the *Application* object</span></span>

<span data-ttu-id="c4f20-p105">Для доступа к объектам OneNote, например к объектам **Notebook**, **Section** и **Page**, используйте объект **Application**. С помощью многофункциональных API вы можете запустить пакетные операции на прокси-объектах. Основной процесс выглядит примерно так, как указано ниже.</span><span class="sxs-lookup"><span data-stu-id="c4f20-p105">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With rich APIs, you run batch operations on proxy objects. The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="c4f20-125">Получение экземпляра приложения из контекста.</span><span class="sxs-lookup"><span data-stu-id="c4f20-125">Get the application instance from the context.</span></span>

2. <span data-ttu-id="c4f20-p106">Создание прокси-объекта, представляющего объект OneNote, с которым вам необходимо работать. Для синхронного взаимодействия с прокси-объектами можно считывать и записывать их свойства и вызывать имеющиеся в них методы.</span><span class="sxs-lookup"><span data-stu-id="c4f20-p106">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span> 

3. <span data-ttu-id="c4f20-p107">Вызовите метод **load** прокси-объекта, чтобы указать для него значения свойств, указанные в параметре. Этот вызов будет добавлен в очередь команд.</span><span class="sxs-lookup"><span data-stu-id="c4f20-p107">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="c4f20-130">Вызовы, которые методы совершают к API (например, `context.application.getActiveSection().pages;`), также добавляются в очередь.</span><span class="sxs-lookup"><span data-stu-id="c4f20-130">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="c4f20-p108">Чтобы запустить все поставленные в очередь команды в том порядке, в котором они находятся в очереди, вызовите метод **context.sync**. Этот метод синхронизирует состояния выполняющихся сценариев и реальных объектов, а также получает свойства загруженных объектов OneNote, которые необходимо использовать в сценарии. Вы можете использовать возвращенный объект обещания для связывания дополнительных действий в цепочку.</span><span class="sxs-lookup"><span data-stu-id="c4f20-p108">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="c4f20-134">Примеры:</span><span class="sxs-lookup"><span data-stu-id="c4f20-134">For example:</span></span> 

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

<span data-ttu-id="c4f20-135">Сведения о поддерживаемых объектах и операциях OneNote см. в [справочнике по API](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).</span><span class="sxs-lookup"><span data-stu-id="c4f20-135">You can find supported OneNote objects and operations in the [API reference](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="c4f20-136">Получение доступа к стандартному API с помощью объекта *Document*</span><span class="sxs-lookup"><span data-stu-id="c4f20-136">Accessing the common API through the *Document* object</span></span>

<span data-ttu-id="c4f20-137">Для доступа к стандартному API, например к методам **getSelectedDataAsync** и [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), используйте объект [Document](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync).</span><span class="sxs-lookup"><span data-stu-id="c4f20-137">Use the **Document** object to access the common API, such as the [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) and [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) methods.</span></span> 

<span data-ttu-id="c4f20-138">Например:</span><span class="sxs-lookup"><span data-stu-id="c4f20-138">For example:</span></span>  

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
<span data-ttu-id="c4f20-139">Надстройки OneNote поддерживают только указанные ниже стандартные API.</span><span class="sxs-lookup"><span data-stu-id="c4f20-139">OneNote add-ins support only the following common APIs:</span></span>

| <span data-ttu-id="c4f20-140">API</span><span class="sxs-lookup"><span data-stu-id="c4f20-140">API</span></span> | <span data-ttu-id="c4f20-141">Примечания</span><span class="sxs-lookup"><span data-stu-id="c4f20-141">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="c4f20-142">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c4f20-142">Office.context.document.getSelectedDataAsync</span></span>](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) | <span data-ttu-id="c4f20-143">Только **Office.CoercionType.Text** и **Office.CoercionType.Matrix**</span><span class="sxs-lookup"><span data-stu-id="c4f20-143">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="c4f20-144">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c4f20-144">Office.context.document.setSelectedDataAsync</span></span>](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) | <span data-ttu-id="c4f20-145">Только **Office.CoercionType.Text**, **Office.CoercionType.Image** и **Office.CoercionType.Html**</span><span class="sxs-lookup"><span data-stu-id="c4f20-145">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| [<span data-ttu-id="c4f20-146">var mySetting = Office.context.document.settings.get(имя);</span><span class="sxs-lookup"><span data-stu-id="c4f20-146">var mySetting = Office.context.document.settings.get(name);</span></span>](https://dev.office.com/reference/add-ins/shared/settings.get) | <span data-ttu-id="c4f20-147">Параметры поддерживаются только контентными надстройками</span><span class="sxs-lookup"><span data-stu-id="c4f20-147">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="c4f20-148">Office.context.document.settings.set(имя, значение);</span><span class="sxs-lookup"><span data-stu-id="c4f20-148">Office.context.document.settings.set(name, value);</span></span>](https://dev.office.com/reference/add-ins/shared/settings.set) | <span data-ttu-id="c4f20-149">Параметры поддерживаются только контентными надстройками</span><span class="sxs-lookup"><span data-stu-id="c4f20-149">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="c4f20-150">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="c4f20-150">Office.EventType.DocumentSelectionChanged</span></span>](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

<span data-ttu-id="c4f20-p109">В общем случае стандартный API следует использовать только тогда, когда необходимые возможности не поддерживаются в многофункциональном API. Дополнительные сведения об использовании стандартного API см. в [документации](../overview/office-add-ins.md) и [справочнике](https://dev.office.com/reference/add-ins/javascript-api-for-office) по надстройкам Office.</span><span class="sxs-lookup"><span data-stu-id="c4f20-p109">In general, you only use the common API to do something that isn't supported in the rich API. To learn more about using the common API, see the Office Add-ins [documentation](../overview/office-add-ins.md) and [reference](https://dev.office.com/reference/add-ins/javascript-api-for-office).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="c4f20-153">Схема объектной модели OneNote</span><span class="sxs-lookup"><span data-stu-id="c4f20-153">OneNote object model diagram</span></span> 
<span data-ttu-id="c4f20-154">На схеме ниже показаны возможности, которые на данный момент доступны в API JavaScript для OneNote .</span><span class="sxs-lookup"><span data-stu-id="c4f20-154">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Схема объектной модели OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="c4f20-156">См. также</span><span class="sxs-lookup"><span data-stu-id="c4f20-156">See also</span></span>

- [<span data-ttu-id="c4f20-157">Создание первой надстройки OneNote</span><span class="sxs-lookup"><span data-stu-id="c4f20-157">Build your first OneNote add-in</span></span>](onenote-add-ins-getting-started.md)
- [<span data-ttu-id="c4f20-158">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="c4f20-158">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="c4f20-159">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="c4f20-159">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="c4f20-160">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c4f20-160">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

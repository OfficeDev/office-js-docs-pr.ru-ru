---
title: Обзор API JavaScript для Visio
description: Обзор API JavaScript для Visio
ms.date: 06/03/2020
ms.prod: visio
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 9d0abb5ddc93419f5acd38a8c0134941e15be48b
ms.sourcegitcommit: fecad2afa7938d7178456c11ba52b558224813b4
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/09/2020
ms.locfileid: "49603794"
---
# <a name="visio-javascript-api-overview"></a><span data-ttu-id="0cdfa-103">Обзор API JavaScript для Visio</span><span class="sxs-lookup"><span data-stu-id="0cdfa-103">Visio JavaScript API overview</span></span>

<span data-ttu-id="0cdfa-104">С помощью API JavaScript для Visio вы можете встраивать схемы Visio в *классические* страницы SharePoint в SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-104">You can use the Visio JavaScript APIs to embed Visio diagrams in *classic* SharePoint pages in SharePoint Online.</span></span> <span data-ttu-id="0cdfa-105">(эта возможность расширяемости не поддерживается на страницах локальный среды SharePoint или на страницах SharePoint Framework.)</span><span class="sxs-lookup"><span data-stu-id="0cdfa-105">(This extensibility feature is not supported in on-premise SharePoint or on SharePoint Framework pages.)</span></span>

<span data-ttu-id="0cdfa-106">Внедренный документ Visio — схема, которая хранится в библиотеке документов SharePoint и отображается на странице SharePoint.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-106">An embedded Visio diagram is a diagram that is stored in a SharePoint document library and displayed on a SharePoint page.</span></span> <span data-ttu-id="0cdfa-107">Чтобы внедрить документ Visio, отобразите его в элементе `<iframe>` HTML.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-107">To embed a Visio diagram, display it in an HTML `<iframe>` element.</span></span> <span data-ttu-id="0cdfa-108">После этого вы сможете программным способом работать с внедренным документом при помощи API JavaScript для Visio.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-108">Then you can use Visio JavaScript APIs to programmatically work with the embedded diagram.</span></span>

![Документ Visio в iframe на странице SharePoint вместе с веб-частью редактора сценариев](../images/visio-api-block-diagram.png)

<span data-ttu-id="0cdfa-110">API JavaScript для Visio позволяет следующее:</span><span class="sxs-lookup"><span data-stu-id="0cdfa-110">You can use the Visio JavaScript APIs to:</span></span>

* <span data-ttu-id="0cdfa-111">работать с элементами документа Visio как со страницами и фигурами;</span><span class="sxs-lookup"><span data-stu-id="0cdfa-111">Interact with Visio diagram elements like pages and shapes.</span></span>
* <span data-ttu-id="0cdfa-112">создавать визуальную разметку на холсте документа Visio;</span><span class="sxs-lookup"><span data-stu-id="0cdfa-112">Create visual markup on the Visio diagram canvas.</span></span>
* <span data-ttu-id="0cdfa-113">создавать специальные обработчики событий мыши для документа;</span><span class="sxs-lookup"><span data-stu-id="0cdfa-113">Write custom handlers for mouse events within the drawing.</span></span>
* <span data-ttu-id="0cdfa-114">предоставлять своему решению данные документа, такие как текст фигуры, данные фигуры и гиперссылки.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-114">Expose diagram data, such as shape text, shape data, and hyperlinks, to your solution.</span></span>

<span data-ttu-id="0cdfa-p103">В этой статье описано, как использовать API JavaScript для Visio с приложением Visio в Интернете, чтобы создавать решения для SharePoint Online. В ней рассматриваются ключевые понятия, понимание роли которых крайне важно при использовании API, такие как прокси-объекты JavaScript, `EmbeddedSession`, `RequestContext`, а также методы `sync()`, `Visio.run()` и `load()`. В приведенных ниже примерах кода показано применение этих элементов.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-p103">This article describes how to use the Visio JavaScript APIs with Visio on the web to build your solutions for SharePoint Online. It introduces key concepts that are fundamental to using the APIs, such as `EmbeddedSession`, `RequestContext`, and JavaScript proxy objects, and the `sync()`, `Visio.run()`, and `load()` methods. The code examples show you how to apply these concepts.</span></span>

## <a name="embeddedsession"></a><span data-ttu-id="0cdfa-118">EmbeddedSession</span><span class="sxs-lookup"><span data-stu-id="0cdfa-118">EmbeddedSession</span></span>

<span data-ttu-id="0cdfa-119">Объект EmbeddedSession инициализирует взаимодействие между фреймом разработчика и фреймом Visio в браузере.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-119">The EmbeddedSession object initializes communication between the developer frame and the Visio frame in the browser.</span></span>

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a><span data-ttu-id="0cdfa-120">Visio.run(session, function(context) { batch })</span><span class="sxs-lookup"><span data-stu-id="0cdfa-120">Visio.run(session, function(context) { batch })</span></span>

<span data-ttu-id="0cdfa-121">Метод `Visio.run()` выполняет пакетный сценарий, совершающий действия с объектной моделью Visio.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-121">`Visio.run()` executes a batch script that performs actions on the Visio object model.</span></span> <span data-ttu-id="0cdfa-122">Пакетные команды включают определения локальных прокси-объектов JavaScript и методов `sync()`, синхронизирующих состояние объектов Visio и локальных объектов, а также разрешение обещания.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-122">The batch commands include definitions of local JavaScript proxy objects and `sync()` methods that synchronize the state between local and Visio objects and promise resolution.</span></span> <span data-ttu-id="0cdfa-123">Преимущество пакетной обработки запросов в методе `Visio.run()` состоит в том, что при разрешении обещания все отслеживаемые объекты страницы, выделенные во время выполнения, автоматически освобождаются.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-123">The advantage of batching requests in `Visio.run()` is that when the promise is resolved, any tracked page objects that were allocated during the execution will be automatically released.</span></span>

<span data-ttu-id="0cdfa-124">Метод run использует объект session и RequestContext и возвращает обещание (как правило, просто результат выполнения метода `context.sync()`).</span><span class="sxs-lookup"><span data-stu-id="0cdfa-124">The run method takes in session and RequestContext object and returns a promise (typically, just the result of `context.sync()`).</span></span> <span data-ttu-id="0cdfa-125">Пакетную операцию можно выполнить, не указывая ее в методе `Visio.run()`.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-125">It is possible to run the batch operation outside of the `Visio.run()`.</span></span> <span data-ttu-id="0cdfa-126">Однако в этом случае все ссылки на объекты страницы требуют отслеживания и управления вручную.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-126">However, in such a scenario, any page object references needs to be manually tracked and managed.</span></span>

## <a name="requestcontext"></a><span data-ttu-id="0cdfa-127">RequestContext</span><span class="sxs-lookup"><span data-stu-id="0cdfa-127">RequestContext</span></span>

<span data-ttu-id="0cdfa-128">Объект RequestContext обеспечивает отправку запросов приложению Visio.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-128">The RequestContext object facilitates requests to the Visio application.</span></span> <span data-ttu-id="0cdfa-129">Так как фрейм разработчика и веб-клиент Visio отображаются в двух разных элементах iframe, для получения доступа через фрейм разработчика к Visio и связанным объектам, таким как страницы и фигуры, требуется объект RequestContext (контекст показан в следующем примере).</span><span class="sxs-lookup"><span data-stu-id="0cdfa-129">Because the developer frame and the Visio web client run in two different iframes, the RequestContext object (context in next example) is required to get access to Visio and related objects such as pages and shapes, from the developer frame.</span></span>

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a><span data-ttu-id="0cdfa-130">Прокси-объекты</span><span class="sxs-lookup"><span data-stu-id="0cdfa-130">Proxy objects</span></span>

<span data-ttu-id="0cdfa-131">Объекты JavaScript для Visio, объявленные и использованные во встроенном сеансе, — это прокси-объекты для реальных объектов в документе Visio.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-131">The Visio JavaScript objects declared and used in an embedded session are proxy objects for the real objects in a Visio document.</span></span> <span data-ttu-id="0cdfa-132">Все действия над прокси-объектами не реализуются в Visio, а состояние документа Visio — в прокси-объектах, пока оно не будет синхронизировано.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-132">All actions taken on proxy objects are not realized in Visio, and the state of the Visio document is not realized in the proxy objects until the document state has been synchronized.</span></span> <span data-ttu-id="0cdfa-133">Состояние документа синхронизируется при выполнении `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-133">The document state is synchronized when `context.sync()` is run.</span></span>

<span data-ttu-id="0cdfa-134">Например, локальный объект JavaScript getActivePage объявлен в качестве ссылки на выбранную страницу.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-134">For example, the local JavaScript object getActivePage is declared to reference the selected page.</span></span> <span data-ttu-id="0cdfa-135">Это можно использовать для добавления в очередь настройки его свойств и вызова методов.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-135">This can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="0cdfa-136">Действия над такими объектами не реализуются до выполнения метода `sync()`.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-136">The actions on such objects are not realized until the `sync()` method is run.</span></span>

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a><span data-ttu-id="0cdfa-137">sync()</span><span class="sxs-lookup"><span data-stu-id="0cdfa-137">sync()</span></span>

<span data-ttu-id="0cdfa-138">Метод `sync()` синхронизирует состояние прокси-объектов JavaScript и реальных объектов в Visio путем выполнения поставленных в очередь инструкций над контекстом и получения свойств загруженных объектов Office для их использования в коде.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-138">The `sync()` method synchronizes the state between JavaScript proxy objects and real objects in Visio by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.</span></span> <span data-ttu-id="0cdfa-139">Этот метод возвращает обещание, которое выполняется после завершения синхронизации.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-139">This method returns a promise, which is resolved when synchronization is complete.</span></span>

## <a name="load"></a><span data-ttu-id="0cdfa-140">load()</span><span class="sxs-lookup"><span data-stu-id="0cdfa-140">load()</span></span>

<span data-ttu-id="0cdfa-141">Метод `load()` используется для заполнения прокси-объектов, созданных на уровне JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-141">The `load()` method is used to fill in the proxy objects created in the JavaScript layer.</span></span> <span data-ttu-id="0cdfa-142">При попытке получения объекта, такого как документ, сначала на уровне JavaScript создается локальный прокси-объект.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-142">When trying to retrieve an object such as a document, a local proxy object is created first in the JavaScript layer.</span></span> <span data-ttu-id="0cdfa-143">Такой объект можно использовать для добавления в очередь настройки его свойств и вызова методов.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-143">Such an object can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="0cdfa-144">Но для чтения свойств или связей объекта сначала необходимо вызвать методы `load()` и `sync()`.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-144">However, for reading object properties or relations, the `load()` and `sync()` methods need to be invoked first.</span></span> <span data-ttu-id="0cdfa-145">Метод load() использует свойства и связи, которые необходимо загрузить при вызове метода `sync()`.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-145">The load() method takes in the properties and relations that need to be loaded when the `sync()` method is called.</span></span>

<span data-ttu-id="0cdfa-146">Ниже представлен синтаксис метода `load()`.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-146">The following shows the syntax for the `load()` method.</span></span>

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. <span data-ttu-id="0cdfa-147">**properties** — это список имен свойств, которые требуется загрузить, разделенных запятыми, или массив имен.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-147">**properties** is the list of property names to be loaded, specified as comma-delimited strings or array of names.</span></span> <span data-ttu-id="0cdfa-148">Дополнительные сведения см. в описаниях методов `.load()` под каждым объектом.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-148">See `.load()` methods under each object for details.</span></span>

2. <span data-ttu-id="0cdfa-p112">**loadOption** указывает объект, описывающий свойства select, expand, top и skip. Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](/javascript/api/office/officeextension.loadoption).</span><span class="sxs-lookup"><span data-stu-id="0cdfa-p112">**loadOption** specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

## <a name="example-printing-all-shapes-text-in-active-page"></a><span data-ttu-id="0cdfa-151">Пример. Печать текста всех фигур на активной странице</span><span class="sxs-lookup"><span data-stu-id="0cdfa-151">Example: Printing all shapes text in active page</span></span>

<span data-ttu-id="0cdfa-152">Приведенный ниже пример показывает, как распечатать значение текста фигуры из объекта фигур массива.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-152">The following example shows you how to print shape text value from an array shapes object.</span></span>
<span data-ttu-id="0cdfa-153">Метод `Visio.run()` содержит пакет инструкций.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-153">The `Visio.run()` method contains a batch of instructions.</span></span> <span data-ttu-id="0cdfa-154">В рамках этого пакета создается прокси-объект, который ссылается на фигуры в активном документе.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-154">As part of this batch, a proxy object is created that references shapes on the active document.</span></span>

<span data-ttu-id="0cdfa-155">Все эти команды ставятся в очередь и выполняются при вызове метода `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-155">All these commands are queued and run when `context.sync()` is called.</span></span> <span data-ttu-id="0cdfa-156">Метод `sync()` возвращает обещание, с помощью которого его можно связать с другими операциями.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-156">The `sync()` method returns a promise that can be used to chain it with other operations.</span></span>

```js
Visio.run(session, function (context) {
    var page = context.document.getActivePage();
    var shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++) {
            var shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a><span data-ttu-id="0cdfa-157">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="0cdfa-157">Error messages</span></span>

<span data-ttu-id="0cdfa-p115">Ошибки возвращаются с помощью объекта ошибки, состоящего из кода и сообщения. В таблице ниже перечислены возможные ошибки.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-p115">Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.</span></span>

| <span data-ttu-id="0cdfa-160">error.code</span><span class="sxs-lookup"><span data-stu-id="0cdfa-160">error.code</span></span>            | <span data-ttu-id="0cdfa-161">error.message</span><span class="sxs-lookup"><span data-stu-id="0cdfa-161">error.message</span></span> |
|-----------------------|----------------------------------------------------------------|
| <span data-ttu-id="0cdfa-162">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="0cdfa-162">InvalidArgument</span></span>       | <span data-ttu-id="0cdfa-163">Аргумент недопустим, отсутствует или имеет неправильный формат.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-163">The argument is invalid or missing or has an incorrect format.</span></span> |
| <span data-ttu-id="0cdfa-164">GeneralException</span><span class="sxs-lookup"><span data-stu-id="0cdfa-164">GeneralException</span></span>      | <span data-ttu-id="0cdfa-165">При обработке запроса возникла внутренняя ошибка.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-165">There was an internal error while processing the request.</span></span> |
| <span data-ttu-id="0cdfa-166">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="0cdfa-166">NotImplemented</span></span>        | <span data-ttu-id="0cdfa-167">Запрашиваемая функция не реализована.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-167">The requested feature isn't implemented.</span></span>  |
| <span data-ttu-id="0cdfa-168">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="0cdfa-168">UnsupportedOperation</span></span>  | <span data-ttu-id="0cdfa-169">Выполняемая операция не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-169">The operation being attempted is not supported.</span></span> |
| <span data-ttu-id="0cdfa-170">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="0cdfa-170">AccessDenied</span></span>          | <span data-ttu-id="0cdfa-171">Вы не можете выполнить запрашиваемую операцию.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-171">You cannot perform the requested operation.</span></span> |
| <span data-ttu-id="0cdfa-172">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="0cdfa-172">ItemNotFound</span></span>          | <span data-ttu-id="0cdfa-173">Запрашиваемый ресурс не существует.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-173">The requested resource doesn't exist.</span></span> |

## <a name="get-started"></a><span data-ttu-id="0cdfa-174">Начало работы</span><span class="sxs-lookup"><span data-stu-id="0cdfa-174">Get started</span></span>

<span data-ttu-id="0cdfa-175">Для начала просмотрите пример в этом разделе.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-175">You can use the example in this section to get started.</span></span> <span data-ttu-id="0cdfa-176">В этом примере показано, как программно отобразить текст выбранной фигуры в документе Visio.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-176">This example shows you how to programmatically display the shape text of the selected shape in a Visio diagram.</span></span> <span data-ttu-id="0cdfa-177">Сперва создайте классическую страницу в SharePoint Online или отредактируйте существующую страницу.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-177">To begin, create a classic page in SharePoint Online or edit an existing page.</span></span> <span data-ttu-id="0cdfa-178">Добавьте веб-часть редактора сценариев на эту страницу, а затем скопируйте и вставьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-178">Add a script editor webpart on the page and copy-paste the following code.</span></span>

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        var page = context.document.getActivePage();
        var shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(var i=0; i<shapes.items.length;i++) {
                var shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

<span data-ttu-id="0cdfa-179">После этого вам нужен только URL-адрес документа Visio, с которым вы хотите работать.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-179">After that, all you need is the URL of a Visio diagram that you want to work with.</span></span> <span data-ttu-id="0cdfa-180">Просто отправьте документ Visio в SharePoint Online и откройте его в Visio в Интернете.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-180">Just upload the Visio diagram to SharePoint Online and open it in Visio on the web.</span></span> <span data-ttu-id="0cdfa-181">Оттуда откройте диалоговое окно внедрения и используйте URL-адрес внедрения в приведенном выше примере.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-181">From there, open the Embed dialog and use the Embed URL in the above example.</span></span>

![Копирование URL-адреса файла Visio из диалогового окна внедрения](../images/Visio-embed-url.png)

<span data-ttu-id="0cdfa-183">Если вы используете Visio в Интернете в режиме правки, откройте диалоговое окно внедрения, выбрав **Файл** > **Общий доступ** > **Внедрить**.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-183">If you are using Visio on the web in Edit mode, open the Embed dialog by choosing **File** > **Share** > **Embed**.</span></span> <span data-ttu-id="0cdfa-184">Если вы используете Visio в Интернете в режиме просмотра, откройте диалоговое окно внедрения, выбрав элемент "..." а затем — команду **Внедрить**.</span><span class="sxs-lookup"><span data-stu-id="0cdfa-184">If you are using Visio on the web in View mode, open the Embed dialog by choosing '...' and then **Embed**.</span></span>

## <a name="visio-javascript-api-reference"></a><span data-ttu-id="0cdfa-185">Справочник по API JavaScript для Visio</span><span class="sxs-lookup"><span data-stu-id="0cdfa-185">Visio JavaScript API reference</span></span>

<span data-ttu-id="0cdfa-186">Дополнительные сведения об API JavaScript для Visio см. в [справочной документации по API JavaScript для Visio](/javascript/api/visio).</span><span class="sxs-lookup"><span data-stu-id="0cdfa-186">For detailed information about Visio JavaScript API, see the [Visio JavaScript API reference documentation](/javascript/api/visio).</span></span>

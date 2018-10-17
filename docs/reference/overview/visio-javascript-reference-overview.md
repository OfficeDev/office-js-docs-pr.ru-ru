# <a name="visio-javascript-api-overview"></a><span data-ttu-id="19ac5-101">Обзор API JavaScript для Visio</span><span class="sxs-lookup"><span data-stu-id="19ac5-101">Word-specific JavaScript API overview</span></span>

<span data-ttu-id="19ac5-102">С помощью API JavaScript для Visio вы можете внедрять схемы Visio в SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="19ac5-102">You can use the Visio JavaScript APIs to embed Visio diagrams in SharePoint Online.</span></span> <span data-ttu-id="19ac5-103">Внедренный документ Visio — схема, которая хранится в библиотеке документов SharePoint и отображается на странице SharePoint.</span><span class="sxs-lookup"><span data-stu-id="19ac5-103">An embedded Visio diagram is a diagram that is stored in a SharePoint document library and displayed on a SharePoint page.</span></span> <span data-ttu-id="19ac5-104">Чтобы внедрить документ Visio, отобразите его в элементе `<iframe>` HTML.</span><span class="sxs-lookup"><span data-stu-id="19ac5-104">To embed a Visio diagram, display it in an HTML `<iframe>`iframe element.</span></span> <span data-ttu-id="19ac5-105">После этого вы сможете программным способом работать с внедренным документом при помощи API JavaScript для Visio.</span><span class="sxs-lookup"><span data-stu-id="19ac5-105">Then you can use Visio JavaScript APIs to programmatically work with the embedded diagram.</span></span>

![Документ Visio в iframe на странице SharePoint вместе с веб-частью редактора сценариев](../images/visio-api-block-diagram.png)


<span data-ttu-id="19ac5-107">API JavaScript для Visio позволяет следующее:</span><span class="sxs-lookup"><span data-stu-id="19ac5-107">You can use the Visio JavaScript APIs to:</span></span>

* <span data-ttu-id="19ac5-108">взаимодействовать с элементами документа Visio как со страницами, так и фигурами;</span><span class="sxs-lookup"><span data-stu-id="19ac5-108">Interact with Visio diagram elements like pages and shapes</span></span>
* <span data-ttu-id="19ac5-109">создавать визуальную разметку на холсте документа Visio;</span><span class="sxs-lookup"><span data-stu-id="19ac5-109">Create visual markup on the Visio diagram canvas</span></span>
* <span data-ttu-id="19ac5-110">создавать специальные обработчики событий мыши для документа;</span><span class="sxs-lookup"><span data-stu-id="19ac5-110">Write custom handlers for mouse events within the drawing</span></span>
* <span data-ttu-id="19ac5-111">предоставлять своему решению данные документа, такие как текст фигуры, данные фигуры и гиперссылки.</span><span class="sxs-lookup"><span data-stu-id="19ac5-111">Expose diagram data, such as shape text, shape data, and hyperlinks, to your solution.</span></span>

<span data-ttu-id="19ac5-p102">В этой статье описано, как использовать API JavaScript для Visio с Visio Online, чтобы создавать решения для SharePoint Online. В ней представлены ключевые элементы, понимание роли которых крайне важно при использовании API, такие как прокси-объекты JavaScript, **EmbeddedSession**, **RequestContext**, а также методы **sync()**, **Visio.run()** и **load()**. В приведенных ниже примерах кода показано применение этих элементов.</span><span class="sxs-lookup"><span data-stu-id="19ac5-p102">This article describes how to use the Visio JavaScript APIs with Visio Online to build your solutions for SharePoint Online. It introduces key concepts that are fundamental to using the APIs, such as **EmbeddedSession**, **RequestContext**, and JavaScript proxy objects, and the **sync()**, **Visio.run()**, and **load()** methods. The code examples show you how to apply these concepts.</span></span>

## <a name="embeddedsession"></a><span data-ttu-id="19ac5-115">EmbeddedSession</span><span class="sxs-lookup"><span data-stu-id="19ac5-115">EmbeddedSession</span></span>

<span data-ttu-id="19ac5-116">Объект EmbeddedSession инициализирует взаимодействие между фреймом разработчика и фреймом Visio Online.</span><span class="sxs-lookup"><span data-stu-id="19ac5-116">The EmbeddedSession object initializes communication between the developer frame and the Visio Online frame.</span></span>

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a><span data-ttu-id="19ac5-117">Visio.run(session, function(context) { batch })</span><span class="sxs-lookup"><span data-stu-id="19ac5-117">Visio.run(session, function(context) { batch })</span></span>

<span data-ttu-id="19ac5-118">Метод **Visio.run()** выполняет пакетный сценарий, совершающий действия с объектной моделью Visio.</span><span class="sxs-lookup"><span data-stu-id="19ac5-118">**Visio.run()** executes a batch script that performs actions on the Visio object model.</span></span> <span data-ttu-id="19ac5-119">Пакетные команды включают определения локальных прокси-объектов JavaScript и методов **sync()**, синхронизирующих состояние объектов Visio и локальных объектов, а также разрешение обещания.</span><span class="sxs-lookup"><span data-stu-id="19ac5-119">The batch commands include definitions of local JavaScript proxy objects and **sync()** methods that synchronize the state between local and Visio objects and promise resolution.</span></span> <span data-ttu-id="19ac5-120">Преимущество пакетной обработки запросов в методе **Visio.run()** состоит в том, что при разрешении обещания все отслеживаемые объекты страницы, выделенные во время выполнения, автоматически освобождаются.</span><span class="sxs-lookup"><span data-stu-id="19ac5-120">The advantage of batching requests in **Visio.run()** is that when the promise is resolved, any tracked page objects that were allocated during the execution will be automatically released.</span></span>

<span data-ttu-id="19ac5-121">Метод run принимает объект session и RequestContext и возвращает обещание (обычно это результат **context.sync()**).</span><span class="sxs-lookup"><span data-stu-id="19ac5-121">The run method takes in RequestContext and returns a promise (typically, just the result of **ctx.sync()**).</span></span> <span data-ttu-id="19ac5-122">Пакетную операцию можно выполнить, не указывая ее в методе **Visio.run()**.</span><span class="sxs-lookup"><span data-stu-id="19ac5-122">It is possible to run the batch operation outside of the **Visio.run()**.</span></span> <span data-ttu-id="19ac5-123">Однако в этом случае все ссылки на объекты страницы требуют отслеживания и управления вручную.</span><span class="sxs-lookup"><span data-stu-id="19ac5-123">However, in such a scenario, any page object references needs to be manually tracked and managed.</span></span>

## <a name="requestcontext"></a><span data-ttu-id="19ac5-124">RequestContext</span><span class="sxs-lookup"><span data-stu-id="19ac5-124">RequestContext</span></span>

<span data-ttu-id="19ac5-125">Объект RequestContext облегчает запросы на приложение Visio.</span><span class="sxs-lookup"><span data-stu-id="19ac5-125">Request Context: The RequestContext object facilitates requests to the Excel application.</span></span> <span data-ttu-id="19ac5-126">Поскольку фрейм разработчика и приложение Visio Online выполняются в двух разных iframe, объект RequestContext (контекст в следующем примере) требуется для доступа к Visio и связанным с ним объектам, таким как страницы и фигуры, из фрейма разработчика.</span><span class="sxs-lookup"><span data-stu-id="19ac5-126">The RequestContext object facilitates requests to the Visio application. Because the developer frame and the Visio Online application run in two different iframes, request context is required to get access to Visio and related objects such as pages and shapes, from the developer frame. The following example shows how to create a request context.</span></span>

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

## <a name="proxy-objects"></a><span data-ttu-id="19ac5-127">Прокси-объекты</span><span class="sxs-lookup"><span data-stu-id="19ac5-127">Proxy objects</span></span>

<span data-ttu-id="19ac5-p106">Объекты JavaScript для Visio, объявленные и использованные в надстройке, — это прокси-объекты для реальных объектов в документе Visio. Все действия над прокси-объектами не реализуются в Visio, а состояние документа Visio — в прокси-объектах, пока оно не будет синхронизировано. Состояние документа синхронизируется при выполнении `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="19ac5-p106">The Visio JavaScript objects declared and used in an add-in are proxy objects for the real objects in a Visio document. All actions taken on proxy objects are not realized in Visio, and the state of the Visio document is not realized in the proxy objects until the document state has been synchronized. The document state is synchronized when `context.sync()` is run.</span></span>

<span data-ttu-id="19ac5-131">Например, локальный объект JavaScript getActivePage объявлен в качестве ссылки на выбранный диапазон.</span><span class="sxs-lookup"><span data-stu-id="19ac5-131">For example, the local JavaScript object  is declared to reference the selected range.</span></span> <span data-ttu-id="19ac5-132">Это можно использовать для постановки в очередь настройки его свойств и вызова методов.</span><span class="sxs-lookup"><span data-stu-id="19ac5-132">This can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="19ac5-133">Действия над такими объектами не реализуются до выполнения метода **sync()**.</span><span class="sxs-lookup"><span data-stu-id="19ac5-133">The actions on such objects are not realized until the sync() method is run.</span></span>

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a><span data-ttu-id="19ac5-134">sync()</span><span class="sxs-lookup"><span data-stu-id="19ac5-134">sync()</span></span>

<span data-ttu-id="19ac5-135">Метод **sync()** синхронизирует состояние прокси-объектов JavaScript и реальных объектов в Visio путем выполнения поставленных в очередь инструкций над контекстом и получения свойств загруженных объектов Office для их использования в коде.</span><span class="sxs-lookup"><span data-stu-id="19ac5-135">The **sync()** method, available on the request context, synchronizes the state between JavaScript proxy objects and real objects in Visio by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.</span></span> <span data-ttu-id="19ac5-136">Этот метод возвращает обещание, которое выполняется после завершения синхронизации.</span><span class="sxs-lookup"><span data-stu-id="19ac5-136">This method returns a promise, which is resolved when synchronization is complete.</span></span> 

## <a name="load"></a><span data-ttu-id="19ac5-137">load()</span><span class="sxs-lookup"><span data-stu-id="19ac5-137">load()</span></span>

<span data-ttu-id="19ac5-p109">Метод **load()** используется для заполнения прокси-объектов, созданных на уровне JavaScript надстройки. При попытке получения объекта, такого как документ, сначала на уровне JavaScript создается локальный прокси-объект. Такой объект можно использовать для добавления в очередь настройки его свойств и вызова методов. Но для чтения свойств или связей объекта сначала необходимо вызвать методы **load()** и **sync()**. Метод load() использует свойства и связи, которые требуется загрузить при вызове метода **sync()**.</span><span class="sxs-lookup"><span data-stu-id="19ac5-p109">The **load()** method is used to fill in the proxy objects created in the add-in JavaScript layer. When trying to retrieve an object such as a document, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue the setting of its properties and invoking methods. However, for reading object properties or relations, the **load()** and **sync()** methods need to be invoked first. The load() method takes in the properties and relations that need to be loaded when the **sync()** method is called.</span></span>

<span data-ttu-id="19ac5-143">Ниже представлен синтаксис метода **load()**.</span><span class="sxs-lookup"><span data-stu-id="19ac5-143">The following shows the syntax for the **load()** method.</span></span>

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. <span data-ttu-id="19ac5-144">**properties** - это список имен свойств, которые должны быть загружены, заданные как строки с разделителями-запятыми или массив имен.</span><span class="sxs-lookup"><span data-stu-id="19ac5-144">**properties** is the list of properties and/or relationship names to be loaded, specified as comma-delimited strings or array of names.</span></span> <span data-ttu-id="19ac5-145">Дополнительные сведения см. в описаниях методов **.load()** под каждым объектом.</span><span class="sxs-lookup"><span data-stu-id="19ac5-145">See **.load()** methods under each object for details.</span></span>

2. <span data-ttu-id="19ac5-p111">**loadOption** указывает объект, описывающий свойства select, expand, top и skip. Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](/javascript/api/office/officeextension.loadoption).</span><span class="sxs-lookup"><span data-stu-id="19ac5-p111">**loadOption** specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

## <a name="example-printing-all-shapes-text-in-active-page"></a><span data-ttu-id="19ac5-148">Пример: Печать текста всех фигур на активной странице</span><span class="sxs-lookup"><span data-stu-id="19ac5-148">Example: Printing all shapes text in active page</span></span>

<span data-ttu-id="19ac5-149">Приведенный ниже пример показывает, как распечатать значение текста фигуры из объекта фигур массива.</span><span class="sxs-lookup"><span data-stu-id="19ac5-149">The following example shows you how to print shape text value from an array shapes object.</span></span>
<span data-ttu-id="19ac5-150">Метод **Visio.run()** содержит пакет инструкций.</span><span class="sxs-lookup"><span data-stu-id="19ac5-150">The **Visio.run()** method contains a batch of instructions.</span></span> <span data-ttu-id="19ac5-151">В рамках этого пакета создается прокси-объект, который ссылается на фигуры в активном документе.</span><span class="sxs-lookup"><span data-stu-id="19ac5-151">As part of this batch, a proxy object is created that references shapes on the active document.</span></span>

<span data-ttu-id="19ac5-152">Все эти команды ставятся в очередь и выполняются при вызове метода **ctx.sync()**.</span><span class="sxs-lookup"><span data-stu-id="19ac5-152">All these commands are queued and run when **ctx.sync()** is called.</span></span> <span data-ttu-id="19ac5-153">Метод **sync()** возвращает обещание, с помощью которого его можно связать с другими операциями.</span><span class="sxs-lookup"><span data-stu-id="19ac5-153">The **sync()** method returns a promise that can be used to chain it with other operations.</span></span>

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

## <a name="error-messages"></a><span data-ttu-id="19ac5-154">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="19ac5-154">Error messages</span></span>

<span data-ttu-id="19ac5-p114">Ошибки возвращаются с помощью объекта ошибки, состоящего из кода и сообщения. В таблице ниже перечислены возможные ошибки.</span><span class="sxs-lookup"><span data-stu-id="19ac5-p114">Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.</span></span>

| <span data-ttu-id="19ac5-157">error.code</span><span class="sxs-lookup"><span data-stu-id="19ac5-157">error.code</span></span>            | <span data-ttu-id="19ac5-158">error.message</span><span class="sxs-lookup"><span data-stu-id="19ac5-158">error.message</span></span> |
|-----------------------|----------------------------------------------------------------|
| <span data-ttu-id="19ac5-159">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="19ac5-159">InvalidArgument</span></span>       | <span data-ttu-id="19ac5-160">Аргумент недопустим, отсутствует или имеет неправильный формат.</span><span class="sxs-lookup"><span data-stu-id="19ac5-160">The argument is invalid or missing or has an incorrect format.</span></span> |
| <span data-ttu-id="19ac5-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="19ac5-161">GeneralException</span></span>      | <span data-ttu-id="19ac5-162">При обработке запроса возникла внутренняя ошибка.</span><span class="sxs-lookup"><span data-stu-id="19ac5-162">There was an internal error while processing the request.</span></span> |
| <span data-ttu-id="19ac5-163">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="19ac5-163">NotImplemented</span></span>        | <span data-ttu-id="19ac5-164">Запрашиваемая функция не реализована.</span><span class="sxs-lookup"><span data-stu-id="19ac5-164">The requested feature isn't implemented.</span></span>  |
| <span data-ttu-id="19ac5-165">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="19ac5-165">UnsupportedOperation</span></span>  | <span data-ttu-id="19ac5-166">Выполняемая операция не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="19ac5-166">The operation being attempted is not supported.</span></span> |
| <span data-ttu-id="19ac5-167">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="19ac5-167">AccessDenied</span></span>          | <span data-ttu-id="19ac5-168">Вы не можете выполнить запрашиваемую операцию.</span><span class="sxs-lookup"><span data-stu-id="19ac5-168">You cannot perform the requested operation.</span></span> |
| <span data-ttu-id="19ac5-169">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="19ac5-169">ItemNotFound</span></span>          | <span data-ttu-id="19ac5-170">Запрашиваемый ресурс не существует.</span><span class="sxs-lookup"><span data-stu-id="19ac5-170">The requested resource doesn't exist.</span></span> |

## <a name="get-started"></a><span data-ttu-id="19ac5-171">Начало работы</span><span class="sxs-lookup"><span data-stu-id="19ac5-171">Get started</span></span>

<span data-ttu-id="19ac5-172">Пример в этом разделе можно использовать для начала работы.</span><span class="sxs-lookup"><span data-stu-id="19ac5-172">You can use the example in this section to get started.</span></span> <span data-ttu-id="19ac5-173">В этом примере показано, как программно отобразить текст выбранной фигуры в схеме Visio.</span><span class="sxs-lookup"><span data-stu-id="19ac5-173">This example shows you how to programmatically display the shape text of the selected shape in a Visio diagram.</span></span> <span data-ttu-id="19ac5-174">Чтобы приступить к работе, создайте классическую страницу в SharePoint Online или отредактируйте существующую страницу.</span><span class="sxs-lookup"><span data-stu-id="19ac5-174">To begin, create a classic page in SharePoint Online or edit an existing page.</span></span> <span data-ttu-id="19ac5-175">Добавьте веб-часть редактора сценариев на странице и скопируйте и вставьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="19ac5-175">Add a script editor webpart on the page and copy-paste the following code.</span></span>

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

<span data-ttu-id="19ac5-176">После этого все, что требуется — это URL-адрес схемы Visio, с которой вы хотите работать.</span><span class="sxs-lookup"><span data-stu-id="19ac5-176">After that, all you need is the URL of a Visio diagram that you want to work with.</span></span> <span data-ttu-id="19ac5-177">Просто загрузите схему Visio в SharePoint Online и откройте ее в Visio Online.</span><span class="sxs-lookup"><span data-stu-id="19ac5-177">Just upload the Visio diagram to SharePoint Online and open it in Visio Online.</span></span> <span data-ttu-id="19ac5-178">Оттуда откройте диалоговое окно Внедрить и используйте URL-адрес Внедрить в приведенном выше примере.</span><span class="sxs-lookup"><span data-stu-id="19ac5-178">From there, open the Embed dialog and use the Embed URL in the above example.</span></span>

![Скопируйте URL-адрес файла Visio из диалога Внедрить](../images/Visio-embed-url.png)

<span data-ttu-id="19ac5-180">Если вы используете Visio Online в Режиме правки, откройте диалоговое окно Внедрить, выбрав **Файл** > **Поделиться** > **Внедрить**.</span><span class="sxs-lookup"><span data-stu-id="19ac5-180">If you are using Visio Online in Edit mode, open the Embed dialog by choosing **File** > **Share** > **Embed**.</span></span> <span data-ttu-id="19ac5-181">Если вы используете Visio Online в режиме просмотра, откройте диалоговое окно Внедрить, выбрав '... а затем **Внедрить**.</span><span class="sxs-lookup"><span data-stu-id="19ac5-181">If you are using Visio Online in View mode, open the Embed dialog by choosing '...' and then **Embed**.</span></span>

## <a name="open-api-specifications"></a><span data-ttu-id="19ac5-182">Открытые спецификации API</span><span class="sxs-lookup"><span data-stu-id="19ac5-182">Open API specifications</span></span>

<span data-ttu-id="19ac5-p118">Мы публикуем новые API на странице [Открытые спецификации API](../openspec.md), чтобы вы могли делиться своим мнением о них. Узнайте, над какими функциями мы работаем, и поделитесь своим мнением о спецификациях.</span><span class="sxs-lookup"><span data-stu-id="19ac5-p118">As we design and develop new APIs, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.</span></span>

## <a name="visio-javascript-api-reference"></a><span data-ttu-id="19ac5-185">Ссылка на API JavaScript для Visio</span><span class="sxs-lookup"><span data-stu-id="19ac5-185">Visio JavaScript APIs reference</span></span>

<span data-ttu-id="19ac5-186">Для получения подробной информации об API JavaScript для Visio см. Справочную документацию [Visio JavaScript API](/javascript/api/visio).</span><span class="sxs-lookup"><span data-stu-id="19ac5-186">For detailed information about Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/visio).</span></span>

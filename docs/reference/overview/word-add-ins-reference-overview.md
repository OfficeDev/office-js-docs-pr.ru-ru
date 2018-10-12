# <a name="word-javascript-api-overview"></a><span data-ttu-id="2fe55-101">Обзор API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="2fe55-101">Word JavaScript API usage overview</span></span>

<span data-ttu-id="2fe55-p101">Word предоставляет большой набор API. Вы можете использовать эти API для создания надстроек, взаимодействующих с контентом и метаданными документов. С помощью этих API вы сможете создавать привлекательные приложения, интегрируемые с Word и расширяющие возможности этой программы. Вы можете импортировать и экспортировать контент, собирать новые документы на основе различных источников данных, выполнять интеграцию с рабочими процессами документов и создавать пользовательские решения для работы с документами.</span><span class="sxs-lookup"><span data-stu-id="2fe55-p101">Word provides a rich set of APIs that you can use to create add-ins that interact with document content and metadata. Use these APIs to create compelling experiences that integrate with and extend Word. You can import and export content, assemble new documents from different data sources, and integrate with document workflows to create custom document solutions.</span></span>

<span data-ttu-id="2fe55-105">Для взаимодействия с объектами и метаданными в документе Word вы можете использовать два указанных ниже API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2fe55-105">You can use two JavaScript APIs to interact with the objects and metadata in a Word document:</span></span>

- <span data-ttu-id="2fe55-106">API JavaScript для Word: впервые появился в Office 2016.</span><span class="sxs-lookup"><span data-stu-id="2fe55-106">Word JavaScript API - Introduced in Office 2016.</span></span>
- <span data-ttu-id="2fe55-107">[API JavaScript для Office](../javascript-api-for-office.md) (Office.js): впервые появился в Office 2013.</span><span class="sxs-lookup"><span data-stu-id="2fe55-107">[JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Introduced in Office 2013.</span></span>

## <a name="word-javascript-api"></a><span data-ttu-id="2fe55-108">API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="2fe55-108">Word JavaScript API</span></span>

<span data-ttu-id="2fe55-p102">API JavaScript для Word загружается с помощью файла Office.js. Этот API изменяет способ взаимодействия с объектами, например с документами и абзацами. Вместо набора отдельных асинхронных API для получения и обновления каждого из этих объектов новый API JavaScript для Word предоставляет прокси-объекты JavaScript, которые соответствуют реальным объектам, выполняемым в Word. Вы можете напрямую взаимодействовать с этими прокси-объектами, синхронно считывая и записывая их свойства, а также вызывая синхронные методы для операций над ними. Эти взаимодействия с прокси-объектами не сразу реализуются в выполняющихся сценариях. Метод **context.sync** синхронизирует состояние запущенного JavaScript и реальных объектов в Office, выполняя поставленные в очередь инструкции и получая свойства загруженных объектов Word для их использования в сценарии.</span><span class="sxs-lookup"><span data-stu-id="2fe55-p102">The Word JavaScript API is loaded by Office.js. The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word. You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them. These interactions with proxy objects aren’t immediately realized in the running script. The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

## <a name="javascript-api-for-office"></a><span data-ttu-id="2fe55-115">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="2fe55-115">JavaScript API for Office</span></span>

<span data-ttu-id="2fe55-116">Файл Office.js можно получить из следующих расположений:</span><span class="sxs-lookup"><span data-stu-id="2fe55-116">You can reference Office.js from the following locations:</span></span>

* <span data-ttu-id="2fe55-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js — Используйте этот ресурс для надстроек производства.</span><span class="sxs-lookup"><span data-stu-id="2fe55-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js - use this resource for production add-ins.</span></span>
* <span data-ttu-id="2fe55-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js — Используйте этот ресурс при работе функций предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="2fe55-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - use this resource when you're trying out preview features.</span></span>

<span data-ttu-id="2fe55-p103">Если вы используете [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), чтобы получить шаблоны проектов, включающие файл Office.js, вы можете скачать [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx).  Кроме того, [чтобы получить файл Office.js, вы можете воспользоваться NuGet](https://www.nuget.org/packages/Microsoft.Office.js/).</span><span class="sxs-lookup"><span data-stu-id="2fe55-p103">If you're using [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), you can download the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) to get project templates that include Office.js.  You can also use [nuget to get Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).</span></span>

<span data-ttu-id="2fe55-121">Если вы используете TypeScript и у вас есть npm, то вы можете получить определения TypeScript, выполнив в интерфейсе командной строки следующую команду: `typings install office-js --ambient`.</span><span class="sxs-lookup"><span data-stu-id="2fe55-121">If you use TypeScript and have npm, you can get the the TypeScript definitions by typing this in your command line interface: `typings install office-js --ambient`.</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="2fe55-122">Запуск надстроек Word</span><span class="sxs-lookup"><span data-stu-id="2fe55-122">Running Word add-ins</span></span>

<span data-ttu-id="2fe55-p104">Чтобы запустить надстройку, воспользуйтесь обработчиком событий Office.initialize. Дополнительные сведения об инициализации надстроек см. в статье [Общие сведения об API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="2fe55-p104">To run your add-in, use an Office.initialize event handler. For more information about add-in initialization, see [Understanding the API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) .</span></span>

<span data-ttu-id="2fe55-125">Надстройки, предназначенных для Word 2016 или более поздней версии выполните, передав функции в метод **Word.run()** .</span><span class="sxs-lookup"><span data-stu-id="2fe55-125">Add-ins that target Word 2016 or later execute by passing a function into the **Word.run()** method.</span></span> <span data-ttu-id="2fe55-126">Функция, переданная в метод **run**, должна содержать аргумент контекста.</span><span class="sxs-lookup"><span data-stu-id="2fe55-126">The function passed into the **run** method must have a context argument.</span></span> <span data-ttu-id="2fe55-127">  Этот объект [context](/javascript/api/word/word.requestcontext) отличается от объекта context, который вы получаете из объекта Office, но также используется для взаимодействия со средой выполнения Word.</span><span class="sxs-lookup"><span data-stu-id="2fe55-127">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="2fe55-128">Объект context предоставляет доступ к объектной модели API JavaScript для Word.</span><span class="sxs-lookup"><span data-stu-id="2fe55-128">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="2fe55-129">В следующем примере показано, как инициализировать и выполнить надстройку Word с помощью метода **Word.run()** .</span><span class="sxs-lookup"><span data-stu-id="2fe55-129">The following example shows how to initialize and execute a Word add-in by using the **Word.run()** method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="2fe55-130">Синхронизация документов Word с помощью прокси-объектов API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="2fe55-130">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="2fe55-p106">Объектная модель API JavaScript для Word нестрого связана с объектами в Word. Объекты API JavaScript для Word представляют собой прокси-объекты для объектов в документе Word. Действия, выполняемые над прокси-объектами, не будут реализованы в Word, пока не будет синхронизировано состояние документа. И наоборот, состояние документа Word не будет реализовано в прокси-объектах, пока оно не будет синхронизировано. Чтобы синхронизировать состояние документа, выполните метод **context.sync()**. В примере ниже показано, как создать прокси-объект основного текста и помещенную в очередь команду для загрузки свойства текста в прокси-объекте основного текста и как использовать метод **context.sync()** для синхронизации основного текста документа Word с прокси-объектом основного текста.</span><span class="sxs-lookup"><span data-stu-id="2fe55-p106">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the **context.sync()** method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the **context.sync()** method to synchronize the body of the Word document with the body proxy object.</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="2fe55-137">Выполнение пакета команд</span><span class="sxs-lookup"><span data-stu-id="2fe55-137">Executing a batch of commands</span></span>

<span data-ttu-id="2fe55-p107">У прокси-объектов Word есть методы для доступа к объектной модели и ее обновления. Эти методы выполняются последовательно в том порядке, в котором они были поставлены в очередь в пакете. При вызове метода context.sync() выполняются все команды, помещенные в очередь в пакете.</span><span class="sxs-lookup"><span data-stu-id="2fe55-p107">The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. All of the commands that are queued in the batch are executed when context.sync() is called.</span></span>

<span data-ttu-id="2fe55-p108">В примере ниже показано, как работает очередь команд. При вызове метода **context.sync()** в Word выполняется команда загрузки основного текста. Затем выполняется команда вставки текста в основной текст в Word. Результаты выполнения команд возвращаются в прокси-объект основного текста. Значение свойства **body.text** в API JavaScript для Word представляет собой значение основного текста документа Word <u>перед тем, как</u> текст был вставлен в документ Word.</span><span class="sxs-lookup"><span data-stu-id="2fe55-p108">The following example shows how the command queue works. When **context.sync()** is called, the command to load the body text is executed in Word. Then, the command to insert text into the body in Word occurs. The results are then returned to the body proxy object. The value of the **body.text** property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="word-javascript-api-open-specifications"></a><span data-ttu-id="2fe55-146">Открытые спецификации API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="2fe55-146">Word JavaScript API open specifications</span></span>

<span data-ttu-id="2fe55-p109">Мы публикуем новые API для надстроек Word на странице [Открытые спецификации API](../openspec.md), чтобы вы могли делиться своим мнением. Узнайте, над какими функциями API JavaScript для Word мы работаем, и поделитесь своим мнением о проектируемых спецификациях.</span><span class="sxs-lookup"><span data-stu-id="2fe55-p109">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="word-javascript-api-reference"></a><span data-ttu-id="2fe55-149">Ссылка на API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="2fe55-149">Word JavaScript API reference</span></span>

<span data-ttu-id="2fe55-150">Дополнительные сведения об об интерфейсе API JavaScript для Word см. в  [Справочная документация по  API JavaScript для Word](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="2fe55-150">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="see-also"></a><span data-ttu-id="2fe55-151">См. также</span><span class="sxs-lookup"><span data-stu-id="2fe55-151">See also</span></span>

* [<span data-ttu-id="2fe55-152">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="2fe55-152">Word add-ins overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [<span data-ttu-id="2fe55-153">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="2fe55-153">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [<span data-ttu-id="2fe55-154">Примеры надстроек Word на сайте GitHub</span><span class="sxs-lookup"><span data-stu-id="2fe55-154">Word add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)

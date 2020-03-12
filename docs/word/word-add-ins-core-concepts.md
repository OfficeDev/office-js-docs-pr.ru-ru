---
title: Основные концепции программирования с помощью API JavaScript для Word
description: Создание надстроек для Word с помощью API JavaScript для Word.
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: 319570a7790504bdf95c6a66c07db67ca29dec55
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596769"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="e3f20-103">Основные концепции программирования с помощью API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-103">Fundamental programming concepts with the Word JavaScript API</span></span>

<span data-ttu-id="e3f20-104">В этой статье описаны основные концепции использования [API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) с целью создания надстроек для Word 2016 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="e3f20-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="e3f20-105">Ссылки на Office.js</span><span class="sxs-lookup"><span data-stu-id="e3f20-105">Referencing Office.js</span></span>

<span data-ttu-id="e3f20-106">Файл Office.js можно получить из указанных ниже расположений.</span><span class="sxs-lookup"><span data-stu-id="e3f20-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="e3f20-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`. Используйте этот ресурс для рабочих надстроек.</span><span class="sxs-lookup"><span data-stu-id="e3f20-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="e3f20-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`. Используйте этот ресурс для применения предварительных функций.</span><span class="sxs-lookup"><span data-stu-id="e3f20-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="e3f20-109">Наборы обязательных элементов API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="e3f20-110">Наборы требований — это именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="e3f20-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="e3f20-111">Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API.</span><span class="sxs-lookup"><span data-stu-id="e3f20-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="e3f20-112">Подробнее о наборах обязательных элементов API JavaScript для Word см. в статье [Наборы обязательных элементов API JavaScript для Word](../reference/requirement-sets/word-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e3f20-112">For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="e3f20-113">Запуск надстроек Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-113">Running Word add-ins</span></span>

<span data-ttu-id="e3f20-114">Чтобы запустить надстройку, воспользуйтесь обработчиком событий `Office.initialize`Office.initialize.</span><span class="sxs-lookup"><span data-stu-id="e3f20-114">To run your add-in, use an `Office.initialize` event handler.</span></span> <span data-ttu-id="e3f20-115">Дополнительные сведения об инициализации надстроек см. в статье [Общие сведения об API](../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="e3f20-115">For more information about add-in initialization, see [Understanding the API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

<span data-ttu-id="e3f20-116">Надстройки для Word 2016 и более поздних версий запускаются передачей функции в метод `Word.run()`.</span><span class="sxs-lookup"><span data-stu-id="e3f20-116">Add-ins that target Word 2016 or later run by passing a function into the `Word.run()` method.</span></span> <span data-ttu-id="e3f20-117">Функции, передаваемой в метод `run`, обязательно должен быть присвоен контекстный аргумент.</span><span class="sxs-lookup"><span data-stu-id="e3f20-117">The function passed into the `run` method must have a context argument.</span></span> <span data-ttu-id="e3f20-118">Этот [контекстный объект](/javascript/api/word/word.requestcontext) отличается от контекстного объекта, получаемого из объекта Office, но он также используется для взаимодействия со средой выполнения Word.</span><span class="sxs-lookup"><span data-stu-id="e3f20-118">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="e3f20-119">Контекстный объект предоставляет доступ к объектной модели API JavaScript для Word.</span><span class="sxs-lookup"><span data-stu-id="e3f20-119">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="e3f20-120">В следующем примере показано, как инициализировать и запустить надстройку Word с помощью метода `Word.run()`.</span><span class="sxs-lookup"><span data-stu-id="e3f20-120">The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.</span></span>

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

### <a name="asynchronous-nature-of-word-apis"></a><span data-ttu-id="e3f20-121">Асинхронный характер API Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-121">Asynchronous nature of Word APIs</span></span>

<span data-ttu-id="e3f20-122">API JavaScript для Word загружается с помощью файла Office.js.</span><span class="sxs-lookup"><span data-stu-id="e3f20-122">The Word JavaScript API is loaded by Office.js.</span></span> <span data-ttu-id="e3f20-123">Этот API изменяет способ взаимодействия с объектами, например с документами и абзацами.</span><span class="sxs-lookup"><span data-stu-id="e3f20-123">The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs.</span></span> <span data-ttu-id="e3f20-124">Вместо набора отдельных асинхронных API для получения и обновления каждого из этих объектов новый API JavaScript для Word предоставляет прокси-объекты JavaScript, которые соответствуют действующим объектам, используемым в Word.</span><span class="sxs-lookup"><span data-stu-id="e3f20-124">Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides "proxy" JavaScript objects that correspond to the live objects running in Word.</span></span> <span data-ttu-id="e3f20-125">Вы можете взаимодействовать с этими прокси-объектами, синхронно считывая и записывая их свойства, а также вызывая синхронные методы для операций над ними.</span><span class="sxs-lookup"><span data-stu-id="e3f20-125">You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them.</span></span> <span data-ttu-id="e3f20-126">Эти взаимодействия с прокси-объектами не сразу реализуются в выполняющихся сценариях.</span><span class="sxs-lookup"><span data-stu-id="e3f20-126">These interactions with proxy objects aren't immediately realized in the running script.</span></span> <span data-ttu-id="e3f20-127">Метод `context.sync` синхронизирует состояние запущенного JavaScript и реальных объектов в Office, выполняя поставленные в очередь инструкции и получая свойства загруженных объектов Word для их использования в сценарии.</span><span class="sxs-lookup"><span data-stu-id="e3f20-127">The `context.sync` method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="e3f20-128">Синхронизация документов Word с помощью прокси-объектов API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-128">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="e3f20-p105">Объектная модель API JavaScript для Word нестрого связана с объектами в Word. Объекты API JavaScript для Word представляют собой прокси-объекты для объектов в документе Word. Действия, выполняемые над прокси-объектами, не будут реализованы в Word, пока не будет синхронизировано состояние документа. И наоборот, состояние документа Word не будет реализовано в прокси-объектах, пока оно не будет синхронизировано. Чтобы синхронизировать состояние документа, выполните метод `context.sync()`. В примере ниже показано, как создать прокси-объект основного текста и помещенную в очередь команду для загрузки свойства текста в прокси-объекте основного текста и как использовать метод `context.sync()` для синхронизации основного текста документа Word с прокси-объектом основного текста.</span><span class="sxs-lookup"><span data-stu-id="e3f20-p105">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the `context.sync()` method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the `context.sync()` method to synchronize the body of the Word document with the body proxy object.</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    body.load("text");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="e3f20-135">Выполнение пакета команд</span><span class="sxs-lookup"><span data-stu-id="e3f20-135">Executing a batch of commands</span></span>

<span data-ttu-id="e3f20-136">У прокси-объектов Word есть методы для доступа и обновления объектной модели.</span><span class="sxs-lookup"><span data-stu-id="e3f20-136">The Word proxy objects have methods for accessing and updating the object model.</span></span> <span data-ttu-id="e3f20-137">Эти методы выполняются последовательно в том порядке, в каком они были добавлены в пакет.</span><span class="sxs-lookup"><span data-stu-id="e3f20-137">These methods are run sequentially in the order in which they were queued in the batch.</span></span> <span data-ttu-id="e3f20-138">При вызове метода `context.sync()` выполняются все команды, помещенные в очередь в пакете.</span><span class="sxs-lookup"><span data-stu-id="e3f20-138">All of the commands that are queued in the batch are run when `context.sync()` is called.</span></span>

<span data-ttu-id="e3f20-139">В следующем примере показано, как действует очередь команд.</span><span class="sxs-lookup"><span data-stu-id="e3f20-139">The following example shows how the command queue works.</span></span> <span data-ttu-id="e3f20-140">При вызове метода `context.sync()` в Word выполняется команда загрузки основного текста.</span><span class="sxs-lookup"><span data-stu-id="e3f20-140">When `context.sync()` is called, the command to load the body text is run in Word.</span></span> <span data-ttu-id="e3f20-141">Затем выполняется команда вставки текста в основной текст в Word.</span><span class="sxs-lookup"><span data-stu-id="e3f20-141">Then, the command to insert text into the body in Word occurs.</span></span> <span data-ttu-id="e3f20-142">Результаты возвращаются в прокси-объект.основного текста.</span><span class="sxs-lookup"><span data-stu-id="e3f20-142">The results are then returned to the body proxy object.</span></span> <span data-ttu-id="e3f20-143">Значение свойства `body.text` в API JavaScript для Word представляет собой значение основного текста документа Word <u>перед тем, как</u> текст был вставлен в документ Word.</span><span class="sxs-lookup"><span data-stu-id="e3f20-143">The value of the `body.text` property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    body.load("text");

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

## <a name="see-also"></a><span data-ttu-id="e3f20-144">См. также</span><span class="sxs-lookup"><span data-stu-id="e3f20-144">See also</span></span>

- [<span data-ttu-id="e3f20-145">Обзор API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-145">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="e3f20-146">Создание первой надстройки Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-146">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="e3f20-147">Руководство по надстройкам Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-147">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="e3f20-148">Справочник по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="e3f20-148">Word JavaScript API reference</span></span>](/javascript/api/word)
---
title: Использование модели API для определенных приложений
description: Сведения о модели API на основе обещаний для надстроек Excel, OneNote и Word.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: fb25201174dcd97b40ccf6be69b238951103db07
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408602"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="e9e87-103">Использование модели API для определенных приложений</span><span class="sxs-lookup"><span data-stu-id="e9e87-103">Using the application-specific API model</span></span>

<span data-ttu-id="e9e87-104">В этой статье описано, как использовать модель API для создания надстроек в Excel, Word и OneNote.</span><span class="sxs-lookup"><span data-stu-id="e9e87-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="e9e87-105">Здесь представлены основные понятия, лежащие в основе использования API на основе обещаний.</span><span class="sxs-lookup"><span data-stu-id="e9e87-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="e9e87-106">Эта модель не поддерживается клиентами Office 2013.</span><span class="sxs-lookup"><span data-stu-id="e9e87-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="e9e87-107">Используйте [общую модель API](office-javascript-api-object-model.md) для работы с этими версиями Office.</span><span class="sxs-lookup"><span data-stu-id="e9e87-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="e9e87-108">Полные сведения о доступности платформ см. в статье [Доступность клиентских приложений и платформ Office для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="e9e87-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="e9e87-109">В примерах на этой странице используются API JavaScript для Excel, но эти понятия также относятся к API JavaScript для OneNote, Visio и Word.</span><span class="sxs-lookup"><span data-stu-id="e9e87-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="e9e87-110">Асинхронный характер API на основе обещаний</span><span class="sxs-lookup"><span data-stu-id="e9e87-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="e9e87-111">Надстройки Office — это веб-сайты, отображающиеся внутри контейнера браузера в приложениях Office, таких как Excel.</span><span class="sxs-lookup"><span data-stu-id="e9e87-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="e9e87-112">Этот контейнер внедряется в приложение Office на платформах для классических ПК, например Office для Windows, и запускается в элементе iFrame HTML в Office для Интернета.</span><span class="sxs-lookup"><span data-stu-id="e9e87-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="e9e87-113">Из-за соображений производительности интерфейсы API Office.js не могут синхронно взаимодействовать с приложениями Office на всех платформах.</span><span class="sxs-lookup"><span data-stu-id="e9e87-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="e9e87-114">Таким образом, вызов API `sync()` в Office.js возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается, когда приложение Office выполняет запрошенные действия чтения или записи.</span><span class="sxs-lookup"><span data-stu-id="e9e87-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="e9e87-115">Кроме того, вы можете поместить в очередь несколько действий, например действия настройки свойств или вызова методов, а затем запустить их в виде пакета команд в одном вызове метода `sync()`, а не отправлять отдельные запросы для каждого действия.</span><span class="sxs-lookup"><span data-stu-id="e9e87-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="e9e87-116">В разделах ниже описано, как сделать это, используя API `run()` и `sync()`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="e9e87-117">Функция \*.run</span><span class="sxs-lookup"><span data-stu-id="e9e87-117">\*.run function</span></span>

<span data-ttu-id="e9e87-118">`Excel.run`, `Word.run` и `OneNote.run` исполняют функцию, определяющую действия, выполняемые в Excel, Word и OneNote.</span><span class="sxs-lookup"><span data-stu-id="e9e87-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="e9e87-119">`*.run` автоматически создает контекст запроса, который можно использовать для взаимодействия с объектами Office.</span><span class="sxs-lookup"><span data-stu-id="e9e87-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="e9e87-120">Когда `*.run` завершает работу, обещание разрешается, и все объекты, которые были выделены в среде выполнения, будут автоматически разблокированы.</span><span class="sxs-lookup"><span data-stu-id="e9e87-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="e9e87-121">В следующем примере показано, как использовать шаблон `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="e9e87-122">Такой же шаблон используется в Word и OneNote.</span><span class="sxs-lookup"><span data-stu-id="e9e87-122">The same pattern is also used with Word and OneNote.</span></span>

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a><span data-ttu-id="e9e87-123">Контекст запроса</span><span class="sxs-lookup"><span data-stu-id="e9e87-123">Request context</span></span>

<span data-ttu-id="e9e87-124">Приложение Office и ваша надстройка работают в двух разных процессах.</span><span class="sxs-lookup"><span data-stu-id="e9e87-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="e9e87-125">Так как они используют разные среды выполнения, надстройкам требуется объект `RequestContext`, чтобы можно было подключать надстройку к объектам в Office, например к листам, диапазонам, абзацам и таблицам.</span><span class="sxs-lookup"><span data-stu-id="e9e87-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="e9e87-126">Этот объект `RequestContext` предоставляется в качестве аргумента при вызове `*.run`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="e9e87-127">Прокси-объекты</span><span class="sxs-lookup"><span data-stu-id="e9e87-127">Proxy objects</span></span>

<span data-ttu-id="e9e87-128">Объекты JavaScript для Office, объявляемые и используемые с помощью API на основе обещаний, являются прокси-объектами. </span><span class="sxs-lookup"><span data-stu-id="e9e87-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="e9e87-129">Все методы, которые вы вызываете, либо свойства, которые вы настраиваете либо загружаете, в прокси-объектах просто добавляются в очередь команд, ожидающих выполнения.</span><span class="sxs-lookup"><span data-stu-id="e9e87-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="e9e87-130">Когда вы вызываете метод `sync()` в контексте запроса (например, `context.sync()`), команды, помещенные в очередь, передаются в приложение Office и выполняются.</span><span class="sxs-lookup"><span data-stu-id="e9e87-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="e9e87-131">По существу, эти API ориентированы на работу с пакетами.</span><span class="sxs-lookup"><span data-stu-id="e9e87-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="e9e87-132">Вы можете поместить в очередь любое количество изменений в контексте запроса, а затем вызвать метод `sync()`, чтобы запустить пакет команд, помещенных в очередь.</span><span class="sxs-lookup"><span data-stu-id="e9e87-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="e9e87-133">Например, во фрагменте кода ниже показано, как объявить локальный объект JavaScript [Excel.Range](/javascript/api/excel/excel.range) (`selectedRange`) для ссылки на выделенный диапазон в книге Excel, а затем задать ряд свойств для этого объекта.</span><span class="sxs-lookup"><span data-stu-id="e9e87-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="e9e87-134">Объект `selectedRange` представляет собой прокси-объект, поэтому свойства, заданные в этом объекте, и метод, вызываемый в этом объекте, не будут отображены в документе Excel, пока надстройка не вызовет метод `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="e9e87-135">Совет по производительности: минимизируйте количество созданных прокси-объектов</span><span class="sxs-lookup"><span data-stu-id="e9e87-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="e9e87-136">Избегайте повторного создания одного и того же прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="e9e87-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="e9e87-137">Вместо этого, если вам нужен одинаковый прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.</span><span class="sxs-lookup"><span data-stu-id="e9e87-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

### <a name="sync"></a><span data-ttu-id="e9e87-138">sync()</span><span class="sxs-lookup"><span data-stu-id="e9e87-138">sync()</span></span>

<span data-ttu-id="e9e87-139">При вызове метода `sync()` в контексте запроса будет синхронизировано состояние прокси-объектов и объектов в документе Office.</span><span class="sxs-lookup"><span data-stu-id="e9e87-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="e9e87-140">Метод `sync()` запускает любые команды, помещенные в очередь в контексте запроса, и получает значения для любых свойств, которые следует загрузить в прокси-объектах.</span><span class="sxs-lookup"><span data-stu-id="e9e87-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="e9e87-141">Метод `sync()` выполняется асинхронно и возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается по завершении работы метода `sync()`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="e9e87-142">В примере ниже показана пакетная функция, которая определяет локальный прокси-объект JavaScript (`selectedRange`), загружает свойство этого объекта, а затем использует шаблон обещаний JavaScript для вызова метода `context.sync()` и, соответственно, синхронизации состояния прокси-объектов и объектов в документе Excel.</span><span class="sxs-lookup"><span data-stu-id="e9e87-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

<span data-ttu-id="e9e87-143">В предыдущем примере настроен параметр `selectedRange`, и его свойство `address` загружается при вызове `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="e9e87-144">Так как `sync()` — это асинхронная операция, всегда следует возвращать объект `Promise`, чтобы завершить операцию `sync()`, прежде чем продолжить выполнение сценария.</span><span class="sxs-lookup"><span data-stu-id="e9e87-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="e9e87-145">Если вы используете TypeScript или JavaScript ES6+, вы можете `await` вызов `context.sync()` вместо возврата обещания.</span><span class="sxs-lookup"><span data-stu-id="e9e87-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="e9e87-146">Совет по производительности: минимизируйте количество вызовов синхронизации</span><span class="sxs-lookup"><span data-stu-id="e9e87-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="e9e87-147">В API JavaScript для Excel `sync()` является единственной асинхронной операцией и в некоторых обстоятельствах может выполняться медленно, особенно в случае с Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="e9e87-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="e9e87-148">Для оптимизации производительности минимизируйте количество вызовов `sync()`, поставив в очередь максимально возможное количество изменений до ее вызова.</span><span class="sxs-lookup"><span data-stu-id="e9e87-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="e9e87-149">Дополнительные сведения об оптимизации производительности с помощью `sync()` см. в статье [Избегайте использования метода context.sync в циклах](../concepts/correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="e9e87-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="e9e87-150">load()</span><span class="sxs-lookup"><span data-stu-id="e9e87-150">load()</span></span>

<span data-ttu-id="e9e87-151">Чтобы можно было считывать свойства прокси-объекта, вам необходимо явно загрузить их и заполнить прокси-объект данными из документа Office, а затем вызвать метод `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="e9e87-152">Например, вы создали прокси-объект для ссылки на выделенный диапазон, а затем вам потребовалось считать свойство `address` выделенного диапазона. Прежде чем вы сможете считать свойство `address`, вам потребуется загрузить его.</span><span class="sxs-lookup"><span data-stu-id="e9e87-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="e9e87-153">Чтобы запросить загрузку свойств прокси-объекта, вызовите метод `load()` в объекте и укажите свойства, которые необходимо загрузить.</span><span class="sxs-lookup"><span data-stu-id="e9e87-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="e9e87-154">В следующем примере показана загрузка свойства `Range.address` для `myRange`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> <span data-ttu-id="e9e87-155">Если вы вызываете методы или задаете свойства только в прокси-объекте, вам не нужно вызывать метод `load()`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="e9e87-156">Метод `load()` требуется только тогда, когда вам необходимо считать свойства в прокси-объекте.</span><span class="sxs-lookup"><span data-stu-id="e9e87-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="e9e87-p115">Аналогично запросам для задания свойств или вызова методов в прокси-объектах, запросы на загрузку свойств в прокси-объектах добавляются в очередь команд, ожидающих выполнения, в контексте запроса, который будет запущен, когда вы в следующий раз вызовете метод `sync()`. В очередь можно поставить сколько угодно вызовов `load()` в контексте запроса.</span><span class="sxs-lookup"><span data-stu-id="e9e87-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="e9e87-159">Скалярные и навигационные свойства</span><span class="sxs-lookup"><span data-stu-id="e9e87-159">Scalar and navigation properties</span></span>

<span data-ttu-id="e9e87-160">Существует две категории свойств: **скалярные** и **навигационные**.</span><span class="sxs-lookup"><span data-stu-id="e9e87-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="e9e87-161">К скалярным свойствам относятся назначаемые типы, такие как строки, целые числа и структуры JSON.</span><span class="sxs-lookup"><span data-stu-id="e9e87-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="e9e87-162">Свойства навигации — это объекты и коллекции объектов только для чтения, которым назначаются поля вместо прямого назначения свойства.</span><span class="sxs-lookup"><span data-stu-id="e9e87-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="e9e87-163">Например, элементы `name` и `position` объекта [Excel.Worksheet](/javascript/api/excel/excel.worksheet) являются скалярными свойствами, а `protection` и `tables` — свойствами навигации.</span><span class="sxs-lookup"><span data-stu-id="e9e87-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="e9e87-164">Надстройка может использовать свойства навигации в качестве пути для загрузки определенных скалярных свойств.</span><span class="sxs-lookup"><span data-stu-id="e9e87-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="e9e87-165">Следующий код помещает в очередь команду `load` для имени шрифта, используемого объектом `Excel.Range`, без загрузки каких-либо других сведений.</span><span class="sxs-lookup"><span data-stu-id="e9e87-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="e9e87-166">Вы также можете задавать скалярные свойства из свойства навигации по пути к ним.</span><span class="sxs-lookup"><span data-stu-id="e9e87-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="e9e87-167">Например, вы можете задать размер шрифта для `Excel.Range` с помощью команды `someRange.format.font.size = 10;`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="e9e87-168">Чтобы задать свойство, необязательно загружать его.</span><span class="sxs-lookup"><span data-stu-id="e9e87-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="e9e87-169">Имейте в виду, что некоторые свойства объекта могут совпадать с именем другого объекта.</span><span class="sxs-lookup"><span data-stu-id="e9e87-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="e9e87-170">Например, `format` — это свойство объекта `Excel.Range`, но также имеется и объект `format`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="e9e87-171">Поэтому если вы, например, вызываете `range.load("format")`, это эквивалентно `range.format.load()` (нежелательный пустой оператор `load()`).</span><span class="sxs-lookup"><span data-stu-id="e9e87-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="e9e87-172">Чтобы избежать этого, ваш код должен загружать только "конечные узлы" в дереве объектов.</span><span class="sxs-lookup"><span data-stu-id="e9e87-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="e9e87-173">Вызов метода `load` без параметров (не рекомендуется)</span><span class="sxs-lookup"><span data-stu-id="e9e87-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="e9e87-174">Если вызвать метод `load()` для объекта (или коллекции), не указывая параметры, будут загружены все скалярные свойства объекта или объектов в коллекции.</span><span class="sxs-lookup"><span data-stu-id="e9e87-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="e9e87-175">Загрузка ненужных данных замедлит вашу надстройку.</span><span class="sxs-lookup"><span data-stu-id="e9e87-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="e9e87-176">Необходимо всегда явным образом указывать свойства для загрузки.</span><span class="sxs-lookup"><span data-stu-id="e9e87-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e9e87-177">Объем данных, возвращаемых оператором `load` без параметров, может превышать ограничения по размерам для службы.</span><span class="sxs-lookup"><span data-stu-id="e9e87-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="e9e87-178">Чтобы сократить риски для старых надстроек, некоторые свойства не возвращаются методом `load` без их явного запроса.</span><span class="sxs-lookup"><span data-stu-id="e9e87-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="e9e87-179">Следующие свойства исключены из таких операций загрузки:</span><span class="sxs-lookup"><span data-stu-id="e9e87-179">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="e9e87-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="e9e87-180">ClientResult</span></span>

<span data-ttu-id="e9e87-181">Методы в API на основе обещаний, возвращающие примитивные типы, используют шаблон, похожий на парадигму `load`/`sync`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="e9e87-182">Например, `Excel.TableCollection.getCount` получает количество таблиц в коллекции.</span><span class="sxs-lookup"><span data-stu-id="e9e87-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="e9e87-183">`getCount` возвращает `ClientResult<number>`. Это означает, что свойство `value` возвращаемого [`ClientResult`](/javascript/api/office/officeextension.clientresult) выражено числом.</span><span class="sxs-lookup"><span data-stu-id="e9e87-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="e9e87-184">Сценарий не может получить доступ к этому значению, пока не вызовет `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="e9e87-185">Следующий код получает общее количество таблиц в книге Excel и записывает его в консоль.</span><span class="sxs-lookup"><span data-stu-id="e9e87-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a><span data-ttu-id="e9e87-186">set()</span><span class="sxs-lookup"><span data-stu-id="e9e87-186">set()</span></span>

<span data-ttu-id="e9e87-187">Установка свойств объекта с вложенными свойствами навигации может быть трудоемкой задачей.</span><span class="sxs-lookup"><span data-stu-id="e9e87-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="e9e87-188">Вместо того чтобы задавать отдельные свойства с помощью путей навигации, как описано выше, вы можете использовать метод `object.set()`, доступный для объектов в API JavaScript на основе обещаний.</span><span class="sxs-lookup"><span data-stu-id="e9e87-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="e9e87-189">С помощью этого метода можно задать сразу несколько свойств объекта, передавая другой объект того же типа Office.js или объект JavaScript со свойствами, сходными по структуре со свойствами объекта, для которого вызывается метод.</span><span class="sxs-lookup"><span data-stu-id="e9e87-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="e9e87-p124">В приведенном ниже примере кода показано, как задать несколько свойств формата диапазона, вызвав метод `set()` и передав в него объект JavaScript, имена и типы свойств которого повторяют структуру свойств объекта `Range`. В этом примере предполагается, что данные находятся в диапазоне **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="e9e87-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="e9e87-192">Некоторые свойства невозможно задать напрямую</span><span class="sxs-lookup"><span data-stu-id="e9e87-192">Some properties cannot be set directly</span></span>

<span data-ttu-id="e9e87-193">Некоторые свойства невозможно задать, хотя они и поддерживают запись.</span><span class="sxs-lookup"><span data-stu-id="e9e87-193">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="e9e87-194">Эти свойства являются частью родительского свойства, которое должно быть задано как один объект.</span><span class="sxs-lookup"><span data-stu-id="e9e87-194">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="e9e87-195">Это связано с тем, что родительское свойство использует вложенные свойства с определенными логическими связями.</span><span class="sxs-lookup"><span data-stu-id="e9e87-195">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="e9e87-196">Эти родительские свойства должны быть заданы с помощью нотации литерала объекта, чтобы задать весь объект, а не отдельные вложенные свойства этого объекта. </span><span class="sxs-lookup"><span data-stu-id="e9e87-196">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="e9e87-197">Один из примеров доступен в разделе [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="e9e87-197">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="e9e87-198">Свойство `zoom` должно быть задано с помощью одного объекта [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions), как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="e9e87-198">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="e9e87-199">В предыдущем примере вы ***не*** сможете напрямую присвоить значение `zoom`: `sheet.pageLayout.zoom.scale = 200;`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-199">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="e9e87-200">Этот оператор выдает ошибку, так как `zoom` не загружен.</span><span class="sxs-lookup"><span data-stu-id="e9e87-200">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="e9e87-201">Даже если `zoom` загружен, масштабный набор не будет работать.</span><span class="sxs-lookup"><span data-stu-id="e9e87-201">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="e9e87-202">Все контекстные операции происходят в `zoom`, обновляя прокси-объект в надстройке и переписывая локально установленные значения.</span><span class="sxs-lookup"><span data-stu-id="e9e87-202">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="e9e87-203">Это поведение отличается от [свойств навигации](application-specific-api-model.md#scalar-and-navigation-properties), например [Range.format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="e9e87-203">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="e9e87-204">Свойства `format` можно задать с помощью навигации по объектам, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="e9e87-204">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="e9e87-205">Вы можете определить свойство, для которого невозможно напрямую задать его вложенные свойства, путем проверки модификатора только для чтения.</span><span class="sxs-lookup"><span data-stu-id="e9e87-205">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="e9e87-206">Для всех свойств, доступных только для чтения, можно напрямую задать их вложенные свойства, использующиеся не только для чтения.</span><span class="sxs-lookup"><span data-stu-id="e9e87-206">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="e9e87-207">Записываемые свойства, например `PageLayout.zoom`, должны быть заданы на уровне объекта.</span><span class="sxs-lookup"><span data-stu-id="e9e87-207">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="e9e87-208">Сводка:</span><span class="sxs-lookup"><span data-stu-id="e9e87-208">In summary:</span></span>

- <span data-ttu-id="e9e87-209">Свойство только для чтения: вложенные свойства можно задать с помощью навигации.</span><span class="sxs-lookup"><span data-stu-id="e9e87-209">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="e9e87-210">Записываемое свойство: вложенные свойства нельзя задать с помощью навигации (необходимо установить их в рамках начального назначения родительского объекта).</span><span class="sxs-lookup"><span data-stu-id="e9e87-210">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>



## <a name="42ornullobject-methods-and-properties"></a><span data-ttu-id="e9e87-211">Методы и свойства &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="e9e87-211">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="e9e87-212">Некоторые методы и свойства доступа создают исключение, если нужный объект не существует.</span><span class="sxs-lookup"><span data-stu-id="e9e87-212">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="e9e87-213">Например, если для получения листа Excel указать имя листа, не существующее в книге, метод `getItem()` создаст исключение `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-213">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="e9e87-214">Библиотеки конкретных приложений позволяют коду проверять наличие сущностей документа, не требуя кода обработки исключений. </span><span class="sxs-lookup"><span data-stu-id="e9e87-214">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="e9e87-215">Это достигается с помощью вариантов методов и свойств `*OrNullObject`. </span><span class="sxs-lookup"><span data-stu-id="e9e87-215">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="e9e87-216">Эти варианты вместо создания исключения возвращают объект, свойству `isNullObject` которого присвоено значение `true`, если указанный элемент не существует.</span><span class="sxs-lookup"><span data-stu-id="e9e87-216">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="e9e87-217">Например, вы можете вызвать метод `getItemOrNullObject()` для коллекции, такой как **Worksheets**, чтобы получить элемент из коллекции.</span><span class="sxs-lookup"><span data-stu-id="e9e87-217">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="e9e87-218">Метод `getItemOrNullObject()` возвращает указанный элемент, если он существует. В противном случае возвращается объект, свойству `isNullObject` которого присвоено значение `true`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-218">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="e9e87-219">Затем код может оценить это свойство, чтобы определить, существует ли объект.</span><span class="sxs-lookup"><span data-stu-id="e9e87-219">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="e9e87-220">Варианты `*OrNullObject` никогда не возвращают значение JavaScript `null`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-220">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="e9e87-221">Они возвращают обычные прокси-объекты Office.</span><span class="sxs-lookup"><span data-stu-id="e9e87-221">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="e9e87-222">Если сущность, представляемая объектом, не существует, свойству `isNullObject` объекта присваивается значение `true`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-222">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="e9e87-223">Не проверяйте возвращенный объект на нулевое значение или ложность.</span><span class="sxs-lookup"><span data-stu-id="e9e87-223">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="e9e87-224">Для него никогда не используются значения `null`, `false` или `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-224">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="e9e87-225">В следующем примере кода осуществляется попытка получить лист Excel с именем Data с помощью метода `getItemOrNullObject()`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-225">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="e9e87-226">Если лист с таким именем не существует, создается новый лист.</span><span class="sxs-lookup"><span data-stu-id="e9e87-226">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="e9e87-227">Обратите внимание, что код не загружает свойство `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-227">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="e9e87-228">Office автоматически загружает это свойство, когда вызывается `context.sync`, поэтому вам не нужно явным образом загружать его с помощью `datasheet.load('isNullObject')`.</span><span class="sxs-lookup"><span data-stu-id="e9e87-228">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a><span data-ttu-id="e9e87-229">См. также</span><span class="sxs-lookup"><span data-stu-id="e9e87-229">See also</span></span>

* [<span data-ttu-id="e9e87-230">Общая объектная модель API JavaScript</span><span class="sxs-lookup"><span data-stu-id="e9e87-230">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* [<span data-ttu-id="e9e87-231">Ограничения ресурсов и оптимизация производительности надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e9e87-231">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)

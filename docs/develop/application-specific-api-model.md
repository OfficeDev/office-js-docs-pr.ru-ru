---
title: Использование модели API для определенных приложений
description: Сведения о модели API на основе обещаний для Excel, OneNote и надстроек Word.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: fb25201174dcd97b40ccf6be69b238951103db07
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408602"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="8cd4b-103">Использование модели API для определенных приложений</span><span class="sxs-lookup"><span data-stu-id="8cd4b-103">Using the application-specific API model</span></span>

<span data-ttu-id="8cd4b-104">В этой статье описывается, как использовать модель API для создания надстроек в Excel, Word и OneNote.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="8cd4b-105">В нем представлены основные концепции использования API на основе Promise.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="8cd4b-106">Эта модель не поддерживается клиентами Office 2013.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="8cd4b-107">Используйте [общую модель API](office-javascript-api-object-model.md) для работы с этими версиями Office.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="8cd4b-108">Чтобы ознакомиться с полными сведениями о доступности платформы, ознакомьтесь с разделом [клиентские приложения и платформы Office для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="8cd4b-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="8cd4b-109">В примерах на этой странице используются API JavaScript для Excel, но эти понятия также относятся к API-интерфейсам OneNote, Visio и Word JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="8cd4b-110">Асинхронная природа интерфейсов API на основе обещаний</span><span class="sxs-lookup"><span data-stu-id="8cd4b-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="8cd4b-111">Надстройки Office — это веб-сайты, которые отображаются внутри контейнера браузера в приложениях Office, таких как Excel.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="8cd4b-112">Этот контейнер внедряется в приложение Office на платформах на настольных компьютерах, таких как Office в Windows, и запускается в элементе iFrame HTML в Office в Интернете.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="8cd4b-113">Из-за соображений производительности интерфейсы API Office.js не могут синхронно взаимодействовать с приложениями Office на всех платформах.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="8cd4b-114">Таким образом, `sync()` вызов API в Office.js возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) , которое разрешается при выполнении приложением Office запрошенных действий чтения или записи.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="8cd4b-115">Кроме того, можно поставить в очередь несколько действий, таких как установка свойств или вызов методов, и запускать их как пакет команд с одним вызовом `sync()` , а не отправлять отдельный запрос для каждого действия.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="8cd4b-116">В следующих разделах описано, как это сделать с помощью `run()` `sync()` API-интерфейсов.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="8cd4b-117">функция \*. Run</span><span class="sxs-lookup"><span data-stu-id="8cd4b-117">\*.run function</span></span>

<span data-ttu-id="8cd4b-118">`Excel.run`, `Word.run` и `OneNote.run` выполните функцию, которая определяет действия, выполняемые с помощью Excel, Word и OneNote.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="8cd4b-119">`*.run` автоматически создает контекст запроса, который можно использовать для взаимодействия с объектами Office.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="8cd4b-120">По `*.run` завершении обещание разрешается, и все объекты, которые были выделены во время выполнения, автоматически освобождаются.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="8cd4b-121">В приведенном ниже примере показано, как использовать `Excel.run` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="8cd4b-122">Такой же шаблон также используется с Word и OneNote.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-122">The same pattern is also used with Word and OneNote.</span></span>

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

## <a name="request-context"></a><span data-ttu-id="8cd4b-123">Контекст запроса</span><span class="sxs-lookup"><span data-stu-id="8cd4b-123">Request context</span></span>

<span data-ttu-id="8cd4b-124">Приложение Office и надстройка запускаются в двух различных процессах.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="8cd4b-125">Так как они используют разные среды выполнения, надстройкам требуется `RequestContext` объект, чтобы подключить надстройку к объектам в Office, таким как листы, диапазоны, абзацы и таблицы.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="8cd4b-126">Этот `RequestContext` объект предоставляется в качестве аргумента при вызове `*.run` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="8cd4b-127">Прокси-объекты</span><span class="sxs-lookup"><span data-stu-id="8cd4b-127">Proxy objects</span></span>

<span data-ttu-id="8cd4b-128">Объекты JavaScript для Office, объявляемые и используемые с помощью API на основе Promise, являются прокси-объектами.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="8cd4b-129">Все методы, которые вы вызываете, либо свойства, которые вы настраиваете либо загружаете, в прокси-объектах просто добавляются в очередь команд, ожидающих выполнения.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="8cd4b-130">При вызове `sync()` метода в контексте запроса (например, `context.sync()` ) команды, поставленные в очередь, отправляются в приложение Office и запускаются.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="8cd4b-131">Эти API основаны на пакетной основе.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="8cd4b-132">Вы можете поместить в очередь любое количество изменений, которое требуется в контексте запроса, а затем вызвать `sync()` метод для запуска пакета команд в очереди.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="8cd4b-133">Например, в приведенном ниже фрагменте кода объявляется локальный объект JavaScript [Excel. Range](/javascript/api/excel/excel.range) , `selectedRange` для ссылки на выбранный диапазон в книге Excel, а затем задаются некоторые свойства этого объекта.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="8cd4b-134">`selectedRange`Объект является прокси-объектом, поэтому заданные свойства и метод, вызываемый для этого объекта, не будут отражены в документе Excel до вызова надстройки `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="8cd4b-135">Совет по производительности: Минимизируйте число созданных прокси-объектов</span><span class="sxs-lookup"><span data-stu-id="8cd4b-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="8cd4b-136">Избегайте повторного создания одного и того же прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="8cd4b-137">Вместо этого, если вам нужен одинаковый прокси-объект для нескольких операций, создайте его один раз и назначьте его переменной, а затем используйте эту переменную в своем коде.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

### <a name="sync"></a><span data-ttu-id="8cd4b-138">sync()</span><span class="sxs-lookup"><span data-stu-id="8cd4b-138">sync()</span></span>

<span data-ttu-id="8cd4b-139">При вызове `sync()` метода в контексте запроса выполняется синхронизация состояния между объектами прокси-сервера и объектами в документе Office.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="8cd4b-140">`sync()`Метод выполняет все команды, помещенные в очередь в контексте запроса, и получает значения для всех свойств, которые должны быть загружены в прокси-объекты.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="8cd4b-141">`sync()`Метод выполняется асинхронно и возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается по `sync()` завершении метода.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="8cd4b-142">В следующем примере показана Пакетная функция, которая определяет локальный прокси-сервер JavaScript ( `selectedRange` ), загружает свойство этого объекта, а затем использует шаблон JavaScript для синхронизации для `context.sync()` синхронизации состояния между прокси-объектами и объектами в документе Excel.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

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

<span data-ttu-id="8cd4b-143">В предыдущем примере `selectedRange` установлен, и его параметр `address` загружается при вызове `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="8cd4b-144">Так как `sync()` это асинхронная операция, всегда следует возвращать `Promise` объект, чтобы убедиться, что `sync()` операция завершается, прежде чем продолжить выполнение скрипта.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="8cd4b-145">Если вы используете TypeScript или ES6 + JavaScript, вы можете `await` `context.sync()` позвонить вместо возврата обещаний.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="8cd4b-146">Совет по производительности: Минимизируйте число вызовов синхронизации</span><span class="sxs-lookup"><span data-stu-id="8cd4b-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="8cd4b-147">В API JavaScript для Excel `sync()` является единственной асинхронной операцией и в некоторых обстоятельствах может выполняться медленно, особенно в случае с Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="8cd4b-148">Для оптимизации производительности минимизируйте количество вызовов `sync()`, поставив в очередь максимально возможное количество изменений до ее вызова.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="8cd4b-149">Чтобы получить дополнительные сведения о оптимизации производительности с помощью `sync()` , [не используйте метод Context. Sync в циклах](../concepts/correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="8cd4b-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="8cd4b-150">load()</span><span class="sxs-lookup"><span data-stu-id="8cd4b-150">load()</span></span>

<span data-ttu-id="8cd4b-151">Перед чтением свойств прокси-объекта необходимо явно загрузить свойства для заполнения прокси-объекта данными из документа Office и затем вызвать метод `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="8cd4b-152">Например, если вы создаете прокси-объект для ссылки на выбранный диапазон, а затем хотите прочитать свойство выбранного диапазона, необходимо `address` загрузить `address` свойство, прежде чем его можно будет прочитать.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="8cd4b-153">Чтобы запросить свойства прокси-объекта, вызовите `load()` метод для объекта и укажите свойства для загрузки.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="8cd4b-154">В следующем примере показано `Range.address` свойство, для которого выполняется загрузка `myRange` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

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
> <span data-ttu-id="8cd4b-155">Если вы вызываете только методы или задаете свойства прокси-объекта, вам не нужно вызывать `load()` метод.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="8cd4b-156">`load()`Метод требуется только в том случае, если необходимо прочитать свойства прокси-объекта.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="8cd4b-p115">Аналогично запросам для задания свойств или вызова методов в прокси-объектах, запросы на загрузку свойств в прокси-объектах добавляются в очередь команд, ожидающих выполнения, в контексте запроса, который будет запущен, когда вы в следующий раз вызовете метод `sync()`. В очередь можно поставить сколько угодно вызовов `load()` в контексте запроса.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="8cd4b-159">Скалярные и навигационные свойства</span><span class="sxs-lookup"><span data-stu-id="8cd4b-159">Scalar and navigation properties</span></span>

<span data-ttu-id="8cd4b-160">Существует две категории свойств: **скалярные** и **навигационные**.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="8cd4b-161">К скалярным свойствам относятся назначаемые типы, такие как строки, целые числа и структуры JSON.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="8cd4b-162">Свойства навигации — это объекты, доступные только для чтения, и коллекции объектов, которым назначены поля, а не непосредственное назначение свойства.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="8cd4b-163">Например, `name` `position` элементы в объекте [Excel. лист](/javascript/api/excel/excel.worksheet) являются скалярными свойствами, в то время как `protection` `tables` Свойства навигации.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="8cd4b-164">Надстройка может использовать свойства навигации в качестве пути для загрузки определенных скалярных свойств.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="8cd4b-165">Приведенный ниже код ставит в очередь `load` команду для имени шрифта `Excel.Range` , используемого объектом, без загрузки каких бы то ни было других сведений.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="8cd4b-166">Кроме того, можно задать скалярные свойства свойства навигации, обходим путь.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="8cd4b-167">Например, можно задать размер шрифта для элемента с помощью параметра `Excel.Range` `someRange.format.font.size = 10;` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="8cd4b-168">Вам не нужно загружать свойство перед его заданием.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="8cd4b-169">Обратите внимание, что некоторые свойства объекта могут иметь то же имя, что и другой объект.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="8cd4b-170">Например, `format` является свойством `Excel.Range` объекта, но `format` само по себе также является объектом.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="8cd4b-171">Таким образом, при совершении такого вызова, как `range.load("format")` , это эквивалентно `range.format.load()` (нежелательный пустой `load()` оператор).</span><span class="sxs-lookup"><span data-stu-id="8cd4b-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="8cd4b-172">Чтобы избежать этого, код должен загружать только "конечные узлы" в дереве объектов.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="8cd4b-173">Вызов `load` без параметров (не рекомендуется)</span><span class="sxs-lookup"><span data-stu-id="8cd4b-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="8cd4b-174">При вызове `load()` метода для объекта (или коллекции) без указания каких-либо параметров будут загружены все скалярные свойства объекта или коллекции.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="8cd4b-175">Загрузка ненужных данных приведет к снижению производительности надстройки.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="8cd4b-176">Всегда следует явно указывать свойства для загрузки.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8cd4b-177">Объем данных, возвращаемых оператором `load` без параметров, может превышать ограничения по размерам для службы.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="8cd4b-178">Чтобы сократить риски для старых надстроек, некоторые свойства не возвращаются методом `load` без их явного запроса.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="8cd4b-179">Следующие свойства исключены из таких операций загрузки:</span><span class="sxs-lookup"><span data-stu-id="8cd4b-179">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="8cd4b-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="8cd4b-180">ClientResult</span></span>

<span data-ttu-id="8cd4b-181">Методы в API на основе обещания, возвращающие примитивные типы, имеют похожий шаблон для `load` / `sync` парадигмы.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="8cd4b-182">Например, `Excel.TableCollection.getCount` получает количество таблиц в коллекции.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="8cd4b-183">`getCount` Возвращает значение `ClientResult<number>` , означающее, что `value` возвращаемое свойство [`ClientResult`](/javascript/api/office/officeextension.clientresult) является числом.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="8cd4b-184">Скрипт не может получить доступ к этому значению, пока не вызовет `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="8cd4b-185">Приведенный ниже код получает общее количество таблиц в книге Excel и записывает их в консоль.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

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

### <a name="set"></a><span data-ttu-id="8cd4b-186">set()</span><span class="sxs-lookup"><span data-stu-id="8cd4b-186">set()</span></span>

<span data-ttu-id="8cd4b-187">Установка свойств объекта с вложенными свойствами навигации может быть трудоемкой задачей.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="8cd4b-188">В качестве альтернативы для установки отдельных свойств с помощью путей навигации, описанных выше, можно использовать `object.set()` метод, доступный для объектов в API JavaScript на основе Promise.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="8cd4b-189">С помощью этого метода можно задать сразу несколько свойств объекта, передавая другой объект того же типа Office.js или объект JavaScript со свойствами, сходными по структуре со свойствами объекта, для которого вызывается метод.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="8cd4b-p124">В приведенном ниже примере кода показано, как задать несколько свойств формата диапазона, вызвав метод `set()` и передав в него объект JavaScript, имена и типы свойств которого повторяют структуру свойств объекта `Range`. В этом примере предполагается, что данные находятся в диапазоне **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

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

### <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="8cd4b-192">Некоторые свойства невозможно задать напрямую</span><span class="sxs-lookup"><span data-stu-id="8cd4b-192">Some properties cannot be set directly</span></span>

<span data-ttu-id="8cd4b-193">Некоторые свойства не могут быть заданы, несмотря на то, что они доступны для записи.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-193">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="8cd4b-194">Эти свойства являются частью родительского свойства, которое должно быть задано как один объект.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-194">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="8cd4b-195">Это связано с тем, что родительское свойство использует вложенные свойства с определенными логическими связями.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-195">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="8cd4b-196">Эти родительские свойства должны быть заданы с помощью нотации литерала объекта, чтобы задать весь объект, а не задавать отдельные вложенные свойства этого объекта.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-196">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="8cd4b-197">Один из примеров этого примера находится в файле [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="8cd4b-197">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="8cd4b-198">`zoom`Свойство должно быть задано с помощью одного объекта [пажелайаутзумоптионс](/javascript/api/excel/excel.pagelayoutzoomoptions) , как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="8cd4b-198">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="8cd4b-199">В предыдущем примере вы ***не*** сможете напрямую присвоить `zoom` значение: `sheet.pageLayout.zoom.scale = 200;` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-199">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="8cd4b-200">Этот оператор выдает ошибку, так как `zoom` не загружен.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-200">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="8cd4b-201">Даже если `zoom` были загружены, набор масштабов не вступит в силу.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-201">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="8cd4b-202">Все операции контекста выполняются `zoom` , обновляя прокси-объект в надстройке и перезаписывая локально заданные значения.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-202">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="8cd4b-203">Это поведение отличается от [свойств навигации](application-specific-api-model.md#scalar-and-navigation-properties) , таких как [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="8cd4b-203">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="8cd4b-204">Свойства `format` можно задать с помощью навигации по объектам, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="8cd4b-204">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="8cd4b-205">Можно определить свойство, для которого не могут быть заданы вложенные свойства, путем проверки модификатора только для чтения.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-205">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="8cd4b-206">Все свойства, доступные только для чтения, могут иметь нередактируемые вложенные свойства, не предназначенные только для чтения.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-206">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="8cd4b-207">Записываемые свойства, такие как `PageLayout.zoom` , должны быть заданы на уровне объекта.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-207">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="8cd4b-208">В сводке:</span><span class="sxs-lookup"><span data-stu-id="8cd4b-208">In summary:</span></span>

- <span data-ttu-id="8cd4b-209">Свойство только для чтения: вложенные свойства можно задать с помощью навигации.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-209">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="8cd4b-210">Записываемое свойство: подсвойства невозможно задать с помощью навигации (необходимо задать в качестве части исходного назначения родительского объекта).</span><span class="sxs-lookup"><span data-stu-id="8cd4b-210">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>



## <a name="ornullobject-methods-and-properties"></a><span data-ttu-id="8cd4b-211">&#42;методы и свойства Орнуллобжект</span><span class="sxs-lookup"><span data-stu-id="8cd4b-211">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="8cd4b-212">Некоторые методы и свойства метода доступа создают исключение, если нужный объект не существует.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-212">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="8cd4b-213">Например, если вы попытаетесь получить лист Excel, указав имя листа, которого нет в книге, `getItem()` метод создаст `ItemNotFound` исключение.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-213">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="8cd4b-214">Библиотеки, зависящие от приложения, позволяют коду проверять наличие сущностей документа, не требуя кода обработки исключений.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-214">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="8cd4b-215">Это достигается с помощью `*OrNullObject` вариантов методов и свойств.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-215">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="8cd4b-216">Эти варианты возвращают объект, `isNullObject` свойству которого присвоено значение `true` , если указанный элемент не существует, а не создает исключение.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-216">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="8cd4b-217">Например, вы можете вызвать `getItemOrNullObject()` метод для коллекции, например, с помощью **листов** , чтобы получить элемент из коллекции.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-217">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="8cd4b-218">`getItemOrNullObject()`Метод возвращает указанный элемент, если он существует; в противном случае возвращает объект, `isNullObject` свойству которого присвоено значение `true` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-218">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="8cd4b-219">Затем код может оценить это свойство, чтобы определить, существует ли объект.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-219">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="8cd4b-220">`*OrNullObject`Варианты никогда не возвращают значение JavaScript `null` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-220">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="8cd4b-221">Они возвращают обычные прокси-объекты Office.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-221">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="8cd4b-222">Если объект, который представляет объект, не существует, то `isNullObject` для свойства объекта задано значение `true` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-222">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="8cd4b-223">Не проверяйте возвращаемый объект на значение null или фалсити.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-223">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="8cd4b-224">Он никогда `null` `false` или `undefined` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-224">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="8cd4b-225">Следующий пример кода пытается извлечь лист Excel с именем "Data" с помощью `getItemOrNullObject()` метода.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-225">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="8cd4b-226">Если лист с таким именем не существует, создается новый лист.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-226">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="8cd4b-227">Обратите внимание, что код не загружает `isNullObject` свойство.</span><span class="sxs-lookup"><span data-stu-id="8cd4b-227">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="8cd4b-228">Office автоматически загружает это свойство при `context.sync` его вызове, поэтому нет необходимости явно загружать его с аналогичным действием `datasheet.load('isNullObject')` .</span><span class="sxs-lookup"><span data-stu-id="8cd4b-228">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="8cd4b-229">См. также</span><span class="sxs-lookup"><span data-stu-id="8cd4b-229">See also</span></span>

* [<span data-ttu-id="8cd4b-230">Общая объектная модель API JavaScript</span><span class="sxs-lookup"><span data-stu-id="8cd4b-230">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* [<span data-ttu-id="8cd4b-231">Ограничения ресурсов и оптимизация производительности надстроек Office</span><span class="sxs-lookup"><span data-stu-id="8cd4b-231">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)

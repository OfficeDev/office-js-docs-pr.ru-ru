---
title: Основные концепции программирования с помощью API JavaScript для Excel
description: Создание надстроек для Excel с помощью API JavaScript для Excel.
ms.date: 04/25/2019
localization_priority: Priority
ms.openlocfilehash: 26822d9caa91f4a65a9dbb82f82db989b4409214
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353260"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="e292d-103">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e292d-103">Fundamental programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="e292d-104">В этой статье описано, как создавать надстройки для Excel 2016 или более поздней версии с помощью [API JavaScript для Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="e292d-104">This article describes how to use the [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="e292d-105">В статье изложены основные принципы, которые являются фундаментальными при использовании этого API, а также имеются рекомендации по выполнению определенных задач, например чтению данных из большого диапазона или записи данных в него, изменения всех ячеек в диапазоне и много другого.</span><span class="sxs-lookup"><span data-stu-id="e292d-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="e292d-106">Асинхронный характер API Excel</span><span class="sxs-lookup"><span data-stu-id="e292d-106">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="e292d-p102">Веб-надстройки Excel работают в контейнере браузера, внедренном в приложение Office на платформах для настольных ПК, например Office для Windows, и работающем в iFrame HTML в Office Online. Вам не удастся заставить API Office.js синхронно взаимодействовать с ведущим приложением Excel на всех поддерживаемых платформах из-за соображений производительности. Таким образом, вызов API **sync()** в Office.js возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается, когда приложение Excel выполняет запрошенные действия чтения или записи. Кроме того, вы можете поместить в очередь несколько действий, например действия настройки свойств или вызова методов, а затем запустить их в виде пакета команд в одном вызове метода **sync()**, а не отправлять отдельные запросы для каждого действия. В разделах ниже описано, как сделать это, используя API **Excel.run()** и **sync()**.</span><span class="sxs-lookup"><span data-stu-id="e292d-p102">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action. The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>

## <a name="excelrun"></a><span data-ttu-id="e292d-112">Excel.run</span><span class="sxs-lookup"><span data-stu-id="e292d-112">Excel.run</span></span>

<span data-ttu-id="e292d-p103">**Excel.run** выполняет функцию, в которой вы указываете действия, которые необходимо совершить над объектной моделью Excel. **Excel.run** автоматически создает контекст запроса, который вы можете использовать для взаимодействия с объектами Excel. Когда **Excel.run** завершает работу, обещание разрешается, и все объекты, которые были выделены в среде выполнения, будут автоматически разблокированы.</span><span class="sxs-lookup"><span data-stu-id="e292d-p103">**Excel.run** executes a function where you specify the actions to perform against the Excel object model. **Excel.run** automatically creates a request context that you can use to interact with Excel objects. When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="e292d-p104">В примере ниже показано, как использовать **Excel.run**. Оператор catch перехватывает и записывает ошибки, возникающие в **Excel.run**, в журнал.</span><span class="sxs-lookup"><span data-stu-id="e292d-p104">The following example shows how to use **Excel.run**. The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>

```js
Excel.run(function (context) {
    // You can use the Excel JavaScript API here in the batch function
    // to execute actions on the Excel object model.
    console.log('Your code goes here.');
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="run-options"></a><span data-ttu-id="e292d-118">Параметры выполнения</span><span class="sxs-lookup"><span data-stu-id="e292d-118">Run options</span></span>

<span data-ttu-id="e292d-119">В **Excel.run** есть перегрузка, получающая объект [RunOptions](/javascript/api/excel/excel.runoptions).</span><span class="sxs-lookup"><span data-stu-id="e292d-119">**Excel.run** has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="e292d-120">Он содержит набор свойств, влияющих на поведение платформы при выполнении функции.</span><span class="sxs-lookup"><span data-stu-id="e292d-120">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="e292d-121">Ниже перечислены поддерживаемые в настоящее время свойства.</span><span class="sxs-lookup"><span data-stu-id="e292d-121">The following property is currently supported:</span></span>

- <span data-ttu-id="e292d-122">`delayForCellEdit`: определяет, откладывает ли Excel пакетный запрос до выхода пользователя из режима правки ячейки.</span><span class="sxs-lookup"><span data-stu-id="e292d-122">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="e292d-123">Если присвоено значение **true**, пакетный запрос откладывается и запускается, когда пользователь выходит из режима правки ячейки.</span><span class="sxs-lookup"><span data-stu-id="e292d-123">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="e292d-124">Если присвоено значение **false**, происходит автоматический сбой пакетного запроса, если пользователь находится в режиме правки ячейки (приводит к ошибке обращения к пользователю).</span><span class="sxs-lookup"><span data-stu-id="e292d-124">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="e292d-125">Поведение по умолчанию при отсутствии заданного свойства `delayForCellEdit` аналогично поведению при значении **false**.</span><span class="sxs-lookup"><span data-stu-id="e292d-125">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="request-context"></a><span data-ttu-id="e292d-126">Контекст запроса</span><span class="sxs-lookup"><span data-stu-id="e292d-126">Request context</span></span>

<span data-ttu-id="e292d-p107">Excel и ваша надстройка запускаются как два отдельных процесса. Так как они используют разные среды выполнения, надстройкам Excel требуется объект **RequestContext**, чтобы можно было подключать надстройку к объектам в Excel, например к листам, диапазонам, диаграммам и таблицам.</span><span class="sxs-lookup"><span data-stu-id="e292d-p107">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="e292d-129">Прокси-объекты</span><span class="sxs-lookup"><span data-stu-id="e292d-129">Proxy objects</span></span>

<span data-ttu-id="e292d-p108">Объекты JavaScript в Excel, которые вы объявляете и используете в надстройке, представляют собой прокси-объекты. Все методы, которые вы вызываете, либо свойства, которые вы настраиваете либо загружаете, в прокси-объектах просто добавляются в очередь команд, ожидающих выполнения. Когда вы вызываете метод **sync()** в контексте запроса (например, `context.sync()`), команды, помещенные в очередь, передаются в Excel и выполняются. По существу, API JavaScript для Excel ориентирован на работу с пакетами. Вы можете поместить в очередь любое количество изменений в контексте запроса, а затем вызвать метод **sync()**, чтобы запустить пакет команд, помещенных в очередь.</span><span class="sxs-lookup"><span data-stu-id="e292d-p108">The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>

<span data-ttu-id="e292d-p109">Например, во фрагменте кода ниже показано, как объявить локальный объект JavaScript **selectedRange** для ссылки на выделенный диапазон в документе Excel, а затем задать ряд свойств для этого объекта. Объект **selectedRange** представляет собой прокси-объект, поэтому свойства, заданные в этом объекте, и методы, вызванные в этом объекте, не будут отображены в документе Excel, пока надстройка не вызовет метод **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="e292d-p109">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object. The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="sync"></a><span data-ttu-id="e292d-137">sync()</span><span class="sxs-lookup"><span data-stu-id="e292d-137">sync()</span></span>

<span data-ttu-id="e292d-p110">При вызове метода **sync()** в контексте запроса будет синхронизировано состояние прокси-объектов и объектов в документе Excel. Метод **sync()** запускает любые команды, помещенные в очередь в контексте запроса, и получает значения для любых свойств, которые следует загрузить, в прокси-объектах. Метод **sync()** выполняется асинхронно и возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается по завершении работы метода **sync()**.</span><span class="sxs-lookup"><span data-stu-id="e292d-p110">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document. The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>

<span data-ttu-id="e292d-141">В примере ниже показана пакетная функция, которая определяет локальный прокси-объект JavaScript (**selectedRange**), загружает свойство этого объекта, а затем использует шаблон JavaScript Promises для вызова метода **context.sync()** и, соответственно, синхронизации состояния прокси-объектов и объектов в документе Excel.</span><span class="sxs-lookup"><span data-stu-id="e292d-141">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>

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

<span data-ttu-id="e292d-142">В предыдущем примере при вызове метода **context.sync()** задается объект **selectedRange** и загружается его свойство **address**.</span><span class="sxs-lookup"><span data-stu-id="e292d-142">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>

<span data-ttu-id="e292d-143">Так как **sync()**  — асинхронная операция, возвращающая обещание, вам всегда следует **возвращать** обещание (в JavaScript).</span><span class="sxs-lookup"><span data-stu-id="e292d-143">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript).</span></span> <span data-ttu-id="e292d-144">Это гарантирует, что операция **sync()** будет завершена до того как продолжится выполнение скрипта.</span><span class="sxs-lookup"><span data-stu-id="e292d-144">Doing so ensures that the **sync()** operation completes before the script continues to run.</span></span> <span data-ttu-id="e292d-145">Дополнительные сведения об оптимизации производительности с помощью метода **sync()** см. в статье [Оптимизация производительности API JavaScript для Excel](/office/dev/add-ins/excel/performance).</span><span class="sxs-lookup"><span data-stu-id="e292d-145">For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](/office/dev/add-ins/excel/performance).</span></span>

### <a name="load"></a><span data-ttu-id="e292d-146">load()</span><span class="sxs-lookup"><span data-stu-id="e292d-146">load()</span></span>

<span data-ttu-id="e292d-p112">Чтобы можно было считывать свойства прокси-объекта, вам необходимо явно загрузить их и заполнить прокси-объект данными из документа Excel, а затем вызвать метод **context.sync()**. Например, вы создали прокси-объект для ссылки на выделенный диапазон, а затем вам потребовалось считать свойство **address** выделенного диапазона. Прежде чем вы сможете считать свойство **address**, вам потребуется загрузить его. Чтобы запросить загрузку свойств прокси-объекта, вызовите метод **load()** в объекте и укажите свойства, которые необходимо загрузить.</span><span class="sxs-lookup"><span data-stu-id="e292d-p112">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it. To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="e292d-p113">Если вы вызываете методы или задаете свойства только в прокси-объекте, вам не нужно вызывать объект **load()**. Метод **load()** требуется только тогда, когда вам необходимо считать свойства в прокси-объекте.</span><span class="sxs-lookup"><span data-stu-id="e292d-p113">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method. The **load()** method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="e292d-p114">Аналогично запросам для задания свойств или вызова методов в прокси-объектах, запросы на загрузку свойств в прокси-объектах добавляются в очередь команд, ожидающих выполнения, в контексте запроса, который будет запущен, когда вы в следующий раз вызовете метод **sync()**. В очередь можно поставить сколько угодно вызовов **load()** в контексте запроса.</span><span class="sxs-lookup"><span data-stu-id="e292d-p114">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>

<span data-ttu-id="e292d-154">В примере ниже загружаются только определенные свойства диапазона.</span><span class="sxs-lookup"><span data-stu-id="e292d-154">In the following example, only specific properties of the range are loaded.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);

    return context.sync()
      .then(function () {
        console.log (myRange.address);              // ok
        console.log (myRange.format.wrapText);      // ok
        console.log (myRange.format.fill.color);    // ok
        //console.log (myRange.format.font.color);  // not ok as it was not loaded
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

<span data-ttu-id="e292d-155">Так как в предыдущем примере в вызове метода **myRange.load()** не указан `format/font`, вам не удастся считать свойство `format.font.color`.</span><span class="sxs-lookup"><span data-stu-id="e292d-155">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="e292d-156">Чтобы оптимизировать производительность, при использовании метода **load()** в объекте вам следует явно указать свойства и связи, которые необходимо загрузить, как описано в статье [Оптимизация производительности API JavaScript для Excel](performance.md).</span><span class="sxs-lookup"><span data-stu-id="e292d-156">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md).</span></span> <span data-ttu-id="e292d-157">Дополнительные сведения о методе **load()** см. в статье [Дополнительные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-advanced-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="e292d-157">For more information about the **load()** method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="e292d-158">Значения null или пустые значения свойств</span><span class="sxs-lookup"><span data-stu-id="e292d-158">null or blank property values</span></span>

### <a name="null-input-in-2-d-array"></a><span data-ttu-id="e292d-159">Входное значение null в двумерном массиве</span><span class="sxs-lookup"><span data-stu-id="e292d-159">null input in 2-D Array</span></span>

<span data-ttu-id="e292d-p116">В Excel диапазон представлен двумерным массивом, в котором первое измерение — это строки, а второе — столбцы. Чтобы задать значения, формат чисел или формулу только для определенных ячеек в диапазоне, укажите значения, формат чисел или формулу для этих ячеек в двумерном массиве, а для всех остальных ячеек в этом массиве укажите значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e292d-p116">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="e292d-p117">Например, чтобы изменить формат чисел только для одной ячейки в диапазоне и сохранить существующий формат чисел для всех остальных ячеек в диапазоне, укажите новый формат чисел для ячейки, которую необходимо изменить, а для всех остальных ячеек укажите значение `null`. Во фрагменте кода ниже показано, как задать новый формат чисел для четвертой ячейки в диапазоне, при этом формат чисел для первых трех ячеек в диапазоне останется неизменным.</span><span class="sxs-lookup"><span data-stu-id="e292d-p117">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a><span data-ttu-id="e292d-164">Входное значение null для свойства</span><span class="sxs-lookup"><span data-stu-id="e292d-164">null input for a property</span></span>

<span data-ttu-id="e292d-p118">`null` не является допустимым входным значением для одного свойства. Например, указанный ниже фрагмент кода не является допустимым, так как свойство **values** диапазона не должно иметь значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e292d-p118">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>

```js
range.values = null;
```

<span data-ttu-id="e292d-167">Аналогично, указанный ниже фрагмент кода не является допустимым, так как `null` — недопустимое значение для свойства **color**.</span><span class="sxs-lookup"><span data-stu-id="e292d-167">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a><span data-ttu-id="e292d-168">Значения свойств null в ответе</span><span class="sxs-lookup"><span data-stu-id="e292d-168">null property values in the response</span></span>

<span data-ttu-id="e292d-p119">Если в указанном диапазоне имеются другие значения, свойства форматирования, например `size` и `color` будут содержать значения `null` в ответе. Например, если вы получаете диапазон и загружаете его свойство `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="e292d-p119">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

- <span data-ttu-id="e292d-171">Если у всех ячеек в диапазоне один и тот же цвет шрифта, свойство `range.format.font.color` указывает этот цвет.</span><span class="sxs-lookup"><span data-stu-id="e292d-171">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
- <span data-ttu-id="e292d-172">Если в диапазоне используется несколько цветов шрифтов, свойство `range.format.font.color` имеет значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e292d-172">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

### <a name="blank-input-for-a-property"></a><span data-ttu-id="e292d-173">Пустое входное значение для свойства</span><span class="sxs-lookup"><span data-stu-id="e292d-173">Blank input for a property</span></span>

<span data-ttu-id="e292d-p120">Когда вы указываете пустое значение для свойства (то есть две кавычки подряд без других знаков между `''`), это будет интерпретировано как инструкция по очистке или сбросу свойства. Например:</span><span class="sxs-lookup"><span data-stu-id="e292d-p120">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

- <span data-ttu-id="e292d-176">Если вы укажете пустое значение для свойства `values` диапазона, содержимое диапазона будет очищено.</span><span class="sxs-lookup"><span data-stu-id="e292d-176">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>

- <span data-ttu-id="e292d-177">Если вы укажете пустое значение для свойства `numberFormat`, формат чисел будет "сброшен" до формата `General`.</span><span class="sxs-lookup"><span data-stu-id="e292d-177">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>

- <span data-ttu-id="e292d-178">Если вы укажете пустое значение для свойств `formula` и `formulaLocale`, значения формул будут очищены.</span><span class="sxs-lookup"><span data-stu-id="e292d-178">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="e292d-179">Значения пустых свойств в ответе</span><span class="sxs-lookup"><span data-stu-id="e292d-179">Blank property values in the response</span></span>

<span data-ttu-id="e292d-p121">Для операций чтения пустое значение свойства в ответе (то есть две кавычки подряд без других знаков между `''`) указывает, что ячейка не содержит данных или значения. В первом примере ниже первая и последняя ячейки в диапазоне не содержат данных. Во втором примере две первые ячейки в диапазоне не содержат формул.</span><span class="sxs-lookup"><span data-stu-id="e292d-p121">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="e292d-183">Чтение из неограниченного диапазона и запись в него</span><span class="sxs-lookup"><span data-stu-id="e292d-183">Read or write to an unbounded range</span></span>

### <a name="read-an-unbounded-range"></a><span data-ttu-id="e292d-184">Чтение из неограниченного диапазона</span><span class="sxs-lookup"><span data-stu-id="e292d-184">Read an unbounded range</span></span>

<span data-ttu-id="e292d-p122">Адрес неограниченного диапазона представляет собой адрес диапазона, указывающий весь столбец (столбцы) либо всю строку (строки). Например:</span><span class="sxs-lookup"><span data-stu-id="e292d-p122">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>

- <span data-ttu-id="e292d-187">Адреса диапазона включают в себя весь столбец (столбцы):</span><span class="sxs-lookup"><span data-stu-id="e292d-187">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="e292d-188">Адреса диапазона включают в себя всю строку (строки):</span><span class="sxs-lookup"><span data-stu-id="e292d-188">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

<span data-ttu-id="e292d-p123">Когда API отправляет запрос на получение неограниченного диапазона (например, `getRange('C:C')`), ответ будет содержать значения `null` для свойств уровня ячейки, например свойств `values`, `text`, `numberFormat` и `formula`. Другие свойства диапазона, например `address` и `cellCount`, будут содержать допустимые значения для неограниченного диапазона.</span><span class="sxs-lookup"><span data-stu-id="e292d-p123">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="e292d-191">Запись в неограниченный диапазон</span><span class="sxs-lookup"><span data-stu-id="e292d-191">Write to an unbounded range</span></span>

<span data-ttu-id="e292d-p124">Вам не удастся задать свойства уровня ячейки, например `values`, `numberFormat` и `formula`, в неограниченном диапазоне, так как входной запрос слишком велик. Например, приведенный ниже фрагмент кода недопустим, так как он пытается указать свойство `values` для неограниченного диапазона. Если вы попытаетесь задать свойства уровня ячейки для неограниченного диапазона, API возвратит ошибку.</span><span class="sxs-lookup"><span data-stu-id="e292d-p124">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large. For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range. The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="e292d-195">Чтение из большого диапазона и запись в него</span><span class="sxs-lookup"><span data-stu-id="e292d-195">Read or write to a large range</span></span>

<span data-ttu-id="e292d-p125">Если диапазон содержит большое количество ячеек, значений, форматов чисел или формул, то, возможно, не удастся выполнить операции API над этим диапазоном. API всегда делает все возможное, чтобы выполнить запрошенную операцию над диапазоном (то есть получить или записать указанные данные), но попытка выполнить операцию чтения или записи для большого диапазона может привести к ошибке API из-за чрезмерного потребления ресурсов. Чтобы избежать таких ошибок, мы рекомендуем выполнять отдельные операции чтения или записи для небольших подмножеств большого диапазона, а не пытаться выполнить одну операцию чтения или записи для большого диапазона.</span><span class="sxs-lookup"><span data-stu-id="e292d-p125">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e292d-199">В Excel Online применяется ограничение **5 МБ** для размера полезных данных запросов и откликов.</span><span class="sxs-lookup"><span data-stu-id="e292d-199">Excel Online has a payload size limit for requests and responses of **5MB**.</span></span> <span data-ttu-id="e292d-200">При превышении этого ограничения возникает ошибка `RichAPI.Error`.</span><span class="sxs-lookup"><span data-stu-id="e292d-200">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="e292d-201">Изменение всех ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="e292d-201">Update all cells in a range</span></span>

<span data-ttu-id="e292d-202">Если необходимо одинаково изменить все ячейки в диапазоне (например, заполнить все ячейки одним и тем же значением или формулой либо задать один и тот же формат чисел), задайте для соответствующего свойства в объекте **range** (одно) необходимое значение.</span><span class="sxs-lookup"><span data-stu-id="e292d-202">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>

<span data-ttu-id="e292d-203">В примере ниже показано, как получить диапазон, содержащий 20 ячеек, а затем задать формат чисел и заполнить все ячейки в диапазоне значением **3/11/2015** (11.03.2015).</span><span class="sxs-lookup"><span data-stu-id="e292d-203">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = context.workbook.worksheets.getItem(sheetName);

    var range = worksheet.getRange(rangeAddress);
    range.numberFormat = 'm/d/yyyy';
    range.values = '3/11/2015';
    range.load('text');

    return context.sync()
      .then(function () {
        console.log(range.text);
    });
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
      console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="handle-errors"></a><span data-ttu-id="e292d-204">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="e292d-204">Handle errors</span></span>

<span data-ttu-id="e292d-205">При возникновении ошибки в интерфейсе API он возвращает объект **error**, содержащий код и сообщение.</span><span class="sxs-lookup"><span data-stu-id="e292d-205">When an API error occurs, the API returns an **error** object that contains a code and a message.</span></span> <span data-ttu-id="e292d-206">Подробные сведения об обработке ошибок, включая список ошибок API, см. в статье [Обработка ошибок](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="e292d-206">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e292d-207">См. также</span><span class="sxs-lookup"><span data-stu-id="e292d-207">See also</span></span>

- [<span data-ttu-id="e292d-208">Начало работы с надстройками Excel</span><span class="sxs-lookup"><span data-stu-id="e292d-208">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
- [<span data-ttu-id="e292d-209">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="e292d-209">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
- [<span data-ttu-id="e292d-210">Дополнительные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e292d-210">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="e292d-211">Оптимизация производительности API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e292d-211">Excel JavaScript API performance optimization</span></span>](/office/dev/add-ins/excel/performance)
- [<span data-ttu-id="e292d-212">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e292d-212">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

---
title: Основные понятия API JavaScript для Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 37d652d2ad2f323d0f94583e530e91e775e06ddf
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925411"
---
# <a name="excel-javascript-api-core-concepts"></a><span data-ttu-id="7d4ff-102">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="7d4ff-102">Excel JavaScript API core concepts</span></span>
 
<span data-ttu-id="7d4ff-103">В этой статье рассказывается, как создавать надстройки для Excel 2016 с помощью [API JavaScript для Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="7d4ff-103">This article describes how to use the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) to build add-ins for Excel 2016.</span></span> <span data-ttu-id="7d4ff-104">В статье изложены основные принципы, которые являются фундаментальными при использовании этого API, а также имеются рекомендации по выполнению определенных задач, например чтению данных из большого диапазона или записи данных в него, изменения всех ячеек в диапазоне и много другого.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-104">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="7d4ff-105">Асинхронный характер API Excel</span><span class="sxs-lookup"><span data-stu-id="7d4ff-105">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="7d4ff-106">Веб-надстройки Excel работают в контейнере браузера, внедренном в приложение Office на платформах для настольных ПК, например Office для Windows, и работающем в iFrame HTML в Office Online.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-106">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online.</span></span> <span data-ttu-id="7d4ff-107">Вам не удастся заставить API Office.js синхронно взаимодействовать с ведущим приложением Excel на всех поддерживаемых платформах из-за соображений производительности.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-107">Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations.</span></span> <span data-ttu-id="7d4ff-108">Таким образом, вызов API **sync()** в Office.js возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается, когда приложение Excel выполняет запрошенные действия чтения или записи.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-108">Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions.</span></span> <span data-ttu-id="7d4ff-109">Кроме того, вы можете поместить в очередь несколько действий, например действия настройки свойств или вызова методов, а затем запустить их в виде пакета команд в одном вызове метода **sync()**, а не отправлять отдельные запросы для каждого действия.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-109">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action.</span></span> <span data-ttu-id="7d4ff-110">В разделах ниже описано, как сделать это, используя API **Excel.run()** и **sync()**.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-110">The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="7d4ff-111">Excel.run</span><span class="sxs-lookup"><span data-stu-id="7d4ff-111">Excel.run</span></span>
 
<span data-ttu-id="7d4ff-112">**Excel.run** выполняет функцию, в которой вы указываете действия, которые необходимо совершить над объектной моделью Excel.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-112">**Excel.run** executes a function where you specify the actions to perform against the Excel object model.</span></span> <span data-ttu-id="7d4ff-113">**Excel.run** автоматически создает контекст запроса, который вы можете использовать для взаимодействия с объектами Excel.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-113">**Excel.run** automatically creates a request context that you can use to interact with Excel objects.</span></span> <span data-ttu-id="7d4ff-114">Когда **Excel.run** завершает работу, обещание разрешается, и все объекты, которые были выделены в среде выполнения, будут автоматически разблокированы.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-114">When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="7d4ff-115">В примере ниже показано, как использовать **Excel.run**.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-115">The following example shows how to use **Excel.run**.</span></span> <span data-ttu-id="7d4ff-116">Оператор catch перехватывает и записывает ошибки, возникающие в **Excel.run**, в журнал.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-116">The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
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

## <a name="request-context"></a><span data-ttu-id="7d4ff-117">Контекст запроса</span><span class="sxs-lookup"><span data-stu-id="7d4ff-117">Request context</span></span>
 
<span data-ttu-id="7d4ff-p105">Excel и ваша надстройка запускаются как два отдельных процесса. Так как они используют разные среды выполнения, надстройкам Excel требуется объект **RequestContext**, чтобы можно было подключать надстройку к объектам в Excel, например к листам, диапазонам, диаграммам и таблицам.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="7d4ff-120">Прокси-объекты</span><span class="sxs-lookup"><span data-stu-id="7d4ff-120">Proxy objects</span></span>
 
<span data-ttu-id="7d4ff-121">Объекты JavaScript в Excel, которые вы объявляете и используете в надстройке, представляют собой прокси-объекты.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-121">The Excel JavaScript objects that you declare and use in an add-in are proxy objects.</span></span> <span data-ttu-id="7d4ff-122">Все методы, которые вы вызываете, либо свойства, которые вы настраиваете либо загружаете, в прокси-объектах просто добавляются в очередь команд, ожидающих выполнения.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-122">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="7d4ff-123">Когда вы вызываете метод **sync()** в контексте запроса (например, `context.sync()`), команды, помещенные в очередь, передаются в Excel и выполняются.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-123">When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run.</span></span> <span data-ttu-id="7d4ff-124">По существу, API JavaScript для Excel ориентирован на работу с пакетами.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-124">The Excel JavaScript API is fundamentally batch-centric.</span></span> <span data-ttu-id="7d4ff-125">Вы можете поместить в очередь любое количество изменений в контексте запроса, а затем вызвать метод **sync()**, чтобы запустить пакет команд, помещенных в очередь.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-125">You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="7d4ff-126">Например, во фрагменте кода ниже показано, как объявить локальный объект JavaScript **selectedRange** для ссылки на выделенный диапазон в документе Excel, а затем задать ряд свойств для этого объекта.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-126">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object.</span></span> <span data-ttu-id="7d4ff-127">Объект **selectedRange** представляет собой прокси-объект, поэтому свойства, заданные в этом объекте, и методы, вызванные в этом объекте, не будут отображены в документе Excel, пока надстройка не вызовет метод **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-127">The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="7d4ff-128">sync()</span><span class="sxs-lookup"><span data-stu-id="7d4ff-128">sync()</span></span>
 
<span data-ttu-id="7d4ff-129">При вызове метода **sync()** в контексте запроса будет синхронизировано состояние прокси-объектов и объектов в документе Excel.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-129">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document.</span></span> <span data-ttu-id="7d4ff-130">Метод **sync()** запускает любые команды, помещенные в очередь в контексте запроса, и получает значения для любых свойств, которые следует загрузить, в прокси-объектах.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-130">The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="7d4ff-131">Метод **sync()** выполняется асинхронно и возвращает [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается по завершении работы метода **sync()**.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-131">The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="7d4ff-132">В примере ниже показана пакетная функция, которая определяет локальный прокси-объект JavaScript (**selectedRange**), загружает свойство этого объекта, а затем использует шаблон JavaScript Promises для вызова метода **context.sync()** и, соответственно, синхронизации состояния прокси-объектов и объектов в документе Excel.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-132">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
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
 
<span data-ttu-id="7d4ff-133">В предыдущем примере при вызове метода **context.sync()** задается объект **selectedRange** и загружается его свойство **address**.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-133">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="7d4ff-134">Так как **sync()** — асинхронная операция, возвращающая обещание, вам всегда следует **возвращать** обещание (в JavaScript).</span><span class="sxs-lookup"><span data-stu-id="7d4ff-134">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript).</span></span> <span data-ttu-id="7d4ff-135">Это гарантирует, что операция **sync()** будет завершена до того как продолжится выполнение сценария.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-135">Doing so ensures that the **sync()** operation completes before the script continues to run.</span></span> <span data-ttu-id="7d4ff-136">Дополнительную информацию об оптимизации производительности с помощью метода **sync()** см. в статье [Оптимизация производительности API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span><span class="sxs-lookup"><span data-stu-id="7d4ff-136">For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span></span>
 
### <a name="load"></a><span data-ttu-id="7d4ff-137">load()</span><span class="sxs-lookup"><span data-stu-id="7d4ff-137">load()</span></span>
 
<span data-ttu-id="7d4ff-138">Чтобы можно было считывать свойства прокси-объекта, вам необходимо явно загрузить их и заполнить прокси-объект данными из документа Excel, а затем вызвать метод **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-138">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**.</span></span> <span data-ttu-id="7d4ff-139">Например, вы создали прокси-объект для ссылки на выделенный диапазон, а затем вам потребовалось считать свойство **address** выделенного диапазона. Прежде чем вы сможете считать свойство **address**, вам потребуется загрузить его.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-139">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it.</span></span> <span data-ttu-id="7d4ff-140">Чтобы запросить загрузку свойств прокси-объекта, вызовите метод **load()** в объекте и укажите свойства, которые необходимо загрузить.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-140">To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="7d4ff-141">Если вы вызываете методы или задаете свойства только в прокси-объекте, вам не нужно вызывать объект **load()**.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-141">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method.</span></span> <span data-ttu-id="7d4ff-142">Метод **load()** требуется только тогда, когда вам необходимо считать свойства в прокси-объекте.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-142">The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="7d4ff-p112">Аналогично запросам для задания свойств или вызова методов в прокси-объектах, запросы на загрузку свойств в прокси-объектах добавляются в очередь команд, ожидающих выполнения, в контексте запроса, который будет запущен, когда вы в следующий раз вызовете метод **sync()**. В очередь можно поставить сколько угодно вызовов **load()** в контексте запроса.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="7d4ff-145">В примере ниже загружаются только определенные свойства диапазона.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-145">In the following example, only specific properties of the range are loaded.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
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
 
<span data-ttu-id="7d4ff-146">Так как в предыдущем примере в вызове метода **myRange.load()** не указан `format/font`, вам не удастся считать свойство `format.font.color`.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-146">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="7d4ff-147">Чтобы оптимизировать производительность, явно укажите свойства и отношения, загружаемые при использовании метода **load()** для объекта, как описано в статье [Оптимизация производительности API Excel JavaScript](performance.md).</span><span class="sxs-lookup"><span data-stu-id="7d4ff-147">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object.</span></span> <span data-ttu-id="7d4ff-148">Дополнительные сведения о методе **load()** см. в статье [Расширенные концепции API JavaScript для Excel](excel-add-ins-advanced-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="7d4ff-148">For more information about the **load()** method, see [Excel JavaScript API advanced concepts](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="7d4ff-149">Значения null или пустые значения свойств</span><span class="sxs-lookup"><span data-stu-id="7d4ff-149">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="7d4ff-150">Входное значение null в двумерном массиве</span><span class="sxs-lookup"><span data-stu-id="7d4ff-150">null input in 2-D Array</span></span>
 
<span data-ttu-id="7d4ff-151">В Excel диапазон представлен двумерным массивом, в котором первое измерение — это строки, а второе — столбцы.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-151">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns.</span></span> <span data-ttu-id="7d4ff-152">Чтобы задать значения, формат чисел или формулу только для определенных ячеек в диапазоне, укажите значения, формат чисел или формулу для этих ячеек в двумерном массиве, а для всех остальных ячеек в этом массиве укажите значение `null`.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-152">To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="7d4ff-153">Например, чтобы изменить формат чисел только для одной ячейки в диапазоне и сохранить существующий формат чисел для всех остальных ячеек в диапазоне, укажите новый формат чисел для ячейки, которую необходимо изменить, а для всех остальных ячеек укажите значение `null`.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-153">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells.</span></span> <span data-ttu-id="7d4ff-154">Во фрагменте кода ниже показано, как задать новый формат чисел для четвертой ячейки в диапазоне, при этом формат чисел для первых трех ячеек в диапазоне останется неизменным.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-154">The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="7d4ff-155">Входное значение null для свойства</span><span class="sxs-lookup"><span data-stu-id="7d4ff-155">null input for a property</span></span>
 
<span data-ttu-id="7d4ff-p116">`null` не является допустимым входным значением для одного свойства. Например, указанный ниже фрагмент кода не является допустимым, так как свойство **values** диапазона не должно иметь значение `null`.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="7d4ff-158">Аналогично, указанный ниже фрагмент кода не является допустимым, так как `null` — недопустимое значение для свойства **color**.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-158">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="7d4ff-159">Значения свойств null в ответе</span><span class="sxs-lookup"><span data-stu-id="7d4ff-159">null property values in the response</span></span>
 
<span data-ttu-id="7d4ff-160">Если в указанном диапазоне имеются другие значения, свойства форматирования, например `size` и `color` будут содержать значения `null` в ответе.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-160">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range.</span></span> <span data-ttu-id="7d4ff-161">Например, если вы получаете диапазон и загружаете его свойство `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="7d4ff-161">For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="7d4ff-162">Если у всех ячеек в диапазоне один и тот же цвет шрифта, свойство `range.format.font.color` указывает этот цвет.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-162">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="7d4ff-163">Если в диапазоне используется несколько цветов шрифтов, свойство `range.format.font.color` имеет значение `null`.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-163">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="7d4ff-164">Пустое входное значение для свойства</span><span class="sxs-lookup"><span data-stu-id="7d4ff-164">Blank input for a property</span></span>
 
<span data-ttu-id="7d4ff-p118">Когда вы указываете пустое значение для свойства (то есть две кавычки подряд без других знаков между `''`), это будет интерпретировано как инструкция по очистке или сбросу свойства. Например:</span><span class="sxs-lookup"><span data-stu-id="7d4ff-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="7d4ff-167">Если вы укажете пустое значение для свойства `values` диапазона, содержимое диапазона будет очищено.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-167">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="7d4ff-168">Если вы укажете пустое значение для свойства `numberFormat`, формат чисел будет "сброшен" до формата `General`.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-168">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="7d4ff-169">Если вы укажете пустое значение для свойств `formula` и `formulaLocale`, значения формул будут очищены.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-169">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="7d4ff-170">Значения пустых свойств в ответе</span><span class="sxs-lookup"><span data-stu-id="7d4ff-170">Blank property values in the response</span></span>
 
<span data-ttu-id="7d4ff-171">Для операций чтения пустое значение свойства в ответе (то есть две кавычки подряд без других знаков между `''`) указывает, что ячейка не содержит данных или значения.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-171">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value.</span></span> <span data-ttu-id="7d4ff-172">В первом примере ниже первая и последняя ячейки в диапазоне не содержат данных.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-172">In the first example below, the first and last cell in the range contain no data.</span></span> <span data-ttu-id="7d4ff-173">Во втором примере две первые ячейки в диапазоне не содержат формул.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-173">In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="7d4ff-174">Чтение из неограниченного диапазона и запись в него</span><span class="sxs-lookup"><span data-stu-id="7d4ff-174">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="7d4ff-175">Чтение из неограниченного диапазона</span><span class="sxs-lookup"><span data-stu-id="7d4ff-175">Read an unbounded range</span></span>
 
<span data-ttu-id="7d4ff-p120">Адрес неограниченного диапазона представляет собой адрес диапазона, указывающий весь столбец (столбцы) либо всю строку (строки). Например:</span><span class="sxs-lookup"><span data-stu-id="7d4ff-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="7d4ff-178">Адреса диапазона включают в себя весь столбец (столбцы):</span><span class="sxs-lookup"><span data-stu-id="7d4ff-178">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="7d4ff-179">Адреса диапазона включают в себя всю строку (строки):</span><span class="sxs-lookup"><span data-stu-id="7d4ff-179">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="7d4ff-180">Когда API отправляет запрос на получение неограниченного диапазона (например, `getRange('C:C')`), ответ будет содержать значения `null` для свойств уровня ячейки, например свойств `values`, `text`, `numberFormat` и `formula`.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-180">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`.</span></span> <span data-ttu-id="7d4ff-181">Другие свойства диапазона, например `address` и `cellCount`, будут содержать допустимые значения для неограниченного диапазона.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-181">Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="7d4ff-182">Запись в неограниченный диапазон</span><span class="sxs-lookup"><span data-stu-id="7d4ff-182">Write to an unbounded range</span></span>
 
<span data-ttu-id="7d4ff-183">Вам не удастся задать свойства уровня ячейки, например `values`, `numberFormat` и `formula`, в неограниченном диапазоне, так как входной запрос слишком велик.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-183">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large.</span></span> <span data-ttu-id="7d4ff-184">Например, приведенный ниже фрагмент кода недопустим, так как он пытается указать свойство `values` для неограниченного диапазона.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-184">For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="7d4ff-185">Если вы попытаетесь задать свойства уровня ячейки для неограниченного диапазона, API возвратит ошибку.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-185">The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="7d4ff-186">Чтение из большого диапазона и запись в него</span><span class="sxs-lookup"><span data-stu-id="7d4ff-186">Read or write to a large range</span></span>
 
<span data-ttu-id="7d4ff-187">Если диапазон содержит большое количество ячеек, значений, форматов чисел или формул, то, возможно, не удастся выполнить операции API над этим диапазоном.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-187">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="7d4ff-188">API всегда делает все возможное, чтобы выполнить запрошенную операцию над диапазоном (то есть получить или записать указанные данные), но попытка выполнить операцию чтения или записи для большого диапазона может привести к ошибке API из-за чрезмерного потребления ресурсов.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-188">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="7d4ff-189">Чтобы избежать таких ошибок, мы рекомендуем выполнять отдельные операции чтения или записи для небольших подмножеств большого диапазона, а не пытаться выполнить одну операцию чтения или записи для большого диапазона.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-189">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="7d4ff-190">Изменение всех ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="7d4ff-190">Update all cells in a range</span></span>
 
<span data-ttu-id="7d4ff-191">Если необходимо одинаково изменить все ячейки в диапазоне (например, заполнить все ячейки одним и тем же значением или формулой либо задать один и тот же формат чисел), задайте для соответствующего свойства в объекте **range** (одно) необходимое значение.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-191">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="7d4ff-192">В примере ниже показано, как получить диапазон, содержащий 20 ячеек, а затем задать формат чисел и заполнить все ячейки в диапазоне значением **3/11/2015** (11.03.2015).</span><span class="sxs-lookup"><span data-stu-id="7d4ff-192">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
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
 
## <a name="error-messages"></a><span data-ttu-id="7d4ff-193">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="7d4ff-193">Error messages</span></span>
 
<span data-ttu-id="7d4ff-194">При возникновении ошибки в API он возвратит объект **error**, содержащий код и сообщение.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-194">When an API error occurs, the API will return an **error** object that contains a code and a message.</span></span> <span data-ttu-id="7d4ff-195">В таблице ниже перечислены ошибки, которые может возвращать API.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-195">The following table defines a list of errors that the API may return.</span></span>
 
|<span data-ttu-id="7d4ff-196">error.code</span><span class="sxs-lookup"><span data-stu-id="7d4ff-196">error.code</span></span> | <span data-ttu-id="7d4ff-197">error.message</span><span class="sxs-lookup"><span data-stu-id="7d4ff-197">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="7d4ff-198">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="7d4ff-198">InvalidArgument</span></span> |<span data-ttu-id="7d4ff-199">Аргумент недопустим, отсутствует или имеет неправильный формат.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-199">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="7d4ff-200">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="7d4ff-200">InvalidRequest</span></span>  |<span data-ttu-id="7d4ff-201">Не удается обработать запрос.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-201">Cannot process the request.</span></span>|
|<span data-ttu-id="7d4ff-202">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="7d4ff-202">InvalidReference</span></span>|<span data-ttu-id="7d4ff-203">Эта ссылка недопустима для текущей операции.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-203">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="7d4ff-204">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="7d4ff-204">InvalidBinding</span></span>  |<span data-ttu-id="7d4ff-205">Эта привязка объектов недопустима из-за предыдущих обновлений.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-205">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="7d4ff-206">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="7d4ff-206">InvalidSelection</span></span>|<span data-ttu-id="7d4ff-207">Выбранный фрагмент недопустим для этой операции.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-207">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="7d4ff-208">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="7d4ff-208">Unauthenticated</span></span> |<span data-ttu-id="7d4ff-209">Требуемые сведения о проверке подлинности отсутствуют или недопустимы.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-209">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="7d4ff-210">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="7d4ff-210">AccessDenied</span></span> |<span data-ttu-id="7d4ff-211">Вы не можете выполнить запрашиваемую операцию.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-211">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="7d4ff-212">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="7d4ff-212">ItemNotFound</span></span> |<span data-ttu-id="7d4ff-213">Запрашиваемый ресурс не существует.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-213">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="7d4ff-214">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="7d4ff-214">ActivityLimitReached</span></span>|<span data-ttu-id="7d4ff-215">Достигнут предел действий.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-215">Activity limit has been reached.</span></span>|
|<span data-ttu-id="7d4ff-216">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7d4ff-216">GeneralException</span></span>|<span data-ttu-id="7d4ff-217">При обработке запроса возникла внутренняя ошибка.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-217">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="7d4ff-218">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="7d4ff-218">NotImplemented</span></span>  |<span data-ttu-id="7d4ff-219">Запрашиваемая функция не реализована.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-219">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="7d4ff-220">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="7d4ff-220">ServiceNotAvailable</span></span>|<span data-ttu-id="7d4ff-221">Служба недоступна.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-221">The service is unavailable.</span></span>|
|<span data-ttu-id="7d4ff-222">Conflict</span><span class="sxs-lookup"><span data-stu-id="7d4ff-222">Conflict</span></span>              |<span data-ttu-id="7d4ff-223">Запрос не удалось обработать из-за конфликта.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-223">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="7d4ff-224">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="7d4ff-224">ItemAlreadyExists</span></span>|<span data-ttu-id="7d4ff-225">Создаваемый ресурс уже существует.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-225">The resource being created already exists.</span></span>|
|<span data-ttu-id="7d4ff-226">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="7d4ff-226">UnsupportedOperation</span></span>|<span data-ttu-id="7d4ff-227">Выполняемая операция не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-227">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="7d4ff-228">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="7d4ff-228">RequestAborted</span></span>|<span data-ttu-id="7d4ff-229">Запрос прерван во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-229">The request was aborted during run time.</span></span>|
|<span data-ttu-id="7d4ff-230">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="7d4ff-230">ApiNotAvailable</span></span>|<span data-ttu-id="7d4ff-231">Запрашиваемый интерфейс API недоступен.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-231">The requested API is not available.</span></span>|
|<span data-ttu-id="7d4ff-232">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="7d4ff-232">InsertDeleteConflict</span></span>|<span data-ttu-id="7d4ff-233">Операция вставки или удаления привела к конфликту.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-233">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="7d4ff-234">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="7d4ff-234">InvalidOperation</span></span>|<span data-ttu-id="7d4ff-235">Выполняемая операция недопустима для этого объекта.</span><span class="sxs-lookup"><span data-stu-id="7d4ff-235">The operation attempted is invalid on the object.</span></span>|
 
## <a name="see-also"></a><span data-ttu-id="7d4ff-236">См. также</span><span class="sxs-lookup"><span data-stu-id="7d4ff-236">See also</span></span>
 
* [<span data-ttu-id="7d4ff-237">Начало работы с надстройками Excel</span><span class="sxs-lookup"><span data-stu-id="7d4ff-237">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="7d4ff-238">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="7d4ff-238">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [<span data-ttu-id="7d4ff-239">Оптимизация производительности JavaScript API для Excel</span><span class="sxs-lookup"><span data-stu-id="7d4ff-239">Excel JavaScript API performance optimization</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [<span data-ttu-id="7d4ff-240">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="7d4ff-240">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)

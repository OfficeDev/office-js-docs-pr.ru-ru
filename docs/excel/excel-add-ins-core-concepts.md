---
title: Основные принципы программирования с помощью API JavaScript для Excel
description: Вы можете использовать API JavaScript для Excel, чтобы создавать надстройки для Excel.
ms.date: 10/16/2018
ms.openlocfilehash: 21fcbc32dab2057f6d1f5e97a62a6b00bf1f7f03
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579893"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="df7d0-103">Основные принципы программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="df7d0-103">Fundamental programming concepts with the Excel JavaScript API</span></span>
 
<span data-ttu-id="df7d0-p101">В этой статье описывается, как использовать [API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) для построения надстройки для Excel 2016 или более поздней версии. Также приводятся общие сведения об основных концепциях, необходимых для использования API, а также приводятся рекомендации для выполнения определенных задач, таких как чтение или запись большого диапазона, обновление всех ячеек в диапазоне и многое другое.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p101">This article describes how to use the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) to build add-ins for Excel 2016 or later. It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="df7d0-106">Асинхронный характер API Excel</span><span class="sxs-lookup"><span data-stu-id="df7d0-106">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="df7d0-p102">Веб-надстройки Excel работают в контейнере браузера, внедренном в приложение Office на платформах для настольных ПК, например Office для Windows, и работающем в iFrame HTML в Office Online. Вам не удастся заставить API Office.js синхронно взаимодействовать с основным приложением Excel на всех поддерживаемых платформах из-за соображений производительности. Таким образом, вызов API **sync()** в Office.js возвращает  [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается, когда приложение Excel выполняет запрошенные действия чтения или записи. Кроме того, вы можете поместить в очередь несколько действий, например действия настройки свойств или вызова методов, а затем запустить их в виде пакета команд в одном вызове метода  **sync()**, а не отправлять отдельные запросы для каждого действия. В разделах ниже описано, как сделать это, используя API **Excel.run()**  и **sync()**.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p102">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action. The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="df7d0-112">Excel.run</span><span class="sxs-lookup"><span data-stu-id="df7d0-112">Excel.run</span></span>
 
<span data-ttu-id="df7d0-p103">**Excel.run**  выполняет функцию, в которой вы указываете действия, которые необходимо совершить над объектной моделью Excel. **Excel.run**  автоматически создает контекст запроса, который вы можете использовать для взаимодействия с объектами Excel. Когда **Excel.run**  завершает работу, обещание разрешается, и все объекты, которые были выделены в среде выполнения, будут автоматически разблокированы.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p103">**Excel.run** executes a function where you specify the actions to perform against the Excel object model. **Excel.run** automatically creates a request context that you can use to interact with Excel objects. When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="df7d0-p104">В следующем примере показано, как использовать **Excel.run**.  Оператор catch перехватывает и записывает ошибки, возникающие в **Excel.run**,  в журнал.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p104">The following example shows how to use **Excel.run**. The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
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

## <a name="request-context"></a><span data-ttu-id="df7d0-118">Контекст запроса</span><span class="sxs-lookup"><span data-stu-id="df7d0-118">Request context</span></span>
 
<span data-ttu-id="df7d0-p105">Excel и ваша надстройка запускаются как два отдельных процесса. Так как они используют разные среды выполнения, надстройкам Excel требуется объект **RequestContext**, чтобы можно было подключать надстройку к объектам в Excel, например к листам, диапазонам, диаграммам и таблицам.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="df7d0-121">Прокси-объекты</span><span class="sxs-lookup"><span data-stu-id="df7d0-121">Proxy objects</span></span>
 
<span data-ttu-id="df7d0-p106">Объекты JavaScript в Excel, которые вы объявляете и используете в надстройке, представляют собой прокси-объекты. Все методы, которые вы вызываете, либо свойства, которые вы настраиваете либо загружаете, в прокси-объектах просто добавляются в очередь команд, ожидающих выполнения. Когда вы вызываете метод **sync()** в контексте запроса (например,  `context.sync()`), команды, помещенные в очередь, передаются в Excel и выполняются. По существу, API JavaScript для Excel ориентирован на работу с пакетами. Вы можете поместить в очередь любое количество изменений в контексте запроса, а затем вызвать метод **sync()**, чтобы запустить пакет команд, помещенных в очередь.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p106">The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="df7d0-p107">Например, в следующем фрагменте кода ниже показано, как объявить локальный объект JavaScript **selectedRange** для ссылки на выделенный диапазон в документе Excel, а затем задать ряд свойств для этого объекта. Объект **selectedRange**  представляет собой прокси-объект, поэтому заданные в этом объекте свойства и вызванные в этом объекте методы не будут отображены в документе Excel, пока надстройка не вызовет метод **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p107">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object. The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="df7d0-129">sync()</span><span class="sxs-lookup"><span data-stu-id="df7d0-129">sync()</span></span>
 
<span data-ttu-id="df7d0-p108">При вызове метода **sync()**  в контексте запроса будет синхронизировано состояние прокси-объектов и объектов в документе Excel. Метод **sync()**  запускает любые команды, помещенные в очередь в контексте запроса, и получает значения для любых свойств, которые следует загрузить, в прокси-объектах. Метод **sync()** выполняется асинхронно и возвращает  [обещание](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), которое разрешается по завершении работы метода **sync()**.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p108">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document. The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="df7d0-133">В следующем примере показана пакетная функция, которая определяет локальный прокси-объект JavaScript (**selectedRange**), загружает свойство этого объекта, а затем использует шаблон JavaScript Promises для вызова метода **context.sync()** и, соответственно, синхронизации состояния прокси-объектов и объектов в документе Excel.</span><span class="sxs-lookup"><span data-stu-id="df7d0-133">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
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
 
<span data-ttu-id="df7d0-134">В предыдущем примере задается объект **selectedRange** и загружается его свойство **address** при вызове метода **context.sync()** .</span><span class="sxs-lookup"><span data-stu-id="df7d0-134">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="df7d0-p109">Поскольку **sync()** является асинхронной операцией, возвращающей обещание, вам всегда следует **возвращать** обещание (в JavaScript). Это гарантирует, что операция **sync()**  будет завершена до того как продолжится выполнение сценария. Дополнительную информацию об оптимизации производительности с помощью метода **sync()** см. в статье [Оптимизация производительности API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span><span class="sxs-lookup"><span data-stu-id="df7d0-p109">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript). Doing so ensures that the **sync()** operation completes before the script continues to run. For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span></span>
 
### <a name="load"></a><span data-ttu-id="df7d0-138">load()</span><span class="sxs-lookup"><span data-stu-id="df7d0-138">load()</span></span>
 
<span data-ttu-id="df7d0-p110">Перед прочтением свойств прокси-объекта, вам необходимо явно загрузить их и заполнить прокси-объект данными из документа Excel, а затем вызвать метод **context.sync()**. Например, вы создали прокси-объект для ссылки на выделенный диапазон, а затем вам потребовалось считать свойство **address** выделенного диапазона. Прежде чем вы сможете считать свойство **address**, вам потребуется загрузить его. Чтобы запросить загрузку свойств прокси-объекта, вызовите метод **load()** в объекте и укажите свойства, которые необходимо загрузить.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p110">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it. To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="df7d0-p111">Если вы вызываете методы или задаете свойства только в прокси-объекте, вам не нужно вызывать объект **load()**. Метод **load()**  требуется только тогда, когда вам необходимо считать свойства в прокси-объекте.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p111">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method. The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="df7d0-p112">Аналогично запросам для задания свойств или вызова методов в прокси-объектах, запросы на загрузку свойств в прокси-объектах добавляются в очередь команд, ожидающих выполнения, в контексте запроса, который будет запущен, когда вы в следующий раз вызовете метод **sync()**. В очередь можно поставить сколько угодно вызовов **load()** в контексте запроса.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="df7d0-146">В примере ниже загружаются только определенные свойства диапазона.</span><span class="sxs-lookup"><span data-stu-id="df7d0-146">In the following example, only specific properties of the range are loaded.</span></span>
 
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
 
<span data-ttu-id="df7d0-147">Так как в предыдущем примере не указан `format/font` в вызове метода **myRange.load()**, вам не удастся считать свойство `format.font.color`.</span><span class="sxs-lookup"><span data-stu-id="df7d0-147">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="df7d0-p113">Чтобы оптимизировать производительность, явно укажите свойства и отношения, загружаемые при использовании метода **load()** для объекта, как описано в статье [Оптимизация производительности API JavaScript для Excel](performance.md). Дополнительные сведения о методе **load()**  см. в статье [Углубленные принципы программирования с использованием интерфейса API JavaScript для Excel](excel-add-ins-advanced-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="df7d0-p113">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md). For more information about the **load()** method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="df7d0-150">Значения null или пустые значения свойств</span><span class="sxs-lookup"><span data-stu-id="df7d0-150">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="df7d0-151">Входное значение null в двумерном массиве</span><span class="sxs-lookup"><span data-stu-id="df7d0-151">null input in 2-D Array</span></span>
 
<span data-ttu-id="df7d0-p114">В Excel диапазон представлен двумерным массивом, в котором первое измерение — это строки, а второе — столбцы. Чтобы задать значения, формат чисел или формулу только для определенных ячеек в диапазоне, укажите значения, формат чисел или формулу для этих ячеек в двумерном массиве, а для всех остальных ячеек в этом массиве укажите значение `null`.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p114">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="df7d0-p115">Например, чтобы изменить формат чисел только для одной ячейки в диапазоне и сохранить существующий формат чисел для всех остальных ячеек в диапазоне, укажите новый формат чисел для ячейки, которую необходимо изменить, а для всех остальных ячеек укажите значение `null`. В следующем фрагменте кода показано, как задать новый формат чисел для четвертой ячейки в диапазоне, при этом формат чисел для первых трех ячеек в диапазоне останется неизменным.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p115">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="df7d0-156">Входное значение null для свойства</span><span class="sxs-lookup"><span data-stu-id="df7d0-156">null input for a property</span></span>
 
<span data-ttu-id="df7d0-p116">`null` не является допустимым входным значением для одного свойства. Например, указанный ниже фрагмент кода не является допустимым, так как свойство **values** диапазона не должно иметь значение `null`.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="df7d0-159">Аналогично указанный ниже фрагмент кода не является допустимым, так как `null` — недопустимое значение для свойства **color**.</span><span class="sxs-lookup"><span data-stu-id="df7d0-159">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="df7d0-160">Значения свойств null в ответе</span><span class="sxs-lookup"><span data-stu-id="df7d0-160">null property values in the response</span></span>
 
<span data-ttu-id="df7d0-p117">Если в указанном диапазоне имеются другие значения, свойства форматирования, такие как `size`  и `color`, будут содержать значения `null` в ответе. Например, если вы получаете диапазон и загружаете его свойство `format.font.color` :</span><span class="sxs-lookup"><span data-stu-id="df7d0-p117">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="df7d0-163">Если у всех ячеек в диапазоне один и тот же цвет шрифта, свойство `range.format.font.color` указывает этот цвет.</span><span class="sxs-lookup"><span data-stu-id="df7d0-163">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="df7d0-164">Если в диапазоне используется несколько цветов шрифтов, свойство `range.format.font.color` имеет значение `null`.</span><span class="sxs-lookup"><span data-stu-id="df7d0-164">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="df7d0-165">Пустое входное значение для свойства</span><span class="sxs-lookup"><span data-stu-id="df7d0-165">Blank input for a property</span></span>
 
<span data-ttu-id="df7d0-p118">Когда вы указываете пустое значение для свойства (то есть две кавычки подряд без других знаков между `''`), это будет интерпретировано как инструкция по очистке или сбросу свойства. Например:</span><span class="sxs-lookup"><span data-stu-id="df7d0-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="df7d0-168">Если вы укажете пустое значение для свойства `values` диапазона, содержимое диапазона будет очищено.</span><span class="sxs-lookup"><span data-stu-id="df7d0-168">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="df7d0-169">Если вы укажете пустое значение для свойства `numberFormat`, формат чисел будет "сброшен" до формата `General`.</span><span class="sxs-lookup"><span data-stu-id="df7d0-169">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="df7d0-170">Если вы укажете пустое значение для свойств `formula` и `formulaLocale`, значения формул будут очищены.</span><span class="sxs-lookup"><span data-stu-id="df7d0-170">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="df7d0-171">Значения пустых свойств в ответе</span><span class="sxs-lookup"><span data-stu-id="df7d0-171">Blank property values in the response</span></span>
 
<span data-ttu-id="df7d0-p119">Для операций чтения пустое значение свойства в ответе (то есть две кавычки подряд без других знаков между `''`) указывает, что ячейка не содержит данных или значения. В первом примере ниже первая и последняя ячейки в диапазоне не содержат данных. Во втором примере две первые ячейки в диапазоне не содержат формул.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p119">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="df7d0-175">Чтение из неограниченного диапазона и запись в него</span><span class="sxs-lookup"><span data-stu-id="df7d0-175">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="df7d0-176">Чтение из неограниченного диапазона</span><span class="sxs-lookup"><span data-stu-id="df7d0-176">Read an unbounded range</span></span>
 
<span data-ttu-id="df7d0-p120">Адрес неограниченного диапазона представляет собой адрес диапазона, указывающий весь столбец (столбцы) либо всю строку (строки). Например:</span><span class="sxs-lookup"><span data-stu-id="df7d0-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="df7d0-179">Адреса диапазона включают в себя весь столбец (столбцы):</span><span class="sxs-lookup"><span data-stu-id="df7d0-179">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="df7d0-180">Адреса диапазона включают в себя всю строку (строки):</span><span class="sxs-lookup"><span data-stu-id="df7d0-180">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="df7d0-p121">Когда API отправляет запрос на получение неограниченного диапазона (например, `getRange('C:C')`), ответ будет содержать значения `null` для свойств уровня ячейки, такие как `values`, `text`, `numberFormat` и `formula`. Другие свойства диапазона, такие как `address` и `cellCount`, будут содержать допустимые значения для неограниченного диапазона.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p121">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="df7d0-183">Запись в неограниченный диапазон</span><span class="sxs-lookup"><span data-stu-id="df7d0-183">Write to an unbounded range</span></span>
 
<span data-ttu-id="df7d0-p122">Вам не удастся задать свойства уровня ячейки, такие как `values`, `numberFormat` и `formula` , в неограниченном диапазоне, так как входной запрос слишком велик. Например, следующий фрагмент кода недопустим, так как он пытается указать свойство `values` для неограниченного диапазона. Если вы попытаетесь задать свойства уровня ячейки для неограниченного диапазона, API возвратит ошибку.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p122">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large. For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range. The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="df7d0-187">Чтение из большого диапазона и запись в него</span><span class="sxs-lookup"><span data-stu-id="df7d0-187">Read or write to a large range</span></span>
 
<span data-ttu-id="df7d0-p123">Если диапазон содержит большое количество ячеек, значений, форматов чисел или формул, то, возможно, не удастся выполнить операции API над этим диапазоном. API всегда делает все возможное, чтобы выполнить запрошенную операцию над диапазоном (то есть получить или записать указанные данные), но попытка выполнить операцию чтения или записи для большого диапазона может привести к ошибке API из-за чрезмерного потребления ресурсов. Чтобы избежать таких ошибок, мы рекомендуем выполнять отдельные операции чтения или записи для небольших подмножеств большого диапазона, а не пытаться выполнить одну операцию чтения или записи для большого диапазона.</span><span class="sxs-lookup"><span data-stu-id="df7d0-p123">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="df7d0-191">Изменение всех ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="df7d0-191">Update all cells in a range</span></span>
 
<span data-ttu-id="df7d0-192">Если необходимо одинаково изменить все ячейки в диапазоне (например, заполнить все ячейки одним и тем же значением или формулой либо задать один и тот же формат чисел), задайте для соответствующего свойства в объекте **range** (одно) необходимое значение.</span><span class="sxs-lookup"><span data-stu-id="df7d0-192">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="df7d0-193">В примере ниже показано, как получить диапазон, содержащий 20 ячеек, а затем задать формат чисел и заполнить все ячейки в диапазоне значением **3/11/2015** (11.03.2015).</span><span class="sxs-lookup"><span data-stu-id="df7d0-193">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
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
 
##<a name="handle-errors"></a><span data-ttu-id="df7d0-194">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="df7d0-194">Handle Errors</span></span>

<span data-ttu-id="df7d0-195">При возникновении ошибки API возвратит объект **error**, содержащий код и сообщение.</span><span class="sxs-lookup"><span data-stu-id="df7d0-195">When an API error occurs, the API will return an **error** object that contains a code and a message.</span></span> <span data-ttu-id="df7d0-196">Подробные сведения об обработке ошибок, включая список ошибок API, см. в статье [Обработка ошибок](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="df7d0-196">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>
 
## <a name="see-also"></a><span data-ttu-id="df7d0-197">См. также</span><span class="sxs-lookup"><span data-stu-id="df7d0-197">See also</span></span>
 
* [<span data-ttu-id="df7d0-198">Начало работы с надстройками Excel</span><span class="sxs-lookup"><span data-stu-id="df7d0-198">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="df7d0-199">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="df7d0-199">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [<span data-ttu-id="df7d0-200">Углубленные принципы программирования с использованием интерфейса API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="df7d0-200">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
* [<span data-ttu-id="df7d0-201">Оптимизация производительности API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="df7d0-201">Excel JavaScript API performance optimization</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [<span data-ttu-id="df7d0-202">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="df7d0-202">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)

---
ms.date: 05/30/2019
description: Запрос, потоковая передача и отмена потоковой передачи внешних данных к книге с помощью пользовательских функций в Excel
title: Получение и обработка данных с помощью пользовательских функций
localization_priority: Priority
ms.openlocfilehash: 22f79c8b4e7e39569d3b955477e9397a053e1a8f
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910338"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="fdd44-103">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="fdd44-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="fdd44-104">Один из способов, используемых пользовательскими функциями для повышения эффективности Excel, состоит в получении данных из расположений помимо книг, например из Интернета или сервера (через WebSockets).</span><span class="sxs-lookup"><span data-stu-id="fdd44-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="fdd44-105">Пользовательские функции могут запрашивать данные с помощью XHR и запросов `fetch`, а также выполнять потоковую передачу этих данных в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="fdd44-105">Custom functions can request data through XHR and `fetch` requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="fdd44-106">В документах ниже показаны некоторые примеры веб-запросов, но для создания функции потоковой передачи используйте [Руководство по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="fdd44-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="fdd44-107">Функции, которые возвращают данные из внешних источников</span><span class="sxs-lookup"><span data-stu-id="fdd44-107">Functions that return data from external sources</span></span>

<span data-ttu-id="fdd44-108">Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:</span><span class="sxs-lookup"><span data-stu-id="fdd44-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="fdd44-109">Возвращать обещание JavaScript в Excel;</span><span class="sxs-lookup"><span data-stu-id="fdd44-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="fdd44-110">Устранять обещание с итоговым значением с помощью функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fdd44-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="fdd44-111">Можно запрашивать внешние данные с помощью такого API, как [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или с помощью `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="fdd44-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="fdd44-112">В среде выполнения пользовательских функций XHR реализует дополнительные меры по обеспечению безопасности, предъявляя в качестве требования [политику единого домена](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой запрос [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="fdd44-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="fdd44-113">Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="fdd44-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="fdd44-114">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="fdd44-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="fdd44-115">Вы также можете использовать заголовок Content-Type в простом запросе CORS, если используется тип контента `application/x-www-form-urlencoded`, `text/plain` или `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="fdd44-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="fdd44-116">Пример XHR</span><span class="sxs-lookup"><span data-stu-id="fdd44-116">XHR example</span></span>

<span data-ttu-id="fdd44-117">В следующем примере кода функция **getTemperature** вызывает функцию sendWebRequest для получения температуры в определенной области на основе идентификатора термометра.</span><span class="sxs-lookup"><span data-stu-id="fdd44-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="fdd44-118">Функция sendWebRequest использует XHR для отправления запроса GET в конечную точку, которая может предоставить данные.</span><span class="sxs-lookup"><span data-stu-id="fdd44-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```js
/**
 * Receives a temperature from an online source.
 * @customfunction
 * @param {number} thermometerID Identification number of the thermometer.
 */
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions.  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
        };

        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

<span data-ttu-id="fdd44-119">Другой пример запроса XHR с дополнительным контекстом см. в функции `getFile` в [этом файле](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) репозитория Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="fdd44-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="fdd44-120">Пример получения данных</span><span class="sxs-lookup"><span data-stu-id="fdd44-120">Fetch example</span></span>

<span data-ttu-id="fdd44-121">В следующем примере функция `stockPriceStream` использует символ тикера для получения цены акции каждые 1000 миллисекунд.</span><span class="sxs-lookup"><span data-stu-id="fdd44-121">In the following code sample, the `stockPriceStream` function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="fdd44-122">Для получения дополнительных сведений об этом примере см. статью [Руководство по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="fdd44-122">For more details about this sample, see the [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function).</span></span>

```js
/**
 * Streams a stock price.
 * @customfunction 
 * @param {string} ticker Stock ticker.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function stockPriceStream(ticker, invocation) {
    var updateFrequency = 1000 /* milliseconds*/;
    var isPending = false;

    var timer = setInterval(function() {
        // If there is already a pending request, skip this iteration:
        if (isPending) {
            return;
        }

        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        isPending = true;

        fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                invocation.setResult(parseFloat(text));
            })
            .catch(function(error) {
                invocation.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receive-data-via-websockets"></a><span data-ttu-id="fdd44-123">Получение данных через WebSockets</span><span class="sxs-lookup"><span data-stu-id="fdd44-123">Receive data via WebSockets</span></span>

<span data-ttu-id="fdd44-124">В пределах пользовательской функции можно использовать WebSockets для обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="fdd44-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="fdd44-125">С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.</span><span class="sxs-lookup"><span data-stu-id="fdd44-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="fdd44-126">Пример WebSockets</span><span class="sxs-lookup"><span data-stu-id="fdd44-126">WebSockets example</span></span>

<span data-ttu-id="fdd44-127">Следующий примера кода устанавливает соединение WebSocket, а затем заносит в журнал каждое входящее сообщение от сервера.</span><span class="sxs-lookup"><span data-stu-id="fdd44-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="fdd44-128">Создание функции потоковой передачи</span><span class="sxs-lookup"><span data-stu-id="fdd44-128">Make a streaming function</span></span>

<span data-ttu-id="fdd44-129">Пользовательские функции потоковой передачи позволяют выводить данные в ячейки, которые повторно обновляются, не требуя от пользователя явно что-либо обновлять.</span><span class="sxs-lookup"><span data-stu-id="fdd44-129">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="fdd44-130">Такие функции (например, функция из [руководства по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md)) могут быть полезны для проверки данных, обновляемых в реальном времени, из веб-службы.</span><span class="sxs-lookup"><span data-stu-id="fdd44-130">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="fdd44-131">Чтобы объявить функцию потоковой передачи, используйте тег комментария JSDoc `@stream`.</span><span class="sxs-lookup"><span data-stu-id="fdd44-131">To declare a streaming function, use the JSDoc comment tag `@stream`.</span></span> <span data-ttu-id="fdd44-132">Чтобы оповестить пользователей о том, что ваша функция может выполнять повторное вычисление с учетом новой информации, рекомендуем указать поток или другие сведения об этом в имени или описании функции.</span><span class="sxs-lookup"><span data-stu-id="fdd44-132">To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.</span></span>

<span data-ttu-id="fdd44-133">В приведенном ниже примере показана функция потоковой передачи, которая увеличивает определенное число каждую секунду на указанное число.</span><span class="sxs-lookup"><span data-stu-id="fdd44-133">The following example shows a streaming function which increases a given number every second by an amount you specify.</span></span>

```JS
/**
 * Increments a value once a second.
 * @customfunction INC increment
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
CustomFunctions.associate("INC", increment);
```

>[!NOTE]
> <span data-ttu-id="fdd44-134">Обратите внимание, что существует еще одна категория — так называемые отменяемые функции, которые *не* связаны с функциями потоковой передачи.</span><span class="sxs-lookup"><span data-stu-id="fdd44-134">Note that there are also a category of functions called cancelable functions, which are *not* related to streaming functions.</span></span> <span data-ttu-id="fdd44-135">В предыдущих версиях пользовательских функций требовалось объявлять `"cancelable": true` и `"streaming": true` в самостоятельно написанном коде JSON.</span><span class="sxs-lookup"><span data-stu-id="fdd44-135">Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand.</span></span> <span data-ttu-id="fdd44-136">С тех пор, как появились автоматически генерируемые метаданные, можно отменять только асинхронные пользовательские функции, возвращающие одно значение.</span><span class="sxs-lookup"><span data-stu-id="fdd44-136">Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="fdd44-137">Отменяемые функции позволяют прервать выполнение веб-запроса, используя [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation), чтобы решить, что делать после отмены.</span><span class="sxs-lookup"><span data-stu-id="fdd44-137">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="fdd44-138">Для объявления отменяемых функций используется тег `@cancelable`.</span><span class="sxs-lookup"><span data-stu-id="fdd44-138">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="fdd44-139">Использование параметра вызова</span><span class="sxs-lookup"><span data-stu-id="fdd44-139">Using an invocation parameter</span></span>

<span data-ttu-id="fdd44-140">Параметр `invocation` является по умолчанию последним в любой пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="fdd44-140">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="fdd44-141">Параметр `invocation` содержит контекст о ячейке (например, ее адрес), а также позволяет использовать способы `setResult` и `onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="fdd44-141">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="fdd44-142">Эти методы определяют, что делает функция во время ее потоковой передачи (`setResult`) или отмены (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="fdd44-142">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="fdd44-143">При использовании TypeScript требуется обработчик вызовов типа `CustomFunctions.StreamingInvocation` или `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="fdd44-143">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="fdd44-144">Пример потоковой и отменяемой функции</span><span class="sxs-lookup"><span data-stu-id="fdd44-144">Streaming and cancelable function example</span></span>
<span data-ttu-id="fdd44-145">Следующий пример кода — это пользовательская функция, которая добавляет число к результату каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="fdd44-145">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="fdd44-146">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="fdd44-146">Note the following about this code:</span></span>

- <span data-ttu-id="fdd44-147">Excel отображает каждое новое значение автоматически с помощью метода `setResult`.</span><span class="sxs-lookup"><span data-stu-id="fdd44-147">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="fdd44-148">Второй параметр ввода, вызов, не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".</span><span class="sxs-lookup"><span data-stu-id="fdd44-148">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="fdd44-149">Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="fdd44-149">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
    }
}
CustomFunctions.associate("INCREMENT", increment);
```

>[!NOTE]
> <span data-ttu-id="fdd44-150">Excel отменяет выполнение функций в следующих случаях:</span><span class="sxs-lookup"><span data-stu-id="fdd44-150">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="fdd44-151">Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.</span><span class="sxs-lookup"><span data-stu-id="fdd44-151">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="fdd44-152">Когда изменяется один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="fdd44-152">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="fdd44-153">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="fdd44-153">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="fdd44-154">Когда пользователь вручную вызывает пересчет.</span><span class="sxs-lookup"><span data-stu-id="fdd44-154">When the user triggers recalculation manually.</span></span> <span data-ttu-id="fdd44-155">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="fdd44-155">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="fdd44-156">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="fdd44-156">Next steps</span></span>

* <span data-ttu-id="fdd44-157">Ознакомьтесь с [разными типами параметров, которые могут использоваться функциями](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="fdd44-157">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="fdd44-158">Узнайте, как [пакетно обрабатывать несколько вызовов API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="fdd44-158">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fdd44-159">См. также</span><span class="sxs-lookup"><span data-stu-id="fdd44-159">See also</span></span>

* [<span data-ttu-id="fdd44-160">Пересчитываемые значения в функциях</span><span class="sxs-lookup"><span data-stu-id="fdd44-160">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="fdd44-161">Создание метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="fdd44-161">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="fdd44-162">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="fdd44-162">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="fdd44-163">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="fdd44-163">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="fdd44-164">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="fdd44-164">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="fdd44-165">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="fdd44-165">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="fdd44-166">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="fdd44-166">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

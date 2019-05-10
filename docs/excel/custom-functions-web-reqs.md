---
ms.date: 05/07/2019
description: Запрос, потоковая передача и отмена потоковой передачи внешних данных к книге с помощью пользовательских функций в Excel
title: Получение и обработка данных с помощью пользовательских функций
localization_priority: Priority
ms.openlocfilehash: 61f4d0fdaea4277faedddbe075a587fb23842c08
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659637"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="87917-103">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="87917-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="87917-104">Один из способов, используемых пользовательскими функциями для повышения эффективности Excel, состоит в получении данных из расположений помимо книг, например из Интернета или сервера (через WebSockets).</span><span class="sxs-lookup"><span data-stu-id="87917-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="87917-105">Пользовательские функции могут запрашивать данные с помощью XHR и запросов `fetch`, а также выполнять потоковую передачу этих данных в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="87917-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="87917-106">В документах ниже показаны некоторые примеры веб-запросов, но для создания функции потоковой передачи используйте [Руководство по пользовательским функциям](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span><span class="sxs-lookup"><span data-stu-id="87917-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="87917-107">Функции, которые возвращают данные из внешних источников</span><span class="sxs-lookup"><span data-stu-id="87917-107">Functions that return data from external sources</span></span>

<span data-ttu-id="87917-108">Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:</span><span class="sxs-lookup"><span data-stu-id="87917-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="87917-109">Возвращать обещание JavaScript в Excel;</span><span class="sxs-lookup"><span data-stu-id="87917-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="87917-110">Устранять обещание с итоговым значением с помощью функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="87917-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="87917-111">Можно запрашивать внешние данные с помощью такого API, как [`Fetch`](https://developer.mozilla.org/ru-RU/docs/Web/API/Fetch_API), или с помощью `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ru-RU/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="87917-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/ru-RU/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ru-RU/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="87917-112">В среде выполнения пользовательских функций XHR реализует дополнительные меры по обеспечению безопасности, предъявляя в качестве требования [политику единого домена](https://developer.mozilla.org/ru-RU/docs/Web/Security/Same-origin_policy) и простой запрос [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="87917-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/ru-RU/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="87917-113">Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="87917-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="87917-114">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="87917-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="87917-115">Вы также можете использовать заголовок Content-Type в простом запросе CORS, если используется тип контента `application/x-www-form-urlencoded`, `text/plain` или `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="87917-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="87917-116">Пример XHR</span><span class="sxs-lookup"><span data-stu-id="87917-116">XHR example</span></span>

<span data-ttu-id="87917-117">В следующем примере кода функция **getTemperature** вызывает функцию sendWebRequest для получения температуры в определенной области на основе идентификатора термометра.</span><span class="sxs-lookup"><span data-stu-id="87917-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="87917-118">Функция sendWebRequest использует XHR для отправления запроса GET в конечную точку, которая может предоставить данные.</span><span class="sxs-lookup"><span data-stu-id="87917-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

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

<span data-ttu-id="87917-119">Другой пример запроса XHR с дополнительным контекстом см. в функции `getFile` в [этом файле](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) репозитория Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="87917-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="87917-120">Пример получения данных</span><span class="sxs-lookup"><span data-stu-id="87917-120">Fetch example</span></span>

<span data-ttu-id="87917-121">В следующем примере функция `stockPriceStream` использует символ тикера для получения цены акции каждые 1000 миллисекунд.</span><span class="sxs-lookup"><span data-stu-id="87917-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="87917-122">Для получения дополнительных сведений об этом примере см. статью [Руководство по пользовательским функциям](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="87917-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span>

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

## <a name="receive-data-via-websockets"></a><span data-ttu-id="87917-123">Получение данных через WebSockets</span><span class="sxs-lookup"><span data-stu-id="87917-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="87917-124">В пределах пользовательской функции можно использовать WebSockets для обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="87917-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="87917-125">С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.</span><span class="sxs-lookup"><span data-stu-id="87917-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="87917-126">Пример WebSockets</span><span class="sxs-lookup"><span data-stu-id="87917-126">WebSockets example</span></span>

<span data-ttu-id="87917-127">Следующий примера кода устанавливает соединение WebSocket, а затем заносит в журнал каждое входящее сообщение от сервера.</span><span class="sxs-lookup"><span data-stu-id="87917-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="stream-and-cancel-functions"></a><span data-ttu-id="87917-128">Потоковая передача и отмена функций</span><span class="sxs-lookup"><span data-stu-id="87917-128">Stream and cancel functions</span></span>

<span data-ttu-id="87917-129">Потоковая передача пользовательских функций позволяет выводить данные в ячейки, которые повторно обновляются, не требуя от пользователя явно что-либо обновлять.</span><span class="sxs-lookup"><span data-stu-id="87917-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span>

<span data-ttu-id="87917-130">Отменяемые пользовательские функции позволяют отменять выполнение потоковой пользовательской функции, чтобы уменьшить использование пропускной способности, рабочей памяти и загрузку ЦП.</span><span class="sxs-lookup"><span data-stu-id="87917-130">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span>

<span data-ttu-id="87917-131">Чтобы объявить функцию как потоковую или отменяемую, используйте теги комментария JSDOC `@stream` или `@cancelable`.</span><span class="sxs-lookup"><span data-stu-id="87917-131">To declare a function as streaming or cancelable, use the JSDOC comment tags `@stream` or `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="87917-132">Использование параметра вызова</span><span class="sxs-lookup"><span data-stu-id="87917-132">Using an invocation parameter</span></span>

<span data-ttu-id="87917-133">Параметр `invocation` является по умолчанию последним в любой пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="87917-133">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="87917-134">Параметр `invocation` содержит контекст о ячейке (например, ее адрес), а также позволяет использовать способы `setResult` и `onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="87917-134">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="87917-135">Эти методы определяют, что делает функция во время ее потоковой передачи (`setResult`) или отмены (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="87917-135">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="87917-136">При использовании TypeScript требуется обработчик вызовов типа `CustomFunctions.StreamingInvocation` или `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="87917-136">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="87917-137">Пример потоковой и отменяемой функции</span><span class="sxs-lookup"><span data-stu-id="87917-137">Streaming and cancelable function example</span></span>
<span data-ttu-id="87917-138">Следующий пример кода — это пользовательская функция, которая добавляет число к результату каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="87917-138">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="87917-139">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="87917-139">Note the following about this code:</span></span>

- <span data-ttu-id="87917-140">Excel отображает каждое новое значение автоматически с помощью метода `setResult`.</span><span class="sxs-lookup"><span data-stu-id="87917-140">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="87917-141">Второй параметр ввода, вызов, не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".</span><span class="sxs-lookup"><span data-stu-id="87917-141">The second input parameter, , is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="87917-142">Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="87917-142">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

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
> <span data-ttu-id="87917-143">Excel отменяет выполнение функций в следующих случаях:</span><span class="sxs-lookup"><span data-stu-id="87917-143">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="87917-144">Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.</span><span class="sxs-lookup"><span data-stu-id="87917-144">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="87917-145">Когда изменяется один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="87917-145">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="87917-146">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="87917-146">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="87917-147">Когда пользователь вручную вызывает пересчет.</span><span class="sxs-lookup"><span data-stu-id="87917-147">When the user triggers recalculation manually.</span></span> <span data-ttu-id="87917-148">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="87917-148">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="87917-149">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="87917-149">Next steps</span></span>

* <span data-ttu-id="87917-150">Ознакомьтесь с [разными типами параметров, которые могут использоваться функциями](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="87917-150">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="87917-151">Узнайте, как [пакетно обрабатывать несколько вызовов API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="87917-151">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="87917-152">См. также</span><span class="sxs-lookup"><span data-stu-id="87917-152">See also</span></span>

* [<span data-ttu-id="87917-153">Пересчитываемые значения в функциях</span><span class="sxs-lookup"><span data-stu-id="87917-153">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="87917-154">Создание метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="87917-154">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="87917-155">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="87917-155">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="87917-156">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="87917-156">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="87917-157">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="87917-157">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="87917-158">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="87917-158">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="87917-159">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="87917-159">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

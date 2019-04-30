---
ms.date: 04/20/2019
description: Запрос, потоковая передача и отмена потоковой передачи внешних данных к книге с помощью пользовательских функций в Excel
title: Обработка веб-запросов и других данных с помощью пользовательских функций (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: 2942ec56e46d6eb586b516eedab17c1eeb98d9c8
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353267"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a><span data-ttu-id="34853-103">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="34853-103">Receiving and handling data with custom functions</span></span>

<span data-ttu-id="34853-104">Один из способов, используемых пользовательскими функциями для повышения эффективности Excel, состоит в получении данных из расположений помимо книг, например из Интернета или сервера (через WebSockets).</span><span class="sxs-lookup"><span data-stu-id="34853-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="34853-105">Пользовательские функции могут запрашивать данные с помощью XHR и получать запросы, а также выполнять потоковую передачу этих данных в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="34853-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

<span data-ttu-id="34853-106">В документах ниже показаны некоторые примеры веб-запросов, но для создания функции потоковой передачи используйте [Руководство по пользовательским функциям](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span><span class="sxs-lookup"><span data-stu-id="34853-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="34853-107">Функции, которые возвращают данные из внешних источников</span><span class="sxs-lookup"><span data-stu-id="34853-107">Functions that return data from external sources</span></span>

<span data-ttu-id="34853-108">Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:</span><span class="sxs-lookup"><span data-stu-id="34853-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="34853-109">Возвращать обещание JavaScript в Excel;</span><span class="sxs-lookup"><span data-stu-id="34853-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="34853-110">Устранять обещание с итоговым значением с помощью функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="34853-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="34853-111">Можно запрашивать внешние данные с помощью такого API, как [`Fetch`](https://developer.mozilla.org/ru-RU/docs/Web/API/Fetch_API), или с помощью `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ru-RU/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="34853-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/ru-RU/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ru-RU/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="34853-112">В среде выполнения пользовательских функций XHR реализует дополнительные меры по обеспечению безопасности, предъявляя в качестве требования [политику единого домена](https://developer.mozilla.org/ru-RU/docs/Web/Security/Same-origin_policy) и простой запрос [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="34853-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/ru-RU/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="34853-113">Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="34853-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="34853-114">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="34853-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="34853-115">Вы также можете использовать заголовок Content-Type в простом запросе CORS, если используется тип контента `application/x-www-form-urlencoded`, `text/plain` или `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="34853-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="34853-116">Пример XHR</span><span class="sxs-lookup"><span data-stu-id="34853-116">XHR example</span></span>

<span data-ttu-id="34853-117">В следующем примере кода функция **getTemperature** вызывает функцию sendWebRequest для получения температуры в определенной области на основе идентификатора термометра.</span><span class="sxs-lookup"><span data-stu-id="34853-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="34853-118">Функция sendWebRequest использует XHR для отправления запроса GET в конечную точку, которая может предоставить данные.</span><span class="sxs-lookup"><span data-stu-id="34853-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```JavaScript
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ 
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
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

<span data-ttu-id="34853-119">Другой пример запроса XHR с дополнительным контекстом см. в функции `getFile` в [этом файле](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) репозитория Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="34853-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="34853-120">Пример получения данных</span><span class="sxs-lookup"><span data-stu-id="34853-120">Fetch example</span></span>

<span data-ttu-id="34853-121">В следующем примере функция stockPriceStream использует символ тикера для получения цены акции каждые 1000 миллисекунд.</span><span class="sxs-lookup"><span data-stu-id="34853-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="34853-122">Для получения дополнительных сведений об этом примере и соответствующего файла JSON см. статью [Руководство по пользовательским функциям](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="34853-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span> 

```JavaScript
function stockPriceStream(ticker, handler) {
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
                handler.setResult(parseFloat(text));
            })
            .catch(function(error) {
                handler.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    handler.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="34853-123">Получение данных через WebSockets</span><span class="sxs-lookup"><span data-stu-id="34853-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="34853-124">В пределах пользовательской функции можно использовать WebSockets для обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="34853-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="34853-125">С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.</span><span class="sxs-lookup"><span data-stu-id="34853-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="34853-126">Пример WebSockets</span><span class="sxs-lookup"><span data-stu-id="34853-126">WebSockets example</span></span>

<span data-ttu-id="34853-127">Следующий примера кода устанавливает соединение WebSocket, а затем заносит в журнал каждое входящее сообщение от сервера.</span><span class="sxs-lookup"><span data-stu-id="34853-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a><span data-ttu-id="34853-128">Потоковая передача функций</span><span class="sxs-lookup"><span data-stu-id="34853-128">Streaming functions</span></span>

<span data-ttu-id="34853-129">Потоковая передача пользовательских функций позволяет выводить данные в ячейки несколько раз с течением времени, избавляя пользователя от необходимости явным образом запрашивать обновление данных.</span><span class="sxs-lookup"><span data-stu-id="34853-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="34853-130">Следующий пример кода — это пользовательская функция, которая добавляет число к результату каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="34853-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="34853-131">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="34853-131">Note the following about this code:</span></span>

- <span data-ttu-id="34853-132">Excel отображает каждое новое значение автоматически с помощью обратного вызова setResult.</span><span class="sxs-lookup"><span data-stu-id="34853-132">Excel displays each new value automatically using the setResult callback.</span></span>
- <span data-ttu-id="34853-133">Второй параметр ввода (handler) не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".</span><span class="sxs-lookup"><span data-stu-id="34853-133">The second input parameter, handler, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="34853-134">Обратный вызов onCanceled определяет функцию, которая выполняется при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="34853-134">The onCanceled callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="34853-135">Вам необходимо реализовать уведомление об отмене следующим образом для любой функции потоковой передачи.</span><span class="sxs-lookup"><span data-stu-id="34853-135">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="34853-136">Дополнительные сведения см. в разделе [Отмена функции](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="34853-136">For more information, see [Canceling a function](#canceling-a-function).</span></span>

```JavaScript
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}

CustomFunctions.associate("INCREMENTVALUE", incrementValue);
```

<span data-ttu-id="34853-137">Когда вы указываете метаданные для функции потоковой передачи в файле метаданных JSON, это можно автоматически создать с помощью тега комментария JSDOC `@streaming` в файле скрипта функции.</span><span class="sxs-lookup"><span data-stu-id="34853-137">When you specify metadata for a streaming function in the JSON metadata file, you can autogenerate this by using a `@streaming` JSDOC comment tag in your function's script file.</span></span> <span data-ttu-id="34853-138">Дополнительные сведения см. в статье [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="34853-138">For more details, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="canceling-a-function"></a><span data-ttu-id="34853-139">Отмена функции</span><span class="sxs-lookup"><span data-stu-id="34853-139">Canceling a function</span></span>

<span data-ttu-id="34853-140">В некоторых случаях может потребоваться отмена выполнения пользовательских функций потоковой передачи, чтобы уменьшить использования пропускной способности, рабочей памяти и загрузку ЦП.</span><span class="sxs-lookup"><span data-stu-id="34853-140">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="34853-141">Excel отменяет выполнение функций в следующих случаях:</span><span class="sxs-lookup"><span data-stu-id="34853-141">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="34853-142">Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.</span><span class="sxs-lookup"><span data-stu-id="34853-142">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="34853-143">Когда изменяется один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="34853-143">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="34853-144">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="34853-144">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="34853-145">Когда пользователь вручную вызывает пересчет.</span><span class="sxs-lookup"><span data-stu-id="34853-145">When the user triggers recalculation manually.</span></span> <span data-ttu-id="34853-146">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="34853-146">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="34853-147">Чтобы сделать функцию отменяемой, нужно реализовать обработчик в коде функции с указанием действий при ее отмене.</span><span class="sxs-lookup"><span data-stu-id="34853-147">To make a function cancelable, implement a handler in your function's code to tell it what to do when it is canceled.</span></span> <span data-ttu-id="34853-148">Также можно использовать тег комментария JSDOC `@cancelable` в файле скрипта функции.</span><span class="sxs-lookup"><span data-stu-id="34853-148">Additionally, use the `@cancelable` JSDOC comment tag in your function's script file.</span></span> <span data-ttu-id="34853-149">Дополнительные сведения см. в статье [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="34853-149">For more details, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="34853-150">См. также</span><span class="sxs-lookup"><span data-stu-id="34853-150">See also</span></span>

* [<span data-ttu-id="34853-151">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="34853-151">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="34853-152">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="34853-152">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="34853-153">Создание метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="34853-153">Create JSON metadata for custom functions (preview)</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="34853-154">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="34853-154">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="34853-155">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="34853-155">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="34853-156">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="34853-156">Custom functions changelog</span></span>](custom-functions-changelog.md)

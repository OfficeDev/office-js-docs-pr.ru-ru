---
ms.date: 01/14/2020
description: Запрос, потоковая передача и отмена потоковой передачи внешних данных к книге с помощью пользовательских функций в Excel
title: Получение и обработка данных с помощью пользовательских функций
localization_priority: Normal
ms.openlocfilehash: c59133389311721c4d0507337a17b346af97e690
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41949717"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="3f920-103">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="3f920-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="3f920-104">Один из способов, используемых пользовательскими функциями для повышения эффективности Excel, состоит в получении данных из расположений помимо книг, например из Интернета или сервера (через WebSockets).</span><span class="sxs-lookup"><span data-stu-id="3f920-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="3f920-105">Можно запрашивать внешние данные с помощью такого API, как [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или с помощью `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="3f920-105">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![GIF с пользовательской функцией, отправляющей время из API](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="3f920-107">Функции, которые возвращают данные из внешних источников</span><span class="sxs-lookup"><span data-stu-id="3f920-107">Functions that return data from external sources</span></span>

<span data-ttu-id="3f920-108">Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:</span><span class="sxs-lookup"><span data-stu-id="3f920-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="3f920-109">Возвращать обещание JavaScript в Excel;</span><span class="sxs-lookup"><span data-stu-id="3f920-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="3f920-110">Устранять обещание с итоговым значением с помощью функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="3f920-110">Resolve the Promise with the final value using the callback function.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="3f920-111">Пример получения данных</span><span class="sxs-lookup"><span data-stu-id="3f920-111">Fetch example</span></span>

<span data-ttu-id="3f920-112">В следующем примере кода функция **webRequest** обращается к гипотетическому API "Число людей в космосе" компании Contoso, которое отслеживает число людей, находящихся в настоящее время на Международной космической станции.</span><span class="sxs-lookup"><span data-stu-id="3f920-112">In the following code sample, the **webRequest** function reaches out to the hypothetical Contoso "Number of People in Space" API, which tracks the number of people currently on the International Space Station.</span></span> <span data-ttu-id="3f920-113">Функция возвращает обещание JavaScript и использует метод Fetch для запроса сведений из API.</span><span class="sxs-lookup"><span data-stu-id="3f920-113">The function returns a JavaScript Promise and uses fetch to request information from the API.</span></span> <span data-ttu-id="3f920-114">Итоговые данные преобразуются в формат JSON, а свойство `names` преобразуется в строку, использующуюся для разрешения обещания.</span><span class="sxs-lookup"><span data-stu-id="3f920-114">The resulting data is transformed into JSON and the `names` property is converted into a string, which is used to resolve the Promise.</span></span>

<span data-ttu-id="3f920-115">При разработке собственных функций может потребоваться выполнение действия, если веб-запрос не завершается своевременно. Также можно рассмотреть [совмещение нескольких запросов API](./custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="3f920-115">When developing your own functions, you may want to perform an action if the web request does not complete in a timely manner or consider [batching up multiple API requests](./custom-functions-batching.md).</span></span>

```JS
/**
 * Requests the names of the people currently on the International Space Station from a hypothetical API.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace";
  return new Promise(function (resolve, reject) {
    fetch(url)
      .then(function (response){
        return response.json();
        }
      )
      .then(function (json) {
        resolve(JSON.stringify(json.names));
      })
  })
}
```

>[!NOTE]
><span data-ttu-id="3f920-116">При использовании метода `Fetch` не создаются вложенные обратные вызовы, что в некоторых случаях может быть предпочтительнее, чем использование метода XHR.</span><span class="sxs-lookup"><span data-stu-id="3f920-116">Using `Fetch` avoids nested callbacks and may be preferable to XHR in some cases.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="3f920-117">Пример XHR</span><span class="sxs-lookup"><span data-stu-id="3f920-117">XHR example</span></span>

<span data-ttu-id="3f920-118">В среде выполнения пользовательских функций XHR реализует дополнительные меры по обеспечению безопасности, предъявляя в качестве требования [политику единого домена](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой запрос [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="3f920-118">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="3f920-119">Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="3f920-119">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="3f920-120">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="3f920-120">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="3f920-121">Вы также можете использовать заголовок Content-Type в простом запросе CORS, если используется тип контента `application/x-www-form-urlencoded`, `text/plain` или `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="3f920-121">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

<span data-ttu-id="3f920-122">В следующем примере кода функция **getStarCount** вызывает API GitHub для определения числа звезд, присвоенных репозиторию определенного пользователя.</span><span class="sxs-lookup"><span data-stu-id="3f920-122">In the following code sample, the **getStarCount** function calls the Github API to discover the amount of stars given to a particular user's repository.</span></span> <span data-ttu-id="3f920-123">Это асинхронная функция, возвращающая обещание JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3f920-123">This is an asynchronous function which returns a JavaScript Promise.</span></span> <span data-ttu-id="3f920-124">При получении данных из веб-вызова обещание разрешается, что возвращает данные в ячейку.</span><span class="sxs-lookup"><span data-stu-id="3f920-124">When data is obtained from the web call, the Promise is resolved which returns the data to the cell.</span></span>

```TS
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */

async function getStarCount(userName: string, repoName: string) {

  const url = "https://api.github.com/repos/" + userName + "/" + repoName;

  let xhttp = new XMLHttpRequest();

  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        resolve(JSON.parse(xhttp.responseText).watchers_count);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="3f920-125">Создание функции потоковой передачи</span><span class="sxs-lookup"><span data-stu-id="3f920-125">Make a streaming function</span></span>

<span data-ttu-id="3f920-126">Пользовательские функции потоковой передачи позволяют выводить данные в ячейки, которые повторно обновляются, не требуя от пользователя явно что-либо обновлять.</span><span class="sxs-lookup"><span data-stu-id="3f920-126">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="3f920-127">Такие функции (например, функция из [руководства по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md)) могут быть полезны для проверки данных, обновляемых в реальном времени, из веб-службы.</span><span class="sxs-lookup"><span data-stu-id="3f920-127">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="3f920-128">Чтобы объявить функцию потоковой передачи, используйте тег `@streaming` или параметр вызова `CustomFunctions.StreamingInvocation`. Это указывает, что это функция потоковой передачи.</span><span class="sxs-lookup"><span data-stu-id="3f920-128">To declare a streaming function, either use the tag `@streaming` or make use of the `CustomFunctions.StreamingInvocation` invocation parameter, which will indicate that your function is streaming.</span></span> <span data-ttu-id="3f920-129">Чтобы оповестить пользователей о том, что ваша функция может выполнять повторное вычисление с учетом новой информации, рекомендуем указать поток или другие сведения об этом в имени или описании функции.</span><span class="sxs-lookup"><span data-stu-id="3f920-129">To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.</span></span>

<span data-ttu-id="3f920-130">Следующий пример кода — это пользовательская функция, которая добавляет число к результату каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="3f920-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="3f920-131">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="3f920-131">Note the following about this code:</span></span>

- <span data-ttu-id="3f920-132">Excel отображает каждое новое значение автоматически с помощью метода `setResult`.</span><span class="sxs-lookup"><span data-stu-id="3f920-132">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="3f920-133">Второй параметр ввода, вызов, не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".</span><span class="sxs-lookup"><span data-stu-id="3f920-133">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="3f920-134">Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="3f920-134">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>
- <span data-ttu-id="3f920-135">Потоковая передача не обязательно связана с веб-запросом: в этом случае функция не выполняет веб-запрос, но по-прежнему получает данные через заданные интервалы, поэтому для нее требуется использовать параметр потоковой передачи `invocation`.</span><span class="sxs-lookup"><span data-stu-id="3f920-135">Streaming isn't necessarily tied to making a web request: in this case, the function isn't making a web request but is still getting data at set intervals, so it requires the use of the streaming `invocation` parameter.</span></span>

```js
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
```

<span data-ttu-id="3f920-136">В дополнение к сведениям об обратном вызове `onCanceled` вы также должны учитывать, что Excel отменяет выполнение функции в следующих случаях:</span><span class="sxs-lookup"><span data-stu-id="3f920-136">In addition to knowing about the `onCanceled` callback, you should also know that Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="3f920-137">Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.</span><span class="sxs-lookup"><span data-stu-id="3f920-137">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="3f920-138">Когда изменяется один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="3f920-138">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="3f920-139">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="3f920-139">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="3f920-140">Когда пользователь вручную вызывает пересчет.</span><span class="sxs-lookup"><span data-stu-id="3f920-140">When the user triggers recalculation manually.</span></span> <span data-ttu-id="3f920-141">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="3f920-141">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="3f920-142">Также можно настроить стандартное значение потоковой передачи, чтобы обрабатывать случаи выполнения запроса, когда вы находитесь в автономном режиме.</span><span class="sxs-lookup"><span data-stu-id="3f920-142">You can also consider setting a default streaming value to handle cases when a request is made but you are offline.</span></span>

> [!NOTE]
> <span data-ttu-id="3f920-143">Обратите внимание, что существует еще одна категория — так называемые отменяемые функции, которые _не_ связаны с функциями потоковой передачи.</span><span class="sxs-lookup"><span data-stu-id="3f920-143">Note that there are also a category of functions called cancelable functions, which are _not_ related to streaming functions.</span></span> <span data-ttu-id="3f920-144">В предыдущих версиях пользовательских функций требовалось объявлять `"cancelable": true` и `"streaming": true` в самостоятельно написанном коде JSON.</span><span class="sxs-lookup"><span data-stu-id="3f920-144">Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand.</span></span> <span data-ttu-id="3f920-145">С тех пор, как появились автоматически генерируемые метаданные, можно отменять только асинхронные пользовательские функции, возвращающие одно значение.</span><span class="sxs-lookup"><span data-stu-id="3f920-145">Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="3f920-146">Отменяемые функции позволяют прервать выполнение веб-запроса, используя [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation), чтобы решить, что делать после отмены.</span><span class="sxs-lookup"><span data-stu-id="3f920-146">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="3f920-147">Для объявления отменяемых функций используется тег `@cancelable`.</span><span class="sxs-lookup"><span data-stu-id="3f920-147">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="3f920-148">Использование параметра вызова</span><span class="sxs-lookup"><span data-stu-id="3f920-148">Using an invocation parameter</span></span>

<span data-ttu-id="3f920-149">Параметр `invocation` является по умолчанию последним в любой пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="3f920-149">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="3f920-150">Параметр `invocation` содержит контекст о ячейке (например, ее адрес и содержимое), а также позволяет использовать методы `setResult` и `onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="3f920-150">The `invocation` parameter gives context about the cell (such as its address and contents) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="3f920-151">Эти методы определяют, что делает функция во время ее потоковой передачи (`setResult`) или отмены (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="3f920-151">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="3f920-152">При использовании TypeScript требуется обработчик вызовов типа `CustomFunctions.StreamingInvocation` или `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="3f920-152">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

## <a name="receive-data-via-websockets"></a><span data-ttu-id="3f920-153">Получение данных через WebSockets</span><span class="sxs-lookup"><span data-stu-id="3f920-153">Receive data via WebSockets</span></span>

<span data-ttu-id="3f920-154">В пределах пользовательской функции можно использовать WebSockets для обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="3f920-154">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="3f920-155">С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.</span><span class="sxs-lookup"><span data-stu-id="3f920-155">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="3f920-156">Пример WebSockets</span><span class="sxs-lookup"><span data-stu-id="3f920-156">WebSockets example</span></span>

<span data-ttu-id="3f920-157">Следующий примера кода устанавливает соединение WebSocket, а затем заносит в журнал каждое входящее сообщение от сервера.</span><span class="sxs-lookup"><span data-stu-id="3f920-157">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="next-steps"></a><span data-ttu-id="3f920-158">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="3f920-158">Next steps</span></span>

- <span data-ttu-id="3f920-159">Ознакомьтесь с [разными типами параметров, которые могут использоваться функциями](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="3f920-159">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
- <span data-ttu-id="3f920-160">Узнайте, как [пакетно обрабатывать несколько вызовов API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="3f920-160">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3f920-161">См. также</span><span class="sxs-lookup"><span data-stu-id="3f920-161">See also</span></span>

- [<span data-ttu-id="3f920-162">Пересчитываемые значения в функциях</span><span class="sxs-lookup"><span data-stu-id="3f920-162">Volatile values in functions</span></span>](custom-functions-volatile.md)
- [<span data-ttu-id="3f920-163">Создание метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="3f920-163">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="3f920-164">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="3f920-164">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="3f920-165">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="3f920-165">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
- [<span data-ttu-id="3f920-166">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="3f920-166">Create custom functions in Excel</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="3f920-167">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="3f920-167">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

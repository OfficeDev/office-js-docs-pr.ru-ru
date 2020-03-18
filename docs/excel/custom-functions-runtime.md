---
ms.date: 05/08/2019
description: Сведения об основных сценариях при разработке пользовательских функций Excel, которые используют новую среду выполнения JavaScript.
title: Среда выполнения для пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: 2cb950cd6f5f78ed76b19a1fa443720d7cfb86a2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719499"
---
# <a name="runtime-for-excel-custom-functions"></a><span data-ttu-id="f3c1a-103">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="f3c1a-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="f3c1a-104">Пользовательские функции используют новую среду выполнения JavaScript, отличающимся от среды выполнения, используемой другими частями надстройки, такими как область задач или другие элементы пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="f3c1a-105">Эта среда выполнения JavaScript предназначена для оптимизации производительности вычислений в пользовательских функциях и представляет новые API, которые можно использовать для выполнения стандартных действий в Интернете в пределах пользовательских функций, например для отправления запроса внешних данных или обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f3c1a-106">Среда выполнения JavaScript также обеспечивает доступ к новым API в пространстве имен `OfficeRuntime`, которые могут быть использованы в пределах пользовательских функций или другими частями надстройки для хранения данных или отображения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="f3c1a-107">В этой статье объясняется, как использовать такие API в пределах пользовательских функций, а также приводятся другие важные замечания, которые следует учитывать при разработке пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="f3c1a-108">Запрос внешних данных</span><span class="sxs-lookup"><span data-stu-id="f3c1a-108">Requesting external data</span></span>

<span data-ttu-id="f3c1a-109">В пределах пользовательской функции можно запрашивать внешние данные с помощью такого API, как [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или с помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="f3c1a-110">В среде выполнения JavaScript, используемой пользовательскими функциями, XHR реализует дополнительные меры безопасности, требуя [одного и того же политики начала](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="f3c1a-110">Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="f3c1a-111">Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="f3c1a-111">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="f3c1a-112">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-112">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="f3c1a-113">Вы также можете `Content-Type` использовать заголовок в простой CORS, при условии, что тип контента `application/x-www-form-urlencoded`: `text/plain`, или `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-113">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="f3c1a-114">Пример XHR</span><span class="sxs-lookup"><span data-stu-id="f3c1a-114">XHR example</span></span>

<span data-ttu-id="f3c1a-115">В приведенном ниже примере кода функция `getTemperature` вызывает функцию `sendWebRequest` для получения температуры в определенной области на основе идентификатора термометра.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-115">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="f3c1a-116">Функция `sendWebRequest` использует XHR для отправления запроса `GET` в конечную точку, которая может предоставить данные.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-116">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>

> [!NOTE] 
> <span data-ttu-id="f3c1a-117">При использовании Fetch или XHR возвращается новое значение `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-117">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="f3c1a-118">До сентября 2018 г. необходимо было указывать `OfficeExtension.Promise` использовать обещания в пределах API Office JavaScript, но теперь вы можете просто использовать JavaScript `Promise`.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-118">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

```js
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
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="f3c1a-119">Получение данных через WebSockets</span><span class="sxs-lookup"><span data-stu-id="f3c1a-119">Receiving data via WebSockets</span></span>

<span data-ttu-id="f3c1a-120">В пределах пользовательской функции можно использовать [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) для обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-120">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="f3c1a-121">С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-121">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="f3c1a-122">Пример WebSockets</span><span class="sxs-lookup"><span data-stu-id="f3c1a-122">WebSockets example</span></span>

<span data-ttu-id="f3c1a-123">Приведенный ниже примера кода устанавливает соединение `WebSocket`, а затем заносит в журнал каждое входящее сообщение от сервера.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-123">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span>

```js
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="f3c1a-124">Хранения данных и доступ к ним</span><span class="sxs-lookup"><span data-stu-id="f3c1a-124">Storing and accessing data</span></span>

<span data-ttu-id="f3c1a-125">В пределах функции (или в пределах любой другой части надстройки) можно хранить данные и выполнять доступ к ним с помощью объекта `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-125">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="f3c1a-126">`Storage` — это постоянная незашифрованная система-хранилище пары "ключ-значение", обеспечивающая альтернативу хранилищу [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), которое нельзя использовать в пределах пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-126">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="f3c1a-127">`Storage`предоставляет 10 МБ данных для каждого домена.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-127">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="f3c1a-128">Домены могут совместно использоваться несколькими надстройками.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-128">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="f3c1a-129">`Storage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-129">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="f3c1a-130">Например, токены для аутентификации пользователей могут храниться в `storage`, так как доступ к нему могут выполнять и пользовательская функция, и элементы пользовательского интерфейса надстройки, такие как область задач.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-130">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="f3c1a-131">Точно так же, если две надстройки используют один и тот же домен (например, www.contoso.com/addin1, www.contoso.com/addin2), им также разрешается обмен информацией в оба направления через `storage`.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-131">Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="f3c1a-132">Обратите внимание, что надстройки, имеющие разные поддомены, будут иметь разные экземпляры `storage` (например, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span><span class="sxs-lookup"><span data-stu-id="f3c1a-132">Note that add-ins which have different subdomains will have different instances of `storage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span></span>

<span data-ttu-id="f3c1a-133">Так как `storage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".</span><span class="sxs-lookup"><span data-stu-id="f3c1a-133">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="f3c1a-134">Ниже указаны методы, доступные в объекте `storage`.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-134">The following methods are available on the `storage` object:</span></span>

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

<span data-ttu-id="f3c1a-135">.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-135">.</span></span>[!NOTE]
> <span data-ttu-id="f3c1a-136">Нет метода для очистки всей информации (например, `clear`).</span><span class="sxs-lookup"><span data-stu-id="f3c1a-136">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="f3c1a-137">Вместо этого вам следует использовать `removeItems` для одновременного удаления нескольких записей.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-137">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="f3c1a-138">Пример Оффицерунтиме. Storage</span><span class="sxs-lookup"><span data-stu-id="f3c1a-138">OfficeRuntime.storage example</span></span>

<span data-ttu-id="f3c1a-139">В следующем примере кода вызывается `OfficeRuntime.storage.setItem` функция для установки ключа и значения `storage`.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-139">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="f3c1a-140">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="f3c1a-140">Additional considerations</span></span>

<span data-ttu-id="f3c1a-141">Чтобы создать надстройку, которая будет работать на различных платформах (один из основных клиентов надстроек Office), вам не следует выполнять доступ к модели DOM в пользовательских функциях или использовать библиотеки, такие как jQuery, которые используют модель DOM.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-141">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="f3c1a-142">В Excel для Windows, где пользовательские функции используют среду выполнения JavaScript, пользовательские функции не могут получить доступ к модели DOM.</span><span class="sxs-lookup"><span data-stu-id="f3c1a-142">In Excel on Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f3c1a-143">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="f3c1a-143">Next steps</span></span>
<span data-ttu-id="f3c1a-144">Узнайте, как [выполнять веб-запросы с пользовательскими функциями](custom-functions-web-reqs.md).</span><span class="sxs-lookup"><span data-stu-id="f3c1a-144">Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f3c1a-145">См. также</span><span class="sxs-lookup"><span data-stu-id="f3c1a-145">See also</span></span>

* [<span data-ttu-id="f3c1a-146">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="f3c1a-146">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f3c1a-147">Архитектура пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f3c1a-147">Custom functions architecture</span></span>](custom-functions-architecture.md)
* [<span data-ttu-id="f3c1a-148">Отображение диалогового окна в пользовательских функциях</span><span class="sxs-lookup"><span data-stu-id="f3c1a-148">Display a dialog in custom functions</span></span>](custom-functions-dialog.md)
* [<span data-ttu-id="f3c1a-149">Руководство по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="f3c1a-149">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

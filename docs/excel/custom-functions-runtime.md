---
ms.date: 02/06/2019
description: Сведения об основных сценариях при разработке пользовательских функций Excel, которые используют новую среду выполнения JavaScript.
title: Среда выполнения для пользовательских функций Excel (предварительный просмотр)
localization_priority: Normal
ms.openlocfilehash: 85024b6c3559e2a5f32bae9297787f8052bba38d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448219"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="788ab-103">Среда выполнения для пользовательских функций Excel (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="788ab-103">Runtime for Excel custom functions (preview)</span></span>

<span data-ttu-id="788ab-104">Пользовательские функции используют новую среду выполнения JavaScript, отличающимся от среды выполнения, используемой другими частями надстройки, такими как область задач или другие элементы пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="788ab-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="788ab-105">Эта среда выполнения JavaScript предназначена для оптимизации производительности вычислений в пользовательских функциях и представляет новые API, которые можно использовать для выполнения стандартных действий в Интернете в пределах пользовательских функций, например для отправления запроса внешних данных или обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="788ab-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="788ab-106">Среда выполнения JavaScript также обеспечивает доступ к новым API в пространстве имен `OfficeRuntime`, которые могут быть использованы в пределах пользовательских функций или другими частями надстройки для хранения данных или отображения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="788ab-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="788ab-107">В этой статье объясняется, как использовать такие API в пределах пользовательских функций, а также приводятся другие важные замечания, которые следует учитывать при разработке пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="788ab-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="788ab-108">Запрос внешних данных</span><span class="sxs-lookup"><span data-stu-id="788ab-108">Requesting external data</span></span>

<span data-ttu-id="788ab-109">В пределах пользовательской функции можно запрашивать внешние данные с помощью такого API, как [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или с помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="788ab-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="788ab-110">В среде выполнения JavaScript, используемой пользовательскими функциями, XHR реализует дополнительные меры безопасности, требуя [одного и того же политики начала](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="788ab-110">Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="788ab-111">Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="788ab-111">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="788ab-112">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="788ab-112">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="788ab-113">Вы также можете `Content-Type` использовать заголовок в простой CORS, при условии, что тип контента `application/x-www-form-urlencoded`: `text/plain`, или `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="788ab-113">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="788ab-114">Пример XHR</span><span class="sxs-lookup"><span data-stu-id="788ab-114">XHR example</span></span>

<span data-ttu-id="788ab-115">В приведенном ниже примере кода функция `getTemperature` вызывает функцию `sendWebRequest` для получения температуры в определенной области на основе идентификатора термометра.</span><span class="sxs-lookup"><span data-stu-id="788ab-115">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="788ab-116">Функция `sendWebRequest` использует XHR для отправления запроса `GET` в конечную точку, которая может предоставить данные.</span><span class="sxs-lookup"><span data-stu-id="788ab-116">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>

> [!NOTE] 
> <span data-ttu-id="788ab-117">При использовании Fetch или XHR возвращается новое значение `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="788ab-117">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="788ab-118">До сентября 2018 г. необходимо было указывать `OfficeExtension.Promise` использовать обещания в пределах API Office JavaScript, но теперь вы можете просто использовать JavaScript `Promise`.</span><span class="sxs-lookup"><span data-stu-id="788ab-118">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="788ab-119">Получение данных через WebSockets</span><span class="sxs-lookup"><span data-stu-id="788ab-119">Receiving data via WebSockets</span></span>

<span data-ttu-id="788ab-120">В пределах пользовательской функции можно использовать [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) для обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="788ab-120">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="788ab-121">С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.</span><span class="sxs-lookup"><span data-stu-id="788ab-121">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="788ab-122">Пример WebSockets</span><span class="sxs-lookup"><span data-stu-id="788ab-122">WebSockets example</span></span>

<span data-ttu-id="788ab-123">Приведенный ниже примера кода устанавливает соединение `WebSocket`, а затем заносит в журнал каждое входящее сообщение от сервера.</span><span class="sxs-lookup"><span data-stu-id="788ab-123">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="788ab-124">Хранения данных и доступ к ним</span><span class="sxs-lookup"><span data-stu-id="788ab-124">Storing and accessing data</span></span>

<span data-ttu-id="788ab-125">В пределах функции (или в пределах любой другой части надстройки) можно хранить данные и выполнять доступ к ним с помощью объекта `OfficeRuntime.AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="788ab-125">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="788ab-126">`AsyncStorage` — это постоянная незашифрованная система-хранилище пары "ключ-значение", обеспечивающая альтернативу хранилищу [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), которое нельзя использовать в пределах пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="788ab-126">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="788ab-127">Надстройка может хранить до 10 МБ данных, используя `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="788ab-127">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="788ab-128">`AsyncStorage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным.</span><span class="sxs-lookup"><span data-stu-id="788ab-128">`AsyncStorage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="788ab-129">Например, токены для аутентификации пользователей могут храниться в `AsyncStorage`, так как доступ к нему могут выполнять и пользовательская функция, и элементы пользовательского интерфейса надстройки, такие как область задач.</span><span class="sxs-lookup"><span data-stu-id="788ab-129">For example, tokens for user authentication may be stored in `AsyncStorage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="788ab-130">Точно так же, если две надстройки используют один и тот же домен (например, www.contoso.com/addin1, www.contoso.com/addin2), им также разрешается обмен информацией в оба направления через `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="788ab-130">Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `AsyncStorage`.</span></span> <span data-ttu-id="788ab-131">Обратите внимание, что надстройки, имеющие разные поддомены, будут иметь разные экземпляры `AsyncStorage` (например, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span><span class="sxs-lookup"><span data-stu-id="788ab-131">Note that add-ins which have different subdomains will have different instances of `AsyncStorage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span></span> 

<span data-ttu-id="788ab-132">Так как `AsyncStorage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".</span><span class="sxs-lookup"><span data-stu-id="788ab-132">Because `AsyncStorage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="788ab-133">Ниже указаны методы, доступные в объекте `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="788ab-133">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - <span data-ttu-id="788ab-134">`multiRemove`: вы обратите внимание, что реализация метода для очистки всей информации отсутствует (например, `clear`).</span><span class="sxs-lookup"><span data-stu-id="788ab-134">`multiRemove`: You will note that there is no implementation of a method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="788ab-135">Вместо этого вам следует использовать `multiRemove` для одновременного удаления нескольких записей.</span><span class="sxs-lookup"><span data-stu-id="788ab-135">Instead, you should instead use `multiRemove` to remove multiple entries at a time.</span></span>

### <a name="asyncstorage-example"></a><span data-ttu-id="788ab-136">Пример AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="788ab-136">AsyncStorage example</span></span> 

<span data-ttu-id="788ab-137">В следующем примере кода вызывается `AsyncStorage.setItem` функция для установки ключа и значения `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="788ab-137">The following code sample calls the `AsyncStorage.setItem` function to set a key and value into `AsyncStorage`.</span></span>

```JavaScript
function StoreValue(key, value) {

  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="788ab-138">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="788ab-138">Additional considerations</span></span>

<span data-ttu-id="788ab-139">Чтобы создать надстройку, которая будет работать на различных платформах (один из основных клиентов надстроек Office), вам не следует выполнять доступ к модели DOM в пользовательских функциях или использовать библиотеки, такие как jQuery, которые используют модель DOM.</span><span class="sxs-lookup"><span data-stu-id="788ab-139">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="788ab-140">В Excel для Windows, где пользовательские функции используют среду выполнения JavaScript, пользовательские функции не могут выполнять доступ к модели DOM.</span><span class="sxs-lookup"><span data-stu-id="788ab-140">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="788ab-141">См. также</span><span class="sxs-lookup"><span data-stu-id="788ab-141">See also</span></span>

* [<span data-ttu-id="788ab-142">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="788ab-142">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="788ab-143">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="788ab-143">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="788ab-144">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="788ab-144">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="788ab-145">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="788ab-145">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="788ab-146">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="788ab-146">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

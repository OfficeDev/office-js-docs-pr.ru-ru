---
ms.date: 02/06/2019
description: Сведения об основных сценариях при разработке пользовательских функций Excel, которые используют новую среду выполнения JavaScript.
title: Среда выполнения для пользовательских функций Excel (предварительный просмотр)
localization_priority: Normal
ms.openlocfilehash: d891a41dc9e142ef3cfaa00c8b54d8d27913c57d
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982043"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="55698-103">Среда выполнения для пользовательских функций Excel (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="55698-103">Runtime for Excel custom functions (preview)</span></span>

<span data-ttu-id="55698-104">Пользовательские функции используют новую среду выполнения JavaScript, отличающимся от среды выполнения, используемой другими частями надстройки, такими как область задач или другие элементы пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="55698-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="55698-105">Эта среда выполнения JavaScript предназначена для оптимизации производительности вычислений в пользовательских функциях и представляет новые API, которые можно использовать для выполнения стандартных действий в Интернете в пределах пользовательских функций, например для отправления запроса внешних данных или обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="55698-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="55698-106">Среда выполнения JavaScript также обеспечивает доступ к новым API в пространстве имен `OfficeRuntime`, которые могут быть использованы в пределах пользовательских функций или другими частями надстройки для хранения данных или отображения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="55698-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="55698-107">В этой статье объясняется, как использовать такие API в пределах пользовательских функций, а также приводятся другие важные замечания, которые следует учитывать при разработке пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="55698-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="55698-108">Запрос внешних данных</span><span class="sxs-lookup"><span data-stu-id="55698-108">Requesting external data</span></span>

<span data-ttu-id="55698-109">В пределах пользовательской функции можно запрашивать внешние данные с помощью такого API, как [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или с помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="55698-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="55698-110">В среде выполнения JavaScript, используемых настраиваемых функций XHR реализует дополнительные меры безопасности, требуя [Политики единого происхождения](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="55698-110">Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="55698-111">Обратите внимание на то, что простая реализация CORS нельзя использовать файлы cookie и поддерживает только простых методов (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="55698-111">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="55698-112">Простой CORS принимает простой заголовков с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="55698-112">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="55698-113">Вы также можете использовать `Content-Type` предоставляемых верхнего колонтитула в простой CORS, что тип контента является `application/x-www-form-urlencoded`, `text/plain`, или `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="55698-113">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="55698-114">Пример XHR</span><span class="sxs-lookup"><span data-stu-id="55698-114">XHR example</span></span>

<span data-ttu-id="55698-115">В приведенном ниже примере кода функция `getTemperature` вызывает функцию `sendWebRequest` для получения температуры в определенной области на основе идентификатора термометра.</span><span class="sxs-lookup"><span data-stu-id="55698-115">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="55698-116">Функция `sendWebRequest` использует XHR для отправления запроса `GET` в конечную точку, которая может предоставить данные.</span><span class="sxs-lookup"><span data-stu-id="55698-116">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>

> [!NOTE] 
> <span data-ttu-id="55698-117">При использовании Fetch или XHR возвращается новое значение `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="55698-117">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="55698-118">До сентября 2018 г. необходимо было указывать `OfficeExtension.Promise` использовать обещания в пределах API Office JavaScript, но теперь вы можете просто использовать JavaScript `Promise`.</span><span class="sxs-lookup"><span data-stu-id="55698-118">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="55698-119">Получение данных через WebSockets</span><span class="sxs-lookup"><span data-stu-id="55698-119">Receiving data via WebSockets</span></span>

<span data-ttu-id="55698-120">В пределах пользовательской функции можно использовать [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) для обмена данными через постоянное соединение с сервером.</span><span class="sxs-lookup"><span data-stu-id="55698-120">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="55698-121">С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.</span><span class="sxs-lookup"><span data-stu-id="55698-121">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="55698-122">Пример WebSockets</span><span class="sxs-lookup"><span data-stu-id="55698-122">WebSockets example</span></span>

<span data-ttu-id="55698-123">Приведенный ниже примера кода устанавливает соединение `WebSocket`, а затем заносит в журнал каждое входящее сообщение от сервера.</span><span class="sxs-lookup"><span data-stu-id="55698-123">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="55698-124">Хранения данных и доступ к ним</span><span class="sxs-lookup"><span data-stu-id="55698-124">Storing and accessing data</span></span>

<span data-ttu-id="55698-125">В пределах функции (или в пределах любой другой части надстройки) можно хранить данные и выполнять доступ к ним с помощью объекта `OfficeRuntime.AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="55698-125">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="55698-126">`AsyncStorage` — это постоянная незашифрованная система-хранилище пары "ключ-значение", обеспечивающая альтернативу хранилищу [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), которое нельзя использовать в пределах пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="55698-126">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="55698-127">Надстройка может хранить до 10 МБ данных, используя `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="55698-127">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="55698-128">`AsyncStorage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным.</span><span class="sxs-lookup"><span data-stu-id="55698-128">`AsyncStorage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="55698-129">Например, токены для аутентификации пользователей могут храниться в `AsyncStorage`, так как доступ к нему могут выполнять и пользовательская функция, и элементы пользовательского интерфейса надстройки, такие как область задач.</span><span class="sxs-lookup"><span data-stu-id="55698-129">For example, tokens for user authentication may be stored in `AsyncStorage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="55698-130">Точно так же, если две надстройки используют один и тот же домен (например, www.contoso.com/addin1, www.contoso.com/addin2), им также разрешается обмен информацией в оба направления через `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="55698-130">Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `AsyncStorage`.</span></span> <span data-ttu-id="55698-131">Обратите внимание, что надстройки, имеющие разные поддомены, будут иметь разные экземпляры `AsyncStorage` (например, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span><span class="sxs-lookup"><span data-stu-id="55698-131">Note that add-ins which have different subdomains will have different instances of `AsyncStorage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span></span> 

<span data-ttu-id="55698-132">Так как `AsyncStorage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".</span><span class="sxs-lookup"><span data-stu-id="55698-132">Because `AsyncStorage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="55698-133">Ниже указаны методы, доступные в объекте `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="55698-133">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - <span data-ttu-id="55698-134">`multiRemove`: вы обратите внимание, что реализация метода для очистки всей информации отсутствует (например, `clear`).</span><span class="sxs-lookup"><span data-stu-id="55698-134">`multiRemove`: You will note that there is no implementation of a method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="55698-135">Вместо этого вам следует использовать `multiRemove` для одновременного удаления нескольких записей.</span><span class="sxs-lookup"><span data-stu-id="55698-135">Instead, you should instead use `multiRemove` to remove multiple entries at a time.</span></span>

### <a name="asyncstorage-example"></a><span data-ttu-id="55698-136">Пример AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="55698-136">AsyncStorage example</span></span> 

<span data-ttu-id="55698-137">Указанный ниже пример кода вызывает функцию `AsyncStorage.getItem` для получения значения из хранилища.</span><span class="sxs-lookup"><span data-stu-id="55698-137">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

```typescript
_goGetData = async () => {
    try {
        const value = await AsyncStorage.getItem('toDoItem');
        if (value !== null) {
            //data exists and you can do something with it here
        }
    } catch (error) {
        //handle errors here
    }
}
```

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="55698-138">Отображение диалогового окна</span><span class="sxs-lookup"><span data-stu-id="55698-138">Displaying a dialog box</span></span>

<span data-ttu-id="55698-139">В пределах пользовательской функции (или в пределах любой другой части надстройки) можно использовать API `OfficeRuntime.displayWebDialog`, чтобы отобразить диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="55698-139">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialog` API to display a dialog box.</span></span> <span data-ttu-id="55698-140">Этот диалоговый API является альтернативой [Dialog API](../develop/dialog-api-in-office-add-ins.md), который можно использовать в пределах областей задач и команд надстройки, но не в пределах пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="55698-140">This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="55698-141">Пример Dialog API</span><span class="sxs-lookup"><span data-stu-id="55698-141">Dialog API example</span></span>

<span data-ttu-id="55698-142">В приведенном ниже примере кода функция `getTokenViaDialog` использует функцию `displayWebDialog` Dialog API для отображения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="55698-142">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialog` function to display a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
        let timeout = 5;
        let count = 0;
        var intervalId = setInterval(function () {
          count++;
          if(_cachedToken) {
            resolve(_cachedToken);
            clearInterval(intervalId);
          }
          if(count >= timeout) {
            reject("Timeout while waiting for token");
            clearInterval(intervalId);
          }
        }, 1000);
      } else {
        _dialogOpen = true;
        OfficeRuntime.displayWebDialog(url, {
          height: '50%',
          width: '50%',
          onMessage: function (message, dialog) {
            _cachedToken = message;
            resolve(message);
            dialog.close();
            return;
          },
          onRuntimeError: function(error, dialog) {
            reject(error);
          },
        }).catch(function (e) {
          reject(e);
        });
      }
    });
  }
}
```

## <a name="additional-considerations"></a><span data-ttu-id="55698-143">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="55698-143">Additional considerations</span></span>

<span data-ttu-id="55698-144">Чтобы создать надстройку, которая будет работать на различных платформах (один из основных клиентов надстроек Office), вам не следует выполнять доступ к модели DOM в пользовательских функциях или использовать библиотеки, такие как jQuery, которые используют модель DOM.</span><span class="sxs-lookup"><span data-stu-id="55698-144">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="55698-145">В Excel для Windows, где пользовательские функции используют среду выполнения JavaScript, пользовательские функции не могут выполнять доступ к модели DOM.</span><span class="sxs-lookup"><span data-stu-id="55698-145">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="55698-146">См. также</span><span class="sxs-lookup"><span data-stu-id="55698-146">See also</span></span>

* [<span data-ttu-id="55698-147">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="55698-147">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="55698-148">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="55698-148">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="55698-149">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="55698-149">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="55698-150">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="55698-150">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="55698-151">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="55698-151">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

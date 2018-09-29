---
ms.date: 09/27/2018
description: Настраиваемые функции Excel используют новую среду выполнения JavaScript, которая отличается от стандартного управления во время выполнения веб-представления надстройки.
title: Среда выполнения для настраиваемых функций Excel
ms.openlocfilehash: ce9678d68860c0f8f4c868712155af7824ceb93f
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348109"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="6dfb5-103">Среда выполнения для настраиваемых функций Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="6dfb5-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="6dfb5-104">Настраиваемые функции расширяют возможности Excel за счет применения новой среды выполнения JavaScript, использующей изолированную подсистему JavaScript, а не веб-браузер.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-104">Custom functions extend Excel’s capabilities by using a new JavaScript runtime that uses a sandboxed JavaScript engine rather than a web browser.</span></span> <span data-ttu-id="6dfb5-105">Поскольку настраиваемые функции не требуют отображения элементов пользовательского интерфейса, новая среда выполнения JavaScript оптимизирована для выполнения вычислений, что позволяет одновременно выполнять тысячи настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-105">Because custom functions do not need to render UI elements, the new JavaScript runtime is optimized for performing calculations, enabling you to run thousands of custom functions simultaneously.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="key-facts-about-the-new-javascript-runtime"></a><span data-ttu-id="6dfb5-106">Основные факты, связанные с новой средой выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="6dfb5-106">Key facts about the new JavaScript runtime</span></span> 

<span data-ttu-id="6dfb5-107">Только  настраиваемые  функции в рамках надстройки будут использовать новую среду выполнения JavaScript, описанную в данной статье.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-107">Only custom functions within an add-in will use the new JavaScript runtime that's described in this article.</span></span> <span data-ttu-id="6dfb5-108">Если надстройка включает в себя другие компоненты, такие как области задач и другие элементы пользовательского интерфейса, помимо пользовательских функций эти компоненты надстройки будут выполняться в среде выполнения веб-представления браузера.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-108">If an add-in includes other components such as task panes and other UI elements, in addition to custom functions, these other components of the add-in will continue to run in the browser-like WebView runtime.</span></span>  <span data-ttu-id="6dfb5-109">Кроме того:</span><span class="sxs-lookup"><span data-stu-id="6dfb5-109">Additionally:</span></span> 

- <span data-ttu-id="6dfb5-110">Среда выполнения JavaScript не предоставляет доступа к модели объектов документа (DOM), либо поддержке библиотек такие как jQuery, зависящие от DOM.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-110">The JavaScript runtime does not provide access to the Document Object Model (DOM) or support libraries like jQuery that rely on the DOM.</span></span>

- <span data-ttu-id="6dfb5-111">Настраиваемая функция, которая определена в файле JavaScript надстройки, может вернуть регулярный JavaScript `Promise` вместо `OfficeExtension.Promise`.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-111">A custom function that's defined in an add-in's JavaScript file can return a regular JavaScript `Promise` instead of returning `OfficeExtension.Promise`.</span></span>  

- <span data-ttu-id="6dfb5-112">Файл JSON, который определяет настраиваемую функцию метаданных не требуется указывать **sync** или **async** в **в рамках параметров**.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-112">The JSON file that specifies custom function metatdata does not need to specify **sync** or **async** within **options**.</span></span>

## <a name="new-apis"></a><span data-ttu-id="6dfb5-113">Новые API-интерфейсы</span><span class="sxs-lookup"><span data-stu-id="6dfb5-113">New Excel JavaScript APIs</span></span> 

<span data-ttu-id="6dfb5-114">Среда выполнения JavaScript, которую используют настраиваемые функции имеет следующие API-интерфейсы:</span><span class="sxs-lookup"><span data-stu-id="6dfb5-114">The JavaScript runtime that's used by custom functions has the following APIs:</span></span>

- [<span data-ttu-id="6dfb5-115">XHR</span><span class="sxs-lookup"><span data-stu-id="6dfb5-115">XHR</span></span>](#xhr)
- [<span data-ttu-id="6dfb5-116">WebSocket</span><span class="sxs-lookup"><span data-stu-id="6dfb5-116">WebSockets</span></span>](#websockets)
- [<span data-ttu-id="6dfb5-117">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="6dfb5-117">AsyncStorage</span></span>](#asyncstorage)
- [<span data-ttu-id="6dfb5-118">API общих диалогов</span><span class="sxs-lookup"><span data-stu-id="6dfb5-118">Dialog API requirement sets</span></span>](#dialog-api)

### <a name="xhr"></a><span data-ttu-id="6dfb5-119">XHR</span><span class="sxs-lookup"><span data-stu-id="6dfb5-119">XHR</span></span>

<span data-ttu-id="6dfb5-120">XHR означает [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), стандартный веб-API, который выдает HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-120">XHR stands for [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="6dfb5-121">В новой среде выполнения JavaScript XHR реализует дополнительные меры безопасности, требуя [Политику единого происхождения](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="6dfb5-121">In the new JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

<span data-ttu-id="6dfb5-122">В следующем примере кода `getTemperature()` функция отправляет веб-запрос для получения температуры отдельной области на основе идентификатора термометра.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-122">In the following code sample, the `getTemperature()` function sends a web request to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="6dfb5-123">`sendWebRequest()` Функция использует XHR для выдачи `GET` запроса к конечной точке, которая может предоставлять данные.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-123">The `sendWebRequest()` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>  

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ //sendWebRequest is defined later in this code sample
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

//Helper method that uses Office's implementation of XMLHttpRequest in the new JavaScript runtime for custom functions  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
          };
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

```

### <a name="websockets"></a><span data-ttu-id="6dfb5-124">WebSockets</span><span class="sxs-lookup"><span data-stu-id="6dfb5-124">WebSockets</span></span>

<span data-ttu-id="6dfb5-125">[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) — это сетевой протокол, который создает в режиме реального времени обмен данными между сервером и одним или несколькими клиентами.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-125">[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) is a networking protocol that creates real-time communication between a server and one or more clients.</span></span> <span data-ttu-id="6dfb5-126">Его часто используется для приложений чата, так как он позволяет одновременно читать и записывать текст.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-126">It is often used for chat applications because it allows you to read and write text simultaneously.</span></span>  

<span data-ttu-id="6dfb5-127">Как показано в следующем примере  кода, настраиваемые функции могут использовать WebSockets.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-127">As shown in the following code sample, custom functions can use WebSockets.</span></span> <span data-ttu-id="6dfb5-128">В этом примере WebSocket регистрирует каждое сообщение, которое он получает.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-128">In this example, the WebSocket logs each message that it receives.</span></span>

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a><span data-ttu-id="6dfb5-129">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="6dfb5-129">AsyncStorage</span></span>

<span data-ttu-id="6dfb5-130">AsyncStorage — это система хранения "ключ-значение", которая может использоваться для хранения маркеров проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-130">AsyncStorage is a key-value storage system that can be used to store authentication tokens.</span></span> <span data-ttu-id="6dfb5-131">Она:</span><span class="sxs-lookup"><span data-stu-id="6dfb5-131">It is framework-agnostic.</span></span>

- <span data-ttu-id="6dfb5-132">Постоянна</span><span class="sxs-lookup"><span data-stu-id="6dfb5-132">persistent</span></span>
- <span data-ttu-id="6dfb5-133">Без шифрования</span><span class="sxs-lookup"><span data-stu-id="6dfb5-133">Unencrypted</span></span>
- <span data-ttu-id="6dfb5-134">Асинхронна</span><span class="sxs-lookup"><span data-stu-id="6dfb5-134">Asynchronous calls</span></span>

<span data-ttu-id="6dfb5-135">AsyncStorage глобально доступна для всех компонентов надстройки.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-135">AsyncStorage is globally available to all parts of your add-in.</span></span> <span data-ttu-id="6dfb5-136">Для настраиваемых функций `AsyncStorage` представленна как глобальный объект.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-136">For custom functions, `AsyncStorage` is exposed as a global object.</span></span> <span data-ttu-id="6dfb5-137">(Для других компонентов надстройки, такие как области задач и другие элементы, использующие среду выполнения веб-представления AsyncStorage предоставляется через `OfficeRuntime`.) Каждая надстройка имеет свой собственный раздел хранилища, по умолчанию размером 5 МБ.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-137">(For other parts of your add-in, such as task panes and other elements that use the WebView runtime, AsyncStorage is exposed through `OfficeRuntime`.) Each add-in has its own storage partition, with a default size of 5MB.</span></span> 

<span data-ttu-id="6dfb5-138">Доступны следующие методы на `AsyncStorage` объекте:</span><span class="sxs-lookup"><span data-stu-id="6dfb5-138">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`
 
<span data-ttu-id="6dfb5-139">В настоящее время `mergeItem` and `multiMerge` методы не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-139">At this time, the `mergeItem` and `multiMerge` methods are not supported.</span></span>

<span data-ttu-id="6dfb5-140">Следующий пример кода вызывает `AsyncStorage.getItem` функции для извлечения значения из хранилища.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-140">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

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
}
```

### <a name="dialog-api"></a><span data-ttu-id="6dfb5-141">API общих диалогов</span><span class="sxs-lookup"><span data-stu-id="6dfb5-141">Dialog API scenarios</span></span>

<span data-ttu-id="6dfb5-142">API общих диалогов позволяет открыть диалоговое окно с запросом входа пользователя.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-142">The Dialog API enables you to open a dialog box that prompts user sign-in.</span></span> <span data-ttu-id="6dfb5-143">API общих диалогов можно использовать для проверки подлинности пользователей с помощью внешних ресурсов, например Google или Facebook, прежде чем пользователь сможет использовать функцию.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-143">You can use the Dialog API to require user authentication through an outside resource, such as Google or Facebook, before the user can use your function.</span></span>   

<span data-ttu-id="6dfb5-144">В следующем примере кода метод `getTokenViaDialog()` использует метод API общих диалогов `displayWebDialog()`, чтобы открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-144">In the following code sample, the `getTokenViaDialog()` method uses the Dialog API’s `displayWebDialog()` method to open a dialog box.</span></span>

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
            dialog.closeDialog();
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

> [!NOTE]
> <span data-ttu-id="6dfb5-145">API общих диалогов, описанный в данном разделе, являются частью новой среды выполнения JavaScript для настраиваемых функций и может использоваться только в пределах настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-145">The Dialog API described in this section is part of the new JavaScript runtime for custom functions and can be used only within custom functions.</span></span> <span data-ttu-id="6dfb5-146">Этот API-интерфейс  отличается от [API общих диалогов](../develop/dialog-api-in-office-add-ins.md), который может использоваться в области задач и командах надстройки.</span><span class="sxs-lookup"><span data-stu-id="6dfb5-146">This API is different from the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands.</span></span>

## <a name="see-also"></a><span data-ttu-id="6dfb5-147">См. также</span><span class="sxs-lookup"><span data-stu-id="6dfb5-147">See also</span></span>

* [<span data-ttu-id="6dfb5-148">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="6dfb5-148">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="6dfb5-149">Настраиваемые функции метаданных</span><span class="sxs-lookup"><span data-stu-id="6dfb5-149">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="6dfb5-150">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="6dfb5-150">Custom functions best practices</span></span>](custom-functions-best-practices.md)

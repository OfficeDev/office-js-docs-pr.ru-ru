---
ms.date: 10/03/2018
description: Основные сведения о ключевых сценариях разработки настраиваемых функций Excel, использующие новую среду выполнения JavaScript.
title: Среда выполнения для настраиваемых функций Excel
ms.openlocfilehash: a48b02a8ca404b51740d9052d199da934eb9312e
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459107"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="f0d7d-103">Среда выполнения для настраиваемых функций Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f0d7d-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="f0d7d-104">Настраиваемые функции используют новую среду выполнения JavaScript, которая отличается от среды выполнения, используемой другими частями надстройки, такими как область задач или другие элементы пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="f0d7d-105">Эта среда выполнения JavaScript предназначена для оптимизации производительности вычислений в настраиваемых функциях и отображает новые API, которые вы можете использовать для выполнения стандартных действий в Интернете, например, запрашивать внешние данные или обмениваться данными при постоянном подключении к серверу.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="f0d7d-106">Среда выполнения JavaScript также предоставляет доступ к новым API в пространстве имен `OfficeRuntime`, которое может использоваться в настраиваемых функциях или другими частями надстройки для хранения данных или отображения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="f0d7d-107">В этой статье описывается, как использовать эти API в настраиваемых функциях, а также излагаются дополнительные соображения, которые следует учитывать при разработке настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="f0d7d-108">Запрос внешних данных</span><span class="sxs-lookup"><span data-stu-id="f0d7d-108">Requesting external data</span></span>

<span data-ttu-id="f0d7d-109">В настраиваемой функции можно запросить внешние данные с помощью API, например, [Fetch API](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или используя [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)— стандартный API, который выдает HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="f0d7d-110">В новой среде выполнения JavaScript XHR реализует дополнительные меры безопасности, требуя [исходную политику](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой механизм [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="f0d7d-110">In the new JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="f0d7d-111">Пример XHR</span><span class="sxs-lookup"><span data-stu-id="f0d7d-111">XHR example</span></span>

<span data-ttu-id="f0d7d-112">В следующем примере кода функция `getTemperature` вызывает функцию `sendWebRequest` для получения температуры отдельной области на основе идентификатора термометра.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-112">In the following code sample, the  function sends a web request to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="f0d7d-113">Функция `sendWebRequest` использует XHR для выдачи запроса `GET` конечной точке, которая может предоставить данные.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-113">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span> 

> [!NOTE] 
> <span data-ttu-id="f0d7d-114">При использовании Fetch или XHR возвращается новый JavaScript `Promise` .</span><span class="sxs-lookup"><span data-stu-id="f0d7d-114">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="f0d7d-115">До сентября 2018 года необходимо было указывать `OfficeExtension.Promise`, чтобы использовать обещания в API Office JavaScript, но теперь вы можете просто использовать JavaScript `Promise`.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-115">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="f0d7d-116">Получение данных с помощью WebSockets</span><span class="sxs-lookup"><span data-stu-id="f0d7d-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="f0d7d-117">В пользовательской функции можно использовать [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) для обмена данными при постоянном подключении к серверу.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-117">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="f0d7d-118">При использовании WebSockets ваша настраиваемая функция может открывать подключение с сервером, а затем автоматически получать сообщения от сервера, когда происходят определенные события, без необходимости явно запрашивать данные у сервера.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-118">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="f0d7d-119">Пример WebSockets</span><span class="sxs-lookup"><span data-stu-id="f0d7d-119">WebSockets example</span></span>

<span data-ttu-id="f0d7d-120">В следующем примере устанавливается подключение `WebSocket`, а затем регистрируются все входящие сообщения с сервера.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="f0d7d-121">Хранение данных и доступ к ним</span><span class="sxs-lookup"><span data-stu-id="f0d7d-121">Storing and accessing data</span></span>

<span data-ttu-id="f0d7d-122">Вы можете хранить данные и получать к ним доступ в настраиваемой функции (или в любой другой части надстройки), используя объект `OfficeRuntime.AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-122">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="f0d7d-123">`AsyncStorage` — это постоянная, не зашифрованная система хранения с ключевым значением, являющаяся альтернативой для хранилища [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), которое не может использоваться для настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-123">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="f0d7d-124">Надстройка может хранить до 10 МБ данных с помощью `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-124">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="f0d7d-125">Доступны следующие методы на объекте `AsyncStorage`:</span><span class="sxs-lookup"><span data-stu-id="f0d7d-125">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a><span data-ttu-id="f0d7d-126">Пример AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="f0d7d-126">AsyncStorage example</span></span> 

<span data-ttu-id="f0d7d-127">Следующий пример кода вызывает функцию `AsyncStorage.getItem` для извлечения значения из хранилища.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-127">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

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

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="f0d7d-128">Отображение диалогового окна</span><span class="sxs-lookup"><span data-stu-id="f0d7d-128">Open a dialog box</span></span>

<span data-ttu-id="f0d7d-129">В настраиваемой функции (или в любой другой части надстройки) можно использовать API`OfficeRuntime.displayWebDialogOptions` для отображения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-129">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box.</span></span> <span data-ttu-id="f0d7d-130">Этот API диалогового окна является альтернативой [API диалогового окна](../develop/dialog-api-in-office-add-ins.md) , который можно использовать в области задач и команд надстроек, но не в настраиваемых функциях.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-130">This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="f0d7d-131">Пример API диалогового окна</span><span class="sxs-lookup"><span data-stu-id="f0d7d-131">Dialog API example</span></span> 

<span data-ttu-id="f0d7d-132">В следующем примере кода функция `getTokenViaDialog` использует функцию API диалогового окна `displayWebDialogOptions`, чтобы открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-132">In the following code sample, the `getTokenViaDialog` method uses the Dialog API’s `displayWebDialogOptions` method to open a dialog box.</span></span>

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
        OfficeRuntime.displayWebDialogOptions(url, {
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

## <a name="additional-considerations"></a><span data-ttu-id="f0d7d-133">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="f0d7d-133">Additional considerations</span></span>

<span data-ttu-id="f0d7d-134">Чтобы создать надстройку, которая будет работать на нескольких платформах (для одного из основных клиентов надстроек Office), вы не должны запрашивать доступ к модели DOM в настраиваемых функциях или использовать библиотеки, такие как jQuery, которые полагаются на DOM.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-134">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="f0d7d-135">В Excel для Windows, где настраиваемые функции используют среду выполнения JavaScript, у настраиваемых функций нет доступа к DOM.</span><span class="sxs-lookup"><span data-stu-id="f0d7d-135">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="f0d7d-136">См. также</span><span class="sxs-lookup"><span data-stu-id="f0d7d-136">See also</span></span>

* [<span data-ttu-id="f0d7d-137">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="f0d7d-137">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f0d7d-138">Настраиваемые функции метаданных</span><span class="sxs-lookup"><span data-stu-id="f0d7d-138">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f0d7d-139">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="f0d7d-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="f0d7d-140">Руководство по настраиваемым функциям Excel</span><span class="sxs-lookup"><span data-stu-id="f0d7d-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

---
ms.date: 09/27/2018
description: Настраиваемые функции Excel используют новую среду выполнения JavaScript, которая отличается от стандартного управления во время выполнения веб-представления надстройки.
title: Среда выполнения для настраиваемых функций Excel
ms.openlocfilehash: 7489cd66851d1e0c24ef573ffa920b794cf749c2
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348761"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Среда выполнения для настраиваемых функций Excel (предварительная версия)

Настраиваемые функции расширяют возможности Excel за счет применения новой среды выполнения JavaScript, использующей изолированную подсистему JavaScript, а не веб-браузер. Поскольку настраиваемые функции не требуют отображения элементов пользовательского интерфейса, новая среда выполнения JavaScript оптимизирована для выполнения вычислений, что позволяет одновременно выполнять тысячи настраиваемых функций.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="key-facts-about-the-new-javascript-runtime"></a>Основные факты, связанные с новой средой выполнения JavaScript 

Только  настраиваемые  функции в рамках надстройки будут использовать новую среду выполнения JavaScript, описанную в данной статье. Если надстройка включает в себя другие компоненты, такие как области задач и другие элементы пользовательского интерфейса, помимо пользовательских функций эти компоненты надстройки будут выполняться в среде выполнения веб-представления браузера.  Кроме того: 

- Среда выполнения JavaScript не предоставляет доступа к модели объектов документа (DOM), либо поддержке библиотек такие как jQuery, зависящие от DOM.

- Настраиваемая функция, которая определена в файле JavaScript надстройки, может вернуть регулярный JavaScript `Promise` вместо `OfficeExtension.Promise`.  

- Файл JSON, который определяет настраиваемую функцию метаданных не требуется указывать **sync** или **async** в **в рамках параметров**.

## <a name="new-apis"></a>Новые API-интерфейсы 

Среда выполнения JavaScript, которую используют настраиваемые функции имеет следующие API-интерфейсы:

- [XHR](#xhr)
- [WebSocket](#websockets)
- [AsyncStorage](#asyncstorage)
- [API общих диалогов](#dialog-api)

### <a name="xhr"></a>XHR

XHR означает [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), стандартный веб-API, который выдает HTTP-запросы для взаимодействия с серверами. В новой среде выполнения JavaScript XHR реализует дополнительные меры безопасности, требуя [Политику единого происхождения](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой [CORS](https://www.w3.org/TR/cors/).  

В следующем примере кода `getTemperature()` функция отправляет веб-запрос для получения температуры отдельной области на основе идентификатора термометра. `sendWebRequest()` Функция использует XHR для выдачи `GET` запроса к конечной точке, которая может предоставлять данные.  

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

### <a name="websockets"></a>WebSockets

[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) — это сетевой протокол, который создает в режиме реального времени обмен данными между сервером и одним или несколькими клиентами. Его часто используется для приложений чата, так как он позволяет одновременно читать и записывать текст.  

Как показано в следующем примере  кода, настраиваемые функции могут использовать WebSockets. В этом примере WebSocket регистрирует каждое сообщение, которое он получает.

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a>AsyncStorage

AsyncStorage — это система хранения "ключ-значение", которая может использоваться для хранения маркеров проверки подлинности. Она:

- Постоянна
- Без шифрования
- Асинхронна

AsyncStorage глобально доступна для всех компонентов надстройки. Для настраиваемых функций `AsyncStorage` представленна как глобальный объект. (Для других компонентов надстройки, такие как области задач и другие элементы, использующие среду выполнения веб-представления AsyncStorage предоставляется через `OfficeRuntime`.) Каждая надстройка имеет свой собственный раздел хранилища, по умолчанию размером 5 МБ. 

Доступны следующие методы на `AsyncStorage` объекте:
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`
 
В настоящее время `mergeItem` and `multiMerge` методы не поддерживаются.

Следующий пример кода вызывает `AsyncStorage.getItem` функции для извлечения значения из хранилища.

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

### <a name="dialog-api"></a>API общих диалогов

API общих диалогов позволяет открыть диалоговое окно с запросом входа пользователя. API общих диалогов можно использовать для проверки подлинности пользователей с помощью внешних ресурсов, например Google или Facebook, прежде чем пользователь сможет использовать функцию.   

В следующем примере кода метод `getTokenViaDialog()` использует метод API общих диалогов `displayWebDialog()`, чтобы открыть диалоговое окно.

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
> API общих диалогов, описанный в данном разделе, являются частью новой среды выполнения JavaScript для настраиваемых функций и может использоваться только в пределах настраиваемых функций. Этот API-интерфейс  отличается от [API общих диалогов](../develop/dialog-api-in-office-add-ins.md), который может использоваться в области задач и командах надстройки.

## <a name="see-also"></a>См. также

* [Создание настраиваемых функций в Excel](custom-functions-overview.md)
* [Настраиваемые функции метаданных](custom-functions-json.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)
* [Руководство по настраиваемым функциям Excel](excel-tutorial-custom-functions.md)

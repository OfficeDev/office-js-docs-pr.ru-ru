---
ms.date: 10/17/2018
description: Основные сведения о ключевых сценариях разработки настраиваемых функций Excel, использующие новую среду выполнения JavaScript.
title: Среда выполнения для настраиваемых функций Excel
ms.openlocfilehash: 333816c3916af1490d14b8344c4bb49094f9a7f9
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640017"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Среда выполнения для настраиваемых функций Excel (предварительная версия)

Настраиваемые функции используют новую среду выполнения JavaScript, которая отличается от среды выполнения, используемой другими частями надстройки, такими как область задач или другие элементы пользовательского интерфейса. Эта среда выполнения JavaScript предназначена для оптимизации производительности вычислений в настраиваемых функциях и отображает новые API, которые вы можете использовать для выполнения стандартных действий в Интернете, например, запрашивать внешние данные или обмениваться данными при постоянном подключении к серверу. Среда выполнения JavaScript также предоставляет доступ к новым API в пространстве имен `OfficeRuntime`, которое может использоваться в настраиваемых функциях или другими частями надстройки для хранения данных или отображения диалогового окна. В этой статье описывается, как использовать эти API в настраиваемых функциях, а также излагаются дополнительные соображения, которые следует учитывать при разработке настраиваемых функций.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>Запрос внешних данных

В настраиваемой функции можно запросить внешние данные с помощью API, например, [Fetch API](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или используя [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)— стандартный API, который выдает HTTP-запросы для взаимодействия с серверами. В новой среде выполнения JavaScript XHR реализует дополнительные меры безопасности, требуя [ исходную политику](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой механизм [ CORS](https://www.w3.org/TR/cors/).  

### <a name="xhr-example"></a>Пример XHR

В следующем примере кода функция `getTemperature` вызывает функцию `sendWebRequest` для получения температуры отдельной области на основе идентификатора термометра. Функция `sendWebRequest` использует XHR для выдачи запроса `GET` к конечной точке, которая может предоставлять данные. 

> [!NOTE] 
> При использовании Fetch или XHR возвращается новый JavaScript `Promise`. До сентября 2018 года необходимо было указывать `OfficeExtension.Promise`, чтобы использовать обещания в API Office JavaScript, но теперь вы можете просто использовать JavaScript `Promise`.

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

## <a name="receiving-data-via-websockets"></a>Получение данных с помощью WebSockets

В пользовательской функции можно использовать [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) для обмена данными при постоянном подключении к серверу. При использовании WebSockets ваша настраиваемая функция может открывать подключение с сервером, а затем автоматически получать сообщения от сервера, когда происходят определенные события, без необходимости прямо запрашивать данные у сервера.

### <a name="websockets-example"></a>Пример WebSockets

В следующем примере устанавливается подключение `WebSocket`, а затем регистрируются все входящие сообщения с сервера. 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Хранение данных и доступ к ним

В настраиваемой функции (или в любой другой части надстройки) можно использовать объект `OfficeRuntime.AsyncStorage` для хранения данных и доступа к ним. `AsyncStorage` — это постоянная, не зашифрованная система хранения с ключевым значением, являющаяся альтернативой для хранилища [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), которое не может использоваться для настраиваемых функций. Надстройка может хранить до 10 МБ данных с помощью `AsyncStorage`.

Доступны следующие методы в объекте `AsyncStorage`:
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a>Пример AsyncStorage 

Следующий пример кода вызывает функцию `AsyncStorage.getItem` для извлечения значения из хранилища.

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

## <a name="displaying-a-dialog-box"></a>Отображение диалогового окна

В настраиваемой функции (или в любой другой части надстройки) можно использовать API `OfficeRuntime.displayWebDialogOptions` для отображения диалогового окна. Этот API диалогового окна является альтернативой [API общих диалогов](../develop/dialog-api-in-office-add-ins.md) , который можно использовать в области задач и команд надстроек, но не в настраиваемых функциях.

### <a name="dialog-api-example"></a>Пример API общих диалогов 

В следующем примере кода функция `getTokenViaDialog` использует функцию API общих диалогов `displayWebDialogOptions`, чтобы открыть диалоговое окно.

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

## <a name="additional-considerations"></a>Дополнительные рекомендации

Чтобы создать надстройку, которая будет работать на нескольких платформах (для одного из основных клиентов надстроек Office), вы не должны запрашивать доступ к модели DOM в настраиваемых функциях или использовать библиотеки, такие как jQuery, которые полагаются на DOM. В Excel для Windows настраиваемые функции, использующие среду выполнения JavaScripte, не могут получить доступ к DOM.

## <a name="see-also"></a>См. также

* [Создание настраиваемых функций в Excel](custom-functions-overview.md)
* [Настраиваемые функции метаданных](custom-functions-json.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)
* [Руководство по настраиваемым функциям Excel](excel-tutorial-custom-functions.md)

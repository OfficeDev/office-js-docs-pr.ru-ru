---
ms.date: 01/08/2019
description: Сведения об основных сценариях при разработке пользовательских функций Excel, которые используют новую среду выполнения JavaScript.
title: Среда выполнения для пользовательских функций Excel (предварительный просмотр)
ms.openlocfilehash: 2610be95ea255d14c577d8b9215f32a79ab04463
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770583"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Среда выполнения для пользовательских функций Excel (предварительный просмотр)

Пользовательские функции используют новую среду выполнения JavaScript, отличающимся от среды выполнения, используемой другими частями надстройки, такими как область задач или другие элементы пользовательского интерфейса. Эта среда выполнения JavaScript предназначена для оптимизации производительности вычислений в пользовательских функциях и представляет новые API, которые можно использовать для выполнения стандартных действий в Интернете в пределах пользовательских функций, например для отправления запроса внешних данных или обмена данными через постоянное соединение с сервером. Среда выполнения JavaScript также обеспечивает доступ к новым API в пространстве имен `OfficeRuntime`, которые могут быть использованы в пределах пользовательских функций или другими частями надстройки для хранения данных или отображения диалогового окна. В этой статье объясняется, как использовать такие API в пределах пользовательских функций, а также приводятся другие важные замечания, которые следует учитывать при разработке пользовательских функций.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>Запрос внешних данных

В пределах пользовательской функции можно запрашивать внешние данные с помощью такого API, как [Fetch](https://developer.mozilla.org/ru-RU/docs/Web/API/Fetch_API), или с помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/ru-RU/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами. В среде выполнения JavaScript XHR реализует дополнительные меры по обеспечению безопасности, предъявляя в качестве требования [Политику единого домена](https://developer.mozilla.org/ru-RU/docs/Web/Security/Same-origin_policy) и простой механизм [CORS](https://www.w3.org/TR/cors/).  

### <a name="xhr-example"></a>Пример XHR

В приведенном ниже примере кода функция `getTemperature` вызывает функцию `sendWebRequest` для получения температуры в определенной области на основе идентификатора термометра. Функция `sendWebRequest` использует XHR для отправления запроса `GET` в конечную точку, которая может предоставить данные.

> [!NOTE] 
> При использовании Fetch или XHR возвращается новое значение `Promise` JavaScript. До сентября 2018 г. необходимо было указывать `OfficeExtension.Promise` использовать обещания в пределах API Office JavaScript, но теперь вы можете просто использовать JavaScript `Promise`.

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

## <a name="receiving-data-via-websockets"></a>Получение данных через WebSockets

В пределах пользовательской функции можно использовать [WebSockets](https://developer.mozilla.org/ru-RU/docs/Web/API/WebSockets_API) для обмена данными через постоянное соединение с сервером. С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.

### <a name="websockets-example"></a>Пример WebSockets

Приведенный ниже примера кода устанавливает соединение `WebSocket`, а затем заносит в журнал каждое входящее сообщение от сервера. 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Хранения данных и доступ к ним

В пределах функции (или в пределах любой другой части надстройки) можно хранить данные и выполнять доступ к ним с помощью объекта `OfficeRuntime.AsyncStorage`. `AsyncStorage` — это постоянная незашифрованная система-хранилище пары "ключ-значение", обеспечивающая альтернативу хранилищу [localStorage](https://developer.mozilla.org/ru-RU/docs/Web/API/Window/localStorage), которое нельзя использовать в пределах пользовательских функций. Надстройка может хранить до 10 МБ данных, используя `AsyncStorage`.

`AsyncStorage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным. Например, токены для аутентификации пользователей могут храниться в `AsyncStorage`, так как доступ к нему могут выполнять и пользовательская функция, и элементы пользовательского интерфейса надстройки, такие как область задач. Точно так же, если две надстройки используют один и тот же домен (например, www.contoso.com/addin1, www.contoso.com/addin2), им также разрешается обмен информацией в оба направления через `AsyncStorage`. Обратите внимание, что надстройки, имеющие разные поддомены, будут иметь разные экземпляры `AsyncStorage` (например, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2). 

Так как `AsyncStorage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".

Ниже указаны методы, доступные в объекте `AsyncStorage`.
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`: вы обратите внимание, что реализация метода для очистки всей информации отсутствует (например, `clear`). Вместо этого вам следует использовать `multiRemove` для одновременного удаления нескольких записей.

### <a name="asyncstorage-example"></a>Пример AsyncStorage 

Указанный ниже пример кода вызывает функцию `AsyncStorage.getItem` для получения значения из хранилища.

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

В пределах пользовательской функции (или в пределах любой другой части надстройки) можно использовать API `OfficeRuntime.displayWebDialogOptions`, чтобы отобразить диалоговое окно. Этот диалоговый API является альтернативой [Dialog API](../develop/dialog-api-in-office-add-ins.md), который можно использовать в пределах областей задач и команд надстройки, но не в пределах пользовательских функций.

### <a name="dialog-api-example"></a>Пример Dialog API

В приведенном ниже примере кода функция `getTokenViaDialog` использует функцию `displayWebDialogOptions` Dialog API для отображения диалогового окна.

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

Чтобы создать надстройку, которая будет работать на различных платформах (один из основных клиентов надстроек Office), вам не следует выполнять доступ к модели DOM в пользовательских функциях или использовать библиотеки, такие как jQuery, которые используют модель DOM. В Excel для Windows, где пользовательские функции используют среду выполнения JavaScript, пользовательские функции не могут выполнять доступ к модели DOM.

## <a name="see-also"></a>См. также

* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)
* [Журнал изменений пользовательских функций](custom-functions-changelog.md)
* [Руководство по настраиваемым функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)

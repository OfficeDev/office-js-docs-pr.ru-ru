---
ms.date: 05/08/2019
description: Сведения об основных сценариях при разработке пользовательских функций Excel, которые используют новую среду выполнения JavaScript.
title: Среда выполнения для пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: bb73ab2f20eadbac3f5fc97e272d69fe8bb983cd
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42689018"
---
# <a name="runtime-for-excel-custom-functions"></a>Среда выполнения для пользовательских функций Excel

Пользовательские функции используют новую среду выполнения JavaScript, отличающимся от среды выполнения, используемой другими частями надстройки, такими как область задач или другие элементы пользовательского интерфейса. Эта среда выполнения JavaScript предназначена для оптимизации производительности вычислений в пользовательских функциях и представляет новые API, которые можно использовать для выполнения стандартных действий в Интернете в пределах пользовательских функций, например для отправления запроса внешних данных или обмена данными через постоянное соединение с сервером.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Среда выполнения JavaScript также обеспечивает доступ к новым API в пространстве имен `OfficeRuntime`, которые могут быть использованы в пределах пользовательских функций или другими частями надстройки для хранения данных или отображения диалогового окна. В этой статье объясняется, как использовать такие API в пределах пользовательских функций, а также приводятся другие важные замечания, которые следует учитывать при разработке пользовательских функций.

## <a name="requesting-external-data"></a>Запрос внешних данных

В пределах пользовательской функции можно запрашивать внешние данные с помощью такого API, как [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или с помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.

В среде выполнения JavaScript, используемой пользовательскими функциями, XHR реализует дополнительные меры безопасности, требуя [одного и того же политики начала](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой [CORS](https://www.w3.org/TR/cors/).

Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST). Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`. Вы также можете `Content-Type` использовать заголовок в простой CORS, при условии, что тип контента `application/x-www-form-urlencoded`: `text/plain`, или `multipart/form-data`.

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
        
        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a>Получение данных через WebSockets

В пределах пользовательской функции можно использовать [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) для обмена данными через постоянное соединение с сервером. С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.

### <a name="websockets-example"></a>Пример WebSockets

Приведенный ниже примера кода устанавливает соединение `WebSocket`, а затем заносит в журнал каждое входящее сообщение от сервера.

```js
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Хранения данных и доступ к ним

В пределах функции (или в пределах любой другой части надстройки) можно хранить данные и выполнять доступ к ним с помощью объекта `OfficeRuntime.storage`. `Storage` — это постоянная незашифрованная система-хранилище пары "ключ-значение", обеспечивающая альтернативу хранилищу [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), которое нельзя использовать в пределах пользовательских функций. `Storage`предоставляет 10 МБ данных для каждого домена. Домены могут совместно использоваться несколькими надстройками.

`Storage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным. Например, токены для аутентификации пользователей могут храниться в `storage`, так как доступ к нему могут выполнять и пользовательская функция, и элементы пользовательского интерфейса надстройки, такие как область задач. Точно так же, если две надстройки используют один и тот же домен (например, www.contoso.com/addin1, www.contoso.com/addin2), им также разрешается обмен информацией в оба направления через `storage`. Обратите внимание, что надстройки, имеющие разные поддомены, будут иметь разные экземпляры `storage` (например, subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).

Так как `storage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".

Ниже указаны методы, доступные в объекте `storage`.

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

.[!NOTE]
> Нет метода для очистки всей информации (например, `clear`). Вместо этого вам следует использовать `removeItems` для одновременного удаления нескольких записей.

### <a name="officeruntimestorage-example"></a>Пример Оффицерунтиме. Storage

В следующем примере кода вызывается `OfficeRuntime.storage.setItem` функция для установки ключа и значения `storage`.

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>Дополнительные рекомендации

Чтобы создать надстройку, которая будет работать на различных платформах (один из основных клиентов надстроек Office), вам не следует выполнять доступ к модели DOM в пользовательских функциях или использовать библиотеки, такие как jQuery, которые используют модель DOM. В Excel для Windows, где пользовательские функции используют среду выполнения JavaScript, пользовательские функции не могут получить доступ к модели DOM.

## <a name="next-steps"></a>Дальнейшие действия
Узнайте, как [выполнять веб-запросы с пользовательскими функциями](custom-functions-web-reqs.md).

## <a name="see-also"></a>См. также

* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Архитектура пользовательских функций](custom-functions-architecture.md)
* [Отображение диалогового окна в пользовательских функциях](custom-functions-dialog.md)
* [Руководство по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md)

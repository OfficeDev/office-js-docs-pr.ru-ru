---
ms.date: 05/30/2019
description: Запрос, потоковая передача и отмена потоковой передачи внешних данных к книге с помощью пользовательских функций в Excel
title: Получение и обработка данных с помощью пользовательских функций
localization_priority: Priority
ms.openlocfilehash: 22f79c8b4e7e39569d3b955477e9397a053e1a8f
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910338"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>Получение и обработка данных с помощью пользовательских функций

Один из способов, используемых пользовательскими функциями для повышения эффективности Excel, состоит в получении данных из расположений помимо книг, например из Интернета или сервера (через WebSockets). Пользовательские функции могут запрашивать данные с помощью XHR и запросов `fetch`, а также выполнять потоковую передачу этих данных в режиме реального времени.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

В документах ниже показаны некоторые примеры веб-запросов, но для создания функции потоковой передачи используйте [Руководство по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md).

## <a name="functions-that-return-data-from-external-sources"></a>Функции, которые возвращают данные из внешних источников

Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:

1. Возвращать обещание JavaScript в Excel;
2. Устранять обещание с итоговым значением с помощью функции обратного вызова.

Можно запрашивать внешние данные с помощью такого API, как [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), или с помощью `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.

В среде выполнения пользовательских функций XHR реализует дополнительные меры по обеспечению безопасности, предъявляя в качестве требования [политику единого домена](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простой запрос [CORS](https://www.w3.org/TR/cors/).

Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST). Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`. Вы также можете использовать заголовок Content-Type в простом запросе CORS, если используется тип контента `application/x-www-form-urlencoded`, `text/plain` или `multipart/form-data`.

### <a name="xhr-example"></a>Пример XHR

В следующем примере кода функция **getTemperature** вызывает функцию sendWebRequest для получения температуры в определенной области на основе идентификатора термометра. Функция sendWebRequest использует XHR для отправления запроса GET в конечную точку, которая может предоставить данные.

```js
/**
 * Receives a temperature from an online source.
 * @customfunction
 * @param {number} thermometerID Identification number of the thermometer.
 */
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions.  
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

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

Другой пример запроса XHR с дополнительным контекстом см. в функции `getFile` в [этом файле](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) репозитория Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).

### <a name="fetch-example"></a>Пример получения данных

В следующем примере функция `stockPriceStream` использует символ тикера для получения цены акции каждые 1000 миллисекунд. Для получения дополнительных сведений об этом примере см. статью [Руководство по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function).

```js
/**
 * Streams a stock price.
 * @customfunction 
 * @param {string} ticker Stock ticker.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function stockPriceStream(ticker, invocation) {
    var updateFrequency = 1000 /* milliseconds*/;
    var isPending = false;

    var timer = setInterval(function() {
        // If there is already a pending request, skip this iteration:
        if (isPending) {
            return;
        }

        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        isPending = true;

        fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                invocation.setResult(parseFloat(text));
            })
            .catch(function(error) {
                invocation.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receive-data-via-websockets"></a>Получение данных через WebSockets

В пределах пользовательской функции можно использовать WebSockets для обмена данными через постоянное соединение с сервером. С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.

### <a name="websockets-example"></a>Пример WebSockets

Следующий примера кода устанавливает соединение WebSocket, а затем заносит в журнал каждое входящее сообщение от сервера.

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="make-a-streaming-function"></a>Создание функции потоковой передачи

Пользовательские функции потоковой передачи позволяют выводить данные в ячейки, которые повторно обновляются, не требуя от пользователя явно что-либо обновлять. Такие функции (например, функция из [руководства по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md)) могут быть полезны для проверки данных, обновляемых в реальном времени, из веб-службы.

Чтобы объявить функцию потоковой передачи, используйте тег комментария JSDoc `@stream`. Чтобы оповестить пользователей о том, что ваша функция может выполнять повторное вычисление с учетом новой информации, рекомендуем указать поток или другие сведения об этом в имени или описании функции.

В приведенном ниже примере показана функция потоковой передачи, которая увеличивает определенное число каждую секунду на указанное число.

```JS
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
CustomFunctions.associate("INC", increment);
```

>[!NOTE]
> Обратите внимание, что существует еще одна категория — так называемые отменяемые функции, которые *не* связаны с функциями потоковой передачи. В предыдущих версиях пользовательских функций требовалось объявлять `"cancelable": true` и `"streaming": true` в самостоятельно написанном коде JSON. С тех пор, как появились автоматически генерируемые метаданные, можно отменять только асинхронные пользовательские функции, возвращающие одно значение. Отменяемые функции позволяют прервать выполнение веб-запроса, используя [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation), чтобы решить, что делать после отмены. Для объявления отменяемых функций используется тег `@cancelable`.

### <a name="using-an-invocation-parameter"></a>Использование параметра вызова

Параметр `invocation` является по умолчанию последним в любой пользовательской функции. Параметр `invocation` содержит контекст о ячейке (например, ее адрес), а также позволяет использовать способы `setResult` и `onCanceled`. Эти методы определяют, что делает функция во время ее потоковой передачи (`setResult`) или отмены (`onCanceled`).

При использовании TypeScript требуется обработчик вызовов типа `CustomFunctions.StreamingInvocation` или `CustomFunctions.CancelableInvocation`.

### <a name="streaming-and-cancelable-function-example"></a>Пример потоковой и отменяемой функции
Следующий пример кода — это пользовательская функция, которая добавляет число к результату каждую секунду. Обратите внимание на следующие особенности этого кода:

- Excel отображает каждое новое значение автоматически с помощью метода `setResult`.
- Второй параметр ввода, вызов, не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".
- Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
    }
}
CustomFunctions.associate("INCREMENT", increment);
```

>[!NOTE]
> Excel отменяет выполнение функций в следующих случаях:
>
> - Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.
> - Когда изменяется один из аргументов (входных параметров) функции. В этом случае после отмены выполняется новый вызов функции.
> - Когда пользователь вручную вызывает пересчет. В этом случае после отмены выполняется новый вызов функции.

## <a name="next-steps"></a>Дальнейшие действия

* Ознакомьтесь с [разными типами параметров, которые могут использоваться функциями](custom-functions-parameter-options.md).
* Узнайте, как [пакетно обрабатывать несколько вызовов API](custom-functions-batching.md).

## <a name="see-also"></a>См. также

* [Пересчитываемые значения в функциях](custom-functions-volatile.md)
* [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Рекомендации по пользовательским функциям](custom-functions-best-practices.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)

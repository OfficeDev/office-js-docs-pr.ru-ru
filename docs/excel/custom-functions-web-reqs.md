---
ms.date: 04/20/2019
description: Запрос, потоковая передача и отмена потоковой передачи внешних данных к книге с помощью пользовательских функций в Excel
title: Обработка веб-запросов и других данных с помощью пользовательских функций (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: 2942ec56e46d6eb586b516eedab17c1eeb98d9c8
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353267"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a>Получение и обработка данных с помощью пользовательских функций

Один из способов, используемых пользовательскими функциями для повышения эффективности Excel, состоит в получении данных из расположений помимо книг, например из Интернета или сервера (через WebSockets). Пользовательские функции могут запрашивать данные с помощью XHR и получать запросы, а также выполнять потоковую передачу этих данных в режиме реального времени.

В документах ниже показаны некоторые примеры веб-запросов, но для создания функции потоковой передачи используйте [Руководство по пользовательским функциям](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).

## <a name="functions-that-return-data-from-external-sources"></a>Функции, которые возвращают данные из внешних источников

Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:

1. Возвращать обещание JavaScript в Excel;
2. Устранять обещание с итоговым значением с помощью функции обратного вызова.

Можно запрашивать внешние данные с помощью такого API, как [`Fetch`](https://developer.mozilla.org/ru-RU/docs/Web/API/Fetch_API), или с помощью `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ru-RU/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.

В среде выполнения пользовательских функций XHR реализует дополнительные меры по обеспечению безопасности, предъявляя в качестве требования [политику единого домена](https://developer.mozilla.org/ru-RU/docs/Web/Security/Same-origin_policy) и простой запрос [CORS](https://www.w3.org/TR/cors/).

Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST). Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`. Вы также можете использовать заголовок Content-Type в простом запросе CORS, если используется тип контента `application/x-www-form-urlencoded`, `text/plain` или `multipart/form-data`.

### <a name="xhr-example"></a>Пример XHR

В следующем примере кода функция **getTemperature** вызывает функцию sendWebRequest для получения температуры в определенной области на основе идентификатора термометра. Функция sendWebRequest использует XHR для отправления запроса GET в конечную точку, которая может предоставить данные.

```JavaScript
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

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

Другой пример запроса XHR с дополнительным контекстом см. в функции `getFile` в [этом файле](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) репозитория Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).

### <a name="fetch-example"></a>Пример получения данных

В следующем примере функция stockPriceStream использует символ тикера для получения цены акции каждые 1000 миллисекунд. Для получения дополнительных сведений об этом примере и соответствующего файла JSON см. статью [Руководство по пользовательским функциям](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function). 

```JavaScript
function stockPriceStream(ticker, handler) {
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
                handler.setResult(parseFloat(text));
            })
            .catch(function(error) {
                handler.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    handler.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receiving-data-via-websockets"></a>Получение данных через WebSockets

В пределах пользовательской функции можно использовать WebSockets для обмена данными через постоянное соединение с сервером. С помощью WebSockets ваша пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий, без необходимости специально запрашивать у сервера данные.

### <a name="websockets-example"></a>Пример WebSockets

Следующий примера кода устанавливает соединение WebSocket, а затем заносит в журнал каждое входящее сообщение от сервера.

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a>Потоковая передача функций

Потоковая передача пользовательских функций позволяет выводить данные в ячейки несколько раз с течением времени, избавляя пользователя от необходимости явным образом запрашивать обновление данных. Следующий пример кода — это пользовательская функция, которая добавляет число к результату каждую секунду. Обратите внимание на следующие особенности этого кода:

- Excel отображает каждое новое значение автоматически с помощью обратного вызова setResult.
- Второй параметр ввода (handler) не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".
- Обратный вызов onCanceled определяет функцию, которая выполняется при отмене функции. Вам необходимо реализовать уведомление об отмене следующим образом для любой функции потоковой передачи. Дополнительные сведения см. в разделе [Отмена функции](#canceling-a-function).

```JavaScript
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}

CustomFunctions.associate("INCREMENTVALUE", incrementValue);
```

Когда вы указываете метаданные для функции потоковой передачи в файле метаданных JSON, это можно автоматически создать с помощью тега комментария JSDOC `@streaming` в файле скрипта функции. Дополнительные сведения см. в статье [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).

## <a name="canceling-a-function"></a>Отмена функции

В некоторых случаях может потребоваться отмена выполнения пользовательских функций потоковой передачи, чтобы уменьшить использования пропускной способности, рабочей памяти и загрузку ЦП. Excel отменяет выполнение функций в следующих случаях:

- Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.
- Когда изменяется один из аргументов (входных параметров) функции. В этом случае после отмены выполняется новый вызов функции.
- Когда пользователь вручную вызывает пересчет. В этом случае после отмены выполняется новый вызов функции.

Чтобы сделать функцию отменяемой, нужно реализовать обработчик в коде функции с указанием действий при ее отмене. Также можно использовать тег комментария JSDOC `@cancelable` в файле скрипта функции. Дополнительные сведения см. в статье [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).

## <a name="see-also"></a>См. также

* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Рекомендации по пользовательским функциям](custom-functions-best-practices.md)
* [Журнал изменений пользовательских функций](custom-functions-changelog.md)

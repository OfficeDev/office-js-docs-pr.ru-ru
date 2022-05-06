---
ms.date: 05/02/2022
description: Запрос, потоковая передача и отмена потоковой передачи внешних данных в книгу с помощью пользовательских функций в Excel.
title: Получение и обработка данных с помощью пользовательских функций
ms.localizationpriority: medium
ms.openlocfilehash: 78f8f5f97bfeb690873091ff7c59555e1683c05f
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244851"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>Получение и обработка данных с помощью пользовательских функций

Одним из способов расширения возможностей Excel функций является получение данных из расположений, отличных от книги, таких как Интернет или сервер (через [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API)). Можно запрашивать внешние данные с помощью такого API, как [`Fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API), или с помощью `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![GIF пользовательской функции, которая выполняет потоковую передачу времени из API.](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a>Функции, которые возвращают данные из внешних источников

Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:

1. Возвращает [код JavaScript `Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) для Excel.
2. Разрешите значение `Promise` с окончательным значением с помощью функции обратного вызова.

### <a name="fetch-example"></a>Пример получения данных

В следующем примере кода `webRequest` функция достигает гипотетического внешнего API, который отслеживает количество людей на международной станции. Функция возвращает JavaScript `Promise` и использует `fetch` для запроса информации из гипотетического API. Полученные данные преобразуются в JSON `names` , а свойство преобразуется в строку, которая используется для разрешения обещания.

При разработке собственных функций может потребоваться выполнение действия, если веб-запрос не завершается своевременно. Также можно рассмотреть [совмещение нескольких запросов API](custom-functions-batching.md).

```JS
/**
 * Requests the names of the people currently on the International Space Station.
 * Note: This function requests data from a hypothetical URL. In practice, replace the URL with a data source for your scenario.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace"; // This is a hypothetical URL.
  return new Promise(function (resolve, reject) {
    fetch(url)
      .then(function (response){
        return response.json();
        }
      )
      .then(function (json) {
        resolve(JSON.stringify(json.names));
      })
  })
}
```

> [!NOTE]
> При использовании метода `fetch` не создаются вложенные обратные вызовы, что в некоторых случаях может быть предпочтительнее, чем использование метода XHR.

### <a name="xhr-example"></a>Пример XHR

В следующем примере кода `getStarCount` функция вызывает API Github для обнаружения количества звезд, присвоенных репозиторию определенного пользователя. Это асинхронная функция, которая возвращает JavaScript `Promise`. При получении данных из веб-вызова обещание разрешается, что возвращает данные в ячейку.

```TS
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */

async function getStarCount(userName: string, repoName: string) {

  const url = "https://api.github.com/repos/" + userName + "/" + repoName;

  let xhttp = new XMLHttpRequest();

  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        resolve(JSON.parse(xhttp.responseText).watchers_count);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
```

## <a name="make-a-streaming-function"></a>Создание функции потоковой передачи

Пользовательские функции потоковой передачи позволяют выводить данные в ячейки, которые повторно обновляются, не требуя от пользователя явно что-либо обновлять. Такие функции (например, функция из [руководства по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md)) могут быть полезны для проверки данных, обновляемых в реальном времени, из веб-службы.

Чтобы объявить функцию потоковой передачи, можно использовать один из следующих двух вариантов.

- Тег `@streaming` .
- Параметр `CustomFunctions.StreamingInvocation` вызова.

Следующий пример кода — это пользовательская функция, которая добавляет число к результату каждую секунду. Обратите внимание на указанные ниже аспекты этого кода.

- Excel отображает каждое новое значение автоматически с помощью метода `setResult`.
- Второй параметр ввода, `invocation`, не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".
- Обратный `onCanceled` вызов определяет функцию, которая выполняется при отмене функции.
- Потоковая передача не обязательно связана с выполнением веб-запроса. В этом случае функция не выполняет веб-запрос, но по-прежнему получает данные с заданным интервалом, поэтому она требует использования параметра потоковой `invocation` передачи.

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
```

## <a name="cancel-a-function"></a>Отмена функции

Excel отменяет выполнение функции в следующих ситуациях.

- Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.
- Когда изменяется один из аргументов (входных параметров) функции. В этом случае после отмены выполняется новый вызов функции.
- Когда пользователь вручную вызывает пересчет. В этом случае после отмены выполняется новый вызов функции.

Также можно настроить стандартное значение потоковой передачи, чтобы обрабатывать случаи выполнения запроса, когда вы находитесь в автономном режиме.

> [!NOTE]
> Существует также категория функций, называемых отменяемыми функциями, и они не _связаны_ с функциями потоковой передачи. Отменяются только асинхронные пользовательские функции, возвращаемые одним значением. Отменяемые функции позволяют прервать выполнение веб-запроса, используя [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation), чтобы решить, что делать после отмены. Для объявления отменяемых функций используется тег `@cancelable`.

### <a name="use-an-invocation-parameter"></a>Использование параметра вызова

Параметр `invocation` является по умолчанию последним в любой пользовательской функции. Параметр `invocation` предоставляет контекст ячейки (например, ее адрес и содержимое) и позволяет использовать `setResult` и методы `onCanceled` . Эти методы определяют, что делает функция во время ее потоковой передачи (`setResult`) или отмены (`onCanceled`).

Если вы используете TypeScript, обработчик вызова должен иметь тип или [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation).

## <a name="receiving-data-via-websockets"></a>Получение данных через WebSockets

В пределах пользовательской функции можно использовать [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API) для обмена данными через постоянное соединение с сервером. С помощью WebSocket пользовательская функция может открыть соединение с сервером, а затем автоматически получать сообщения от сервера при возникновении определенных событий без необходимости явного опроса сервера на наличие данных.

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

## <a name="next-steps"></a>Дальнейшие действия

- Ознакомьтесь с [разными типами параметров, которые могут использоваться функциями](custom-functions-parameter-options.md).
- Узнайте, как [пакетно обрабатывать несколько вызовов API](custom-functions-batching.md).

## <a name="see-also"></a>См. также

- [Пересчитываемые значения в функциях](custom-functions-volatile.md)
- [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
- [Создание метаданных JSON вручную для пользовательских функций](custom-functions-json.md)
- [Создание пользовательских функций в Excel](custom-functions-overview.md)
- [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)

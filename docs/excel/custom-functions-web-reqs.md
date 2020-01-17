---
ms.date: 01/14/2020
description: Запрос, потоковая передача и отмена потоковой передачи внешних данных к книге с помощью пользовательских функций в Excel
title: Получение и обработка данных с помощью пользовательских функций
localization_priority: Priority
ms.openlocfilehash: e57ffb43bfa58e0ff5219aa7fdbcae5fb30f5479
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217048"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>Получение и обработка данных с помощью пользовательских функций

Один из способов, используемых пользовательскими функциями для повышения эффективности Excel, состоит в получении данных из расположений помимо книг, например из Интернета или сервера (через WebSockets). Можно запрашивать внешние данные с помощью такого API, как [`Fetch`](https://developer.mozilla.org/ru-RU/docs/Web/API/Fetch_API), или с помощью `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ru-RU/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![GIF с пользовательской функцией, отправляющей время из API](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a>Функции, которые возвращают данные из внешних источников

Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:

1. Возвращать обещание JavaScript в Excel;
2. Устранять обещание с итоговым значением с помощью функции обратного вызова.

### <a name="fetch-example"></a>Пример получения данных

В следующем примере кода функция **webRequest** обращается к гипотетическому API "Число людей в космосе" компании Contoso, которое отслеживает число людей, находящихся в настоящее время на Международной космической станции. Функция возвращает обещание JavaScript и использует метод Fetch для запроса сведений из API. Итоговые данные преобразуются в формат JSON, а свойство `names` преобразуется в строку, использующуюся для разрешения обещания.

При разработке собственных функций может потребоваться выполнение действия, если веб-запрос не завершается своевременно. Также можно рассмотреть [совмещение нескольких запросов API](./custom-functions-batching.md).

```JS
/**
 * Requests the names of the people currently on the International Space Station from a hypothetical API.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace";
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

>[!NOTE]
>При использовании метода `Fetch` не создаются вложенные обратные вызовы, что в некоторых случаях может быть предпочтительнее, чем использование метода XHR.

### <a name="xhr-example"></a>Пример XHR

В среде выполнения пользовательских функций XHR реализует дополнительные меры по обеспечению безопасности, предъявляя в качестве требования [политику единого домена](https://developer.mozilla.org/ru-RU/docs/Web/Security/Same-origin_policy) и простой запрос [CORS](https://www.w3.org/TR/cors/).

Обратите внимание, что при реализации простых запросов CORS нельзя использовать файлы cookie и поддерживаются только простые методы (GET, HEAD, POST). Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`. Вы также можете использовать заголовок Content-Type в простом запросе CORS, если используется тип контента `application/x-www-form-urlencoded`, `text/plain` или `multipart/form-data`.

В следующем примере кода функция **getStarCount** вызывает API GitHub для определения числа звезд, присвоенных репозиторию определенного пользователя. Это асинхронная функция, возвращающая обещание JavaScript. При получении данных из веб-вызова обещание разрешается, что возвращает данные в ячейку.

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

Чтобы объявить функцию потоковой передачи, используйте тег `@streaming` или параметр вызова `CustomFunctions.StreamingInvocation`. Это указывает, что это функция потоковой передачи. Чтобы оповестить пользователей о том, что ваша функция может выполнять повторное вычисление с учетом новой информации, рекомендуем указать поток или другие сведения об этом в имени или описании функции.

Следующий пример кода — это пользовательская функция, которая добавляет число к результату каждую секунду. Обратите внимание на следующие особенности этого кода:

- Excel отображает каждое новое значение автоматически с помощью метода `setResult`.
- Второй параметр ввода, вызов, не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".
- Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции.
- Потоковая передача не обязательно связана с веб-запросом: в этом случае функция не выполняет веб-запрос, но по-прежнему получает данные через заданные интервалы, поэтому для нее требуется использовать параметр потоковой передачи `invocation`.

```js
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

В дополнение к сведениям об обратном вызове `onCanceled` вы также должны учитывать, что Excel отменяет выполнение функции в следующих случаях:

- Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.
- Когда изменяется один из аргументов (входных параметров) функции. В этом случае после отмены выполняется новый вызов функции.
- Когда пользователь вручную вызывает пересчет. В этом случае после отмены выполняется новый вызов функции.

Также можно настроить стандартное значение потоковой передачи, чтобы обрабатывать случаи выполнения запроса, когда вы находитесь в автономном режиме.

> [!NOTE]
> Обратите внимание, что существует еще одна категория — так называемые отменяемые функции, которые _не_ связаны с функциями потоковой передачи. В предыдущих версиях пользовательских функций требовалось объявлять `"cancelable": true` и `"streaming": true` в самостоятельно написанном коде JSON. С тех пор, как появились автоматически генерируемые метаданные, можно отменять только асинхронные пользовательские функции, возвращающие одно значение. Отменяемые функции позволяют прервать выполнение веб-запроса, используя [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation), чтобы решить, что делать после отмены. Для объявления отменяемых функций используется тег `@cancelable`.

### <a name="using-an-invocation-parameter"></a>Использование параметра вызова

Параметр `invocation` является по умолчанию последним в любой пользовательской функции. Параметр `invocation` содержит контекст о ячейке (например, ее адрес и содержимое), а также позволяет использовать методы `setResult` и `onCanceled`. Эти методы определяют, что делает функция во время ее потоковой передачи (`setResult`) или отмены (`onCanceled`).

При использовании TypeScript требуется обработчик вызовов типа `CustomFunctions.StreamingInvocation` или `CustomFunctions.CancelableInvocation`.

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

## <a name="next-steps"></a>Дальнейшие действия

- Ознакомьтесь с [разными типами параметров, которые могут использоваться функциями](custom-functions-parameter-options.md).
- Узнайте, как [пакетно обрабатывать несколько вызовов API](custom-functions-batching.md).

## <a name="see-also"></a>См. также

- [Пересчитываемые значения в функциях](custom-functions-volatile.md)
- [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
- [Метаданные пользовательских функций](custom-functions-json.md)
- [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
- [Создание пользовательских функций в Excel](custom-functions-overview.md)
- [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)

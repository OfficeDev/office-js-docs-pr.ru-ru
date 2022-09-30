---
ms.date: 09/09/2022
description: Объедините пользовательские функции в пакет, чтобы сократить количество обращений к удаленной службе через сеть.
title: Пакетирование обращений пользовательских функций к удаленной службе
ms.localizationpriority: medium
ms.openlocfilehash: f779351789350bbc591b1b5d7a975ff9f70cda26
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234923"
---
# <a name="batch-custom-function-calls-for-a-remote-service"></a>Пакетные вызовы пользовательских функций для удаленной службы

Если пользовательские функции обращаются к удаленной службе, можно использовать шаблон пакетирования для сокращения количества сетевых вызовов удаленной службы. Для уменьшения объема сетевых операций можно объединить все вызовы в один вызов веб-службы. Это идеальное решение при пересчете электронной таблицы.

Например если пользователь обращается к вашей пользовательской функции в 100 ячейках электронной таблицы, а затем пересчитывает электронную таблицу, эта функция будет выполняться 100 раз и делать 100 сетевых вызовов. С помощью шаблона пакетирования эти вызовы можно объединить так, чтобы делать 100 расчетов в течение одного сетевого вызова.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>Посмотреть готовый пример

Чтобы просмотреть завершенный пример, следуйте инструкциям в этой статье и вставьте примеры кода в собственный проект. Например, чтобы создать проект пользовательской функции для TypeScript, используйте генератор [Yeoman](../develop/yeoman-generator-overview.md) для надстроек Office, а затем добавьте в проект весь код из этой статьи. Запустите код и попробуйте его.

Кроме того, скачайте или просмотрите полный пример проекта в шаблоне [пакетной обработки пользовательских функций](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching). Если вы хотите просмотреть код в целом, прежде чем читать дальше, посмотрите на [файл сценария](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).

## <a name="create-the-batching-pattern-in-this-article"></a>Создание шаблона пакетирования в этой статье

Для реализации пакетирования пользовательских функций необходимо создать три основных раздела кода.

1. Операция [отправки](#add-the-_pushoperation-function) для добавления новой операции в пакет вызовов каждый раз, когда Excel вызывает пользовательскую функцию.
2. Функция [, которая выполняет удаленный запрос,](#make-the-remote-request) когда пакет готов.
3. [Код сервера для ответа на пакетный запрос](#process-the-batch-call-on-the-remote-service), вычисления всех результатов операции и возврата значений.

В следующих разделах вы узнаете, как создать код по одному примеру за раз. Рекомендуется создать новый проект пользовательских функций с помощью генератора [Yeoman](../develop/yeoman-generator-overview.md) для генератора надстроек Office. Сведения о создании проекта см. в статье ["Начало разработки пользовательских функций Excel"](../quickstarts/excel-custom-functions-quickstart.md). Вы можете использовать TypeScript или JavaScript.

## <a name="batch-each-call-to-your-custom-function"></a>Включение в пакет каждого вызова пользовательской функции

Ваши пользовательские функции вызывают удаленную службу для выполнения различных операций и вычисления требуемого результата. Это дает возможность сохранения каждой запрашиваемой операции в пакете. Далее вы узнаете, как создать функцию `_pushOperation` для пакетной обработки операций. Сначала посмотрим на следующий пример кода, где показан вызов `_pushOperation` из пользовательской функции.

В следующем примере пользовательская функция выполняет деление, обращаясь для этой операции к удаленной службе. Она вызывает `_pushOperation` для включения операции вместе с другими операциями в пакет для удаленной службы. Операция здесь называется **div2**. Можно использовать для операций любую схему именования, если только в удаленной службе используется такая же схема (дополнительно об удаленной службе см. далее). Кроме того передаются аргументы, необходимые удаленной службе для выполнения операции.

### <a name="add-the-div2-custom-function"></a>Добавление пользовательской функции div2

Добавьте следующий код в файл **functions.js** **functions.ts** (в зависимости от того, использовали ли вы JavaScript или TypeScript).

```javascript
/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend, divisor) {
  return _pushOperation("div2", [dividend, divisor]);
}
```

### <a name="add-global-variables-for-tracking-batch-requests"></a>Добавление глобальных переменных для отслеживания пакетных запросов

Затем добавьте две глобальные переменные **в файлfunctions.js** **functions.ts** . `_isBatchedRequestScheduled` важно позже для синхронизации пакетных вызовов к удаленной службе.

```javascript
let _batch = [];
let _isBatchedRequestScheduled = false;
```

### <a name="add-the-_pushoperation-function"></a>Добавление функции `_pushOperation`

Когда Excel вызывает пользовательскую функцию, необходимо отправить операцию в пакетный массив. В следующем **_pushOperation** кода функции показано, как добавить новую операцию из пользовательской функции. Здесь создается новый элемент пакета, новое обещание для выполнения или отклонения операции, и элемент вставляется в пакетный массив.

В данном коде также проверяется, является ли пакет плановым. В этом примере выполнение пакете планируется каждые 100 мс. При необходимости этот интервал можно изменить. Чем значение выше, тем больше размер пакета, отправляемого в удаленную службу, и тем дольше пользователь должен ждать результатов. При низком значении в удаленную службу отправляется больше пакетов, но зато время ожидания снижается.

Функция создает объект **invocationEntry** , содержащий строковое имя выполняемой операции. Например, если у вас две пользовательские функции с именами `multiply` и `divide`, их можно использовать как имена операции в элементах пакета. `args` содержит аргументы, которые были переданы в пользовательскую функцию из Excel. И, наконец, `resolve` или методы `reject` сохраняют обещание, содержащее сведения, возвращаемые удаленной службой.

Добавьте следующий код в файл **functions.js** **functions.ts** .

```javascript
// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  console.log("pushOperation");
  const invocationEntry = {
    operation: op, // e.g., sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g., 100 ms.
  if (!_isBatchedRequestScheduled) {
    console.log("schedule remote request");
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a>Проведение удаленного запроса

Цель функции `_makeRemoteRequest` – передать пакет операций в удаленную службу, а затем возвратить результаты в каждую пользовательскую функцию. Сначала она создает копию пакетного массива. Это позволит сразу же начинать включение параллельных вызовов пользовательской функции из Excel в новый массив. Затем копия преобразуется в более простой массив, который не содержит информацию обещания. Не имеет смысла передавать обещания в удаленную службу, так как они не будут работать. Метод `_makeRemoteRequest` будет отклонять или выполнять каждое обещание в зависимости от того, что возвратит удаленная служба.

Добавьте следующий код в файл **functions.js** **functions.ts** .

```javascript
// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  try{
  console.log("makeRemoteRequest");
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });
  console.log("makeRemoteRequest2");
  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      console.log("responseBatch in fetchFromRemoteService");
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
          console.log("rejecting promise");
        } else {
          console.log("fulfilling promise");
          console.log(response);

          batchCopy[index].resolve(response.result);
        }
      });
    });
    console.log("makeRemoteRequest3");
  } catch (error) {
    console.log("error name:" + error.name);
    console.log("error message:" + error.message);
    console.log(error);
  }
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a>Переделка `_makeRemoteRequest` для вашего собственного решения

Функция `_makeRemoteRequest` вызывает метод `_fetchFromRemoteService`, который, как будет видно позже, всего лишь имитирует удаленную службу. Это упрощает изучение и выполнение кода в данной статье. Но если вы хотите использовать этот код для фактической удаленной службы, необходимо внести следующие изменения.

- Выберите способ сериализации пакетных операций по сети. Например может потребоваться поместить массива в текст JSON.
- Вместо вызова `_fetchFromRemoteService` следует сделать сетевой вызов удаленной службы с передачей пакета операций.

## <a name="process-the-batch-call-on-the-remote-service"></a>Обработка пакетного вызова в удаленной службе

Последний шаг – это выполнение пакетного вызова в удаленной службе. В следующем примере кода показана функция `_fetchFromRemoteService`. Эта функция распаковывает каждую операцию, выполняет указанную операцию и возвращает результат. Для учебных целей в данной статье применяется функция `_fetchFromRemoteService`, которая запускается в вашей веб-надстройке и имитирует удаленную службу. Этот код можно добавить в файл **functions.js** **functions.ts** , чтобы можно было изучить и запустить весь код, приведенный в этой статье, без необходимости настройки фактической удаленной службы.

Добавьте следующий код в файл **functions.js** **functions.ts** .

```javascript
// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a promise that may contain a batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
async function _fetchFromRemoteService(requestBatch) {
  // Simulate a slow network request to the server.
  console.log("_fetchFromRemoteService");
  await pause(1000);
  console.log("postpause");
  return requestBatch.map((request) => {
    console.log("requestBatch server side");
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myResult = args[0] * args[1];
        console.log(myResult);
        return {
          result: myResult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms) {
  console.log("pause");
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a>Переделка `_fetchFromRemoteService` для действующей удаленной службы

Чтобы изменить функцию `_fetchFromRemoteService` для запуска в динамической удаленной службе, внесите следующие изменения.

- В зависимости от платформы используемого сервера (Node.js или другая) сопоставьте сетевой вызов клиента с этой функцией.
- Удалите функцию `pause`, которая имитирует задержку в сети.
- Измените объявление функции так, чтобы она работала с переданным параметром, если параметр изменяется для целей сети. Например, это может быть не массив а текст JSON, содержащий требуемые пакетные операции.
- Переделайте функцию для выполнения операций (или вызова функций, которые выполняют операции).
- Примените подходящий механизм проверки подлинности. Убедитесь, что доступ к функции есть только у предусмотренных вами вызывающих пользователей.
- Поместите код в удаленную службу.

## <a name="next-steps"></a>Дальнейшие действия

Узнайте о [различных параметрах](custom-functions-parameter-options.md), которые можно использовать в пользовательских функциях. Или узнайте, что лежит в основе [веб-вызова через пользовательскую функцию](custom-functions-web-reqs.md).

## <a name="see-also"></a>Дополнительные ресурсы

- [Пересчитываемые значения в функциях](custom-functions-volatile.md)
- [Создание пользовательских функций в Excel](custom-functions-overview.md)
- [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)

---
ms.date: 07/08/2021
description: Объедините пользовательские функции в пакет, чтобы сократить количество обращений к удаленной службе через сеть.
title: Пакетирование обращений пользовательских функций к удаленной службе
ms.localizationpriority: medium
ms.openlocfilehash: 0cf1a1df922a08f63af80498da2e357d285775e9
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074233"
---
# <a name="batch-custom-function-calls-for-a-remote-service"></a>Пакетные пользовательские вызовы функции для удаленной службы

Если пользовательские функции обращаются к удаленной службе, можно использовать шаблон пакетирования для сокращения количества сетевых вызовов удаленной службы. Для уменьшения объема сетевых операций можно объединить все вызовы в один вызов веб-службы. Это идеальное решение при пересчете электронной таблицы.

Например если пользователь обращается к вашей пользовательской функции в 100 ячейках электронной таблицы, а затем пересчитывает электронную таблицу, эта функция будет выполняться 100 раз и делать 100 сетевых вызовов. С помощью шаблона пакетирования эти вызовы можно объединить так, чтобы делать 100 расчетов в течение одного сетевого вызова.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>Посмотреть готовый пример

Вы можете изучить эту статью и вставить примеры кода в свой проект. Например можно создать в [генераторе Yo Office](https://github.com/OfficeDev/generator-office) проект пользовательской функции для TypeScript, вставить в него весь код из этой статьи, а затем запустить код и посмотреть на результаты его работы.

Также можно загрузить или просмотреть готовый образец проекта на странице [Custom function batching pattern (Пакетирование пользовательских функций)](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching). Если вы хотите просмотреть код в целом, прежде чем читать дальше, посмотрите на [файл сценария](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).

## <a name="create-the-batching-pattern-in-this-article"></a>Создание шаблона пакетирования в этой статье

Для реализации пакетирования пользовательских функций необходимо создать три основных раздела кода.

1. Push-операция для включения новой операции в пакет вызовов каждый раз, когда Excel вызывает пользовательскую функцию.
2. Функция, которая делает удаленный запрос, когда пакет готов.
3. Код сервера для отклика на пакетный запрос, вычисления результатов всех операций и возвращения значений.

В следующих разделах будет показано создание кода по одному примеру за раз. Добавьте каждый пример кода в файл **functions.ts**. Рекомендуем создавать пользовательские функции заново в генераторе Yo Office. Для создания проекта обратитесь к статье [Начало разработки пользовательских функций Excel](../quickstarts/excel-custom-functions-quickstart.md) и используйте TypeScript вместо JavaScript.

## <a name="batch-each-call-to-your-custom-function"></a>Включение в пакет каждого вызова пользовательской функции

Ваши пользовательские функции вызывают удаленную службу для выполнения различных операций и вычисления требуемого результата. Это дает возможность сохранения каждой запрашиваемой операции в пакете. Далее вы узнаете, как создать функцию `_pushOperation` для пакетной обработки операций. Сначала посмотрим на следующий пример кода, где показан вызов `_pushOperation` из пользовательской функции.

В следующем примере пользовательская функция выполняет деление, обращаясь для этой операции к удаленной службе. Она вызывает `_pushOperation` для включения операции вместе с другими операциями в пакет для удаленной службы. Операция здесь называется **div2**. Можно использовать для операций любую схему именования, если только в удаленной службе используется такая же схема (дополнительно об удаленной службе см. далее). Кроме того передаются аргументы, необходимые удаленной службе для выполнения операции.

### <a name="add-the-div2-custom-function-to-functionsts"></a>Добавление пользовательской функции div2 в functions.ts

```typescript
/**
 * @CustomFunction
 * Divides two numbers using batching
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend: number, divisor: number) {
  return _pushOperation(
    "div2",
    [dividend, divisor]
  );
}
```

После этого следует определить пакетный массив, в котором будут храниться все операции, предназначенные для передачи в одном сетевом вызове. В приведенном ниже коде показано, как определить интерфейс, описывающий каждый элемент пакета в массиве. Интерфейс определяет операцию, которая представляет собой строку-имя запускаемой операции. Например, если у вас две пользовательские функции с именами `multiply` и `divide`, их можно использовать как имена операции в элементах пакета. `args` будет содержать аргументы, переданные в пользовательскую функцию из Excel. И, наконец, в `resolve` или `reject` будет храниться обещание с информацией, возвращаемой удаленной службой.

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

Далее мы создадим пакетный массив, использующий предыдущий интерфейс. Чтобы знать, является ли пакет плановым или нет, создадим переменную `_isBatchedRequestSchedule`. Она понадобится позже для планирования пакетных вызовов удаленной службы.

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

Наконец, когда Excel вызывает пользовательскую функцию, необходимо отправить операцию в пакетный массив. В следующем коде показано, как добавить новую операцию из пользовательской функции. Здесь создается новый элемент пакета, новое обещание для выполнения или отклонения операции, и элемент вставляется в пакетный массив.

В данном коде также проверяется, является ли пакет плановым. В этом примере выполнение пакете планируется каждые 100 мс. При необходимости этот интервал можно изменить. Чем значение выше, тем больше размер пакета, отправляемого в удаленную службу, и тем дольше пользователь должен ждать результатов. При низком значении в удаленную службу отправляется больше пакетов, но зато время ожидания снижается.

### <a name="add-the-_pushoperation-function-to-functionsts"></a>Добавление функции `_pushOperation` в functions.ts

```typescript
function _pushOperation(op: string, args: any[]) {
  // Create an entry for your custom function.
  const invocationEntry: IBatchEntry = {
    operation: op, // e.g. sum
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
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a>Проведение удаленного запроса

Цель функции `_makeRemoteRequest` – передать пакет операций в удаленную службу, а затем возвратить результаты в каждую пользовательскую функцию. Сначала она создает копию пакетного массива. Это позволит сразу же начинать включение параллельных вызовов пользовательской функции из Excel в новый массив. Затем копия преобразуется в более простой массив, который не содержит информацию обещания. Не имеет смысла передавать обещания в удаленную службу, так как они не будут работать. Метод `_makeRemoteRequest` будет отклонять или выполнять каждое обещание в зависимости от того, что возвратит удаленная служба.

### <a name="add-the-following-_makeremoterequest-method-to-functionsts"></a>Добавление следующего метода `_makeRemoteRequest` в functions.ts

```typescript
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });

  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
        } else {
          console.log(response);
          batchCopy[index].resolve(response.result);
        }
      });
    });
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a>Переделка `_makeRemoteRequest` для вашего собственного решения

Функция `_makeRemoteRequest` вызывает метод `_fetchFromRemoteService`, который, как будет видно позже, всего лишь имитирует удаленную службу. Это упрощает изучение и выполнение кода в данной статье. Но если вы хотите использовать этот код для фактической удаленной службы, необходимо внести следующие изменения.

- Выберите способ сериализации пакетных операций по сети. Например может потребоваться поместить массива в текст JSON.
- Вместо вызова `_fetchFromRemoteService` следует сделать сетевой вызов удаленной службы с передачей пакета операций.

## <a name="process-the-batch-call-on-the-remote-service"></a>Обработка пакетного вызова в удаленной службе

Последний шаг – это выполнение пакетного вызова в удаленной службе. В следующем примере кода показана функция `_fetchFromRemoteService`. Эта функция распаковывает каждую операцию, выполняет указанную операцию и возвращает результат. Для учебных целей в данной статье применяется функция `_fetchFromRemoteService`, которая запускается в вашей веб-надстройке и имитирует удаленную службу. Этот код можно добавить в файл **functions.ts**, чтобы изучать и запускать его, не создавая настоящую удаленную службу.

### <a name="add-the-following-_fetchfromremoteservice-function-to-functionsts"></a>Добавление следующей функции `_fetchFromRemoteService` в functions.ts

```typescript
async function _fetchFromRemoteService(
  requestBatch: Array<{ operation: string, args: any[] }>
): Promise<IServerResponse[]> {
  // Simulate a slow network request to the server;
  await pause(1000);

  return requestBatch.map((request): IServerResponse => {
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myresult = args[0] * args[1];
        console.log(myresult);
        return {
          result: myresult
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

function pause(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a>Переделка `_fetchFromRemoteService` для действующей удаленной службы

Чтобы изменить функцию, которая будет работать в вашей удаленной службе `_fetchFromRemoteService` в прямом эфире, внести следующие изменения.

- В зависимости от платформы используемого сервера (Node.js или другая) сопоставьте сетевой вызов клиента с этой функцией.
- Удалите функцию `pause`, которая имитирует задержку в сети.
- Измените объявление функции так, чтобы она работала с переданным параметром, если параметр изменяется для целей сети. Например, это может быть не массив а текст JSON, содержащий требуемые пакетные операции.
- Переделайте функцию для выполнения операций (или вызова функций, которые выполняют операции).
- Примените подходящий механизм проверки подлинности. Убедитесь, что доступ к функции есть только у предусмотренных вами вызывающих пользователей.
- Поместите код в удаленную службу.

## <a name="next-steps"></a>Дальнейшие действия

Узнайте о [различных параметрах](custom-functions-parameter-options.md), которые можно использовать в пользовательских функциях. Или узнайте, что лежит в основе [веб-вызова через пользовательскую функцию](custom-functions-web-reqs.md).

## <a name="see-also"></a>Дополнительные ресурсы

* [Пересчитываемые значения в функциях](custom-functions-volatile.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)

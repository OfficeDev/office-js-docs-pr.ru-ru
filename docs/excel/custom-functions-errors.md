---
title: Обработка и возвращение ошибок из настраиваемой функции
description: 'Обработка и возврат таких ошибок, как #ПУСТО!, из настраиваемой функции.'
ms.date: 08/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: a2f93059f9082bc5a53c07159c9356a41cf16729
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/18/2021
ms.locfileid: "59443547"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Обработка и возвращение ошибок из настраиваемой функции

Если что-то пойдет не так во время работы настраиваемой функции, возвращайте ошибку для информирования пользователя. Если у вас есть определенные требования к параметрам, например только положительные номера, проверьте параметры и в случае их неправильной ошибки. Вы также можете использовать блок, чтобы поймать все ошибки, которые происходят [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) во время работы настраиваемой функции.

## <a name="detect-and-throw-an-error"></a>Обнаружение и возвращение ошибки

Рассмотрим случай, когда необходимо убедиться, что параметр почтовый индекс находится в правильном формате для работы настраиваемой функции. В следующей пользовательской функции используется регулярное выражение для проверки почтового индекса. Если формат почтового кода правильный, он будет искать город с помощью другой функции и возвращать значение. Если формат не действителен, функция возвращает `#VALUE!` ошибку в ячейку.

```typescript
/**
* Gets a city name for the given U.S. zip code.
* @customfunction
* @param {string} zipCode
* @returns The city of the zip code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## <a name="the-customfunctionserror-object"></a>Объект CustomFunctions.Error

Объект [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) используется для возврата ошибки обратно в ячейку. При создании объекта укажите, какую ошибку следует использовать, выбрав одно из следующих `ErrorCode` значений.

|Значение перечисления ErrorCode  |Значение ячейки Excel  |Описание  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | Функция пытается разделить на ноль. |
|`invalidName`    | `#NAME?`  | В имени функции имеется опечатка. Обратите внимание, что эта ошибка поддерживается как настраиваемая ошибка ввода функции, но не как настраиваемая ошибка вывода функции. |
|`invalidNumber`  | `#NUM!`   | Существует проблема с номером в формуле. |
|`invalidReference` | `#REF!` | Функция относится к недействительной ячейке. Обратите внимание, что эта ошибка поддерживается как настраиваемая ошибка ввода функции, но не как настраиваемая ошибка вывода функции.|
|`invalidValue`   | `#VALUE!` | Значение в формуле имеет неправильный тип. |
|`notAvailable`   | `#N/A`    | Функция или служба недоступны. |
|`nullReference`  | `#NULL!`  | Диапазоны в формуле не пересекаются. |

В следующем примере кода показано, как создать и вернуть ошибку для неверного числа (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

Ошибки `#VALUE!` `#N/A` также поддерживают пользовательские сообщения об ошибках. Пользовательские сообщения об ошибке отображаются в меню индикатора ошибок, к которому можно получить доступ, зависая над флагом ошибки на каждой ячейке с ошибкой. В следующем примере показано, как вернуть пользовательское сообщение об `#VALUE!` ошибке.

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### <a name="handle-errors-when-working-with-dynamic-arrays"></a>Обработка ошибок при работе с динамическими массивами

В дополнение к возвращению одной ошибки настраиваемая функция может выводить динамический массив, включающий ошибку. Например, настраиваемая функция может выводить массив `[1],[#NUM!],[3]` . В следующем примере кода показано, как ввести три параметра в настраиваемую функцию, заменить один из параметров ввода ошибкой, а затем вернуть двухмерный массив с результатами обработки каждого параметра `#NUM!` ввода.

```js
/**
* Returns the #NUM! error as part of a 2-dimensional array.
* @customfunction
* @param {number} first First parameter.
* @param {number} second Second parameter.
* @param {number} third Third parameter.
* @returns {number[][]} Three results, as a 2-dimensional array.
*/
function returnInvalidNumberError(first, second, third) {
  // Use the `CustomFunctions.Error` object to retrieve an invalid number error.
  var error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  var firstResult = first;
  var secondResult =  error;
  var thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### <a name="errors-as-custom-function-inputs"></a>Ошибки в качестве пользовательских входных данных функций

Настраиваемая функция может оценить, даже если диапазон ввода содержит ошибку. Например, настраиваемая функция может принимать диапазон **A2:A7** в качестве ввода, даже если **A6:A7** содержит ошибку.

Для обработки входных данных, содержащих ошибки, настраиваемая функция должна иметь свойство метаданных `allowErrorForDataTypeAny` `true` JSON. Дополнительные сведения см. в руководстве по созданию [метаданных JSON для пользовательских](custom-functions-json.md#metadata-reference) функций.

> [!IMPORTANT]
> Свойство `allowErrorForDataTypeAny` можно использовать только с созданными [вручную метаданными JSON.](custom-functions-json.md) Это свойство не работает с процессом автогенерации метаданных JSON.

## <a name="use-trycatch-blocks"></a>Использование `try...catch` блоков

В общем случае используйте [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) блоки в настраиваемой функции, чтобы поймать возможные ошибки. Если вы не обрабатываете исключения в коде, они будут возвращены в Excel. По умолчанию Excel `#VALUE!` возвращается за невыполнение ошибок или исключений.

В следующем примере кода пользовательская функция создает запрос fetch в службу REST. Возможно, что вызов завершится сбоем (например, если служба REST возвращает ошибку или не работает сеть). Если это произойдет, настраиваемая функция вернется, `#N/A` чтобы указать, что веб-вызов не удалось.

```typescript
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + commentID;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    })
}
```

## <a name="next-steps"></a>Дальнейшие действия

Узнайте, как [устранять проблемы с пользовательскими функциями](custom-functions-troubleshooting.md).

## <a name="see-also"></a>Дополнительные ресурсы

* [Отладка пользовательских функций](custom-functions-debugging.md)
* [Наборы обязательных элементов пользовательских функций](../reference/requirement-sets/custom-functions-requirement-sets.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)

---
title: Обработка и возврат ошибок из пользовательской функции
description: 'Обработка и возврат таких ошибок, как #ПУСТО!, из пользовательской функции.'
ms.date: 08/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: c93c13aac1457e776ba8441565c11a23074a8d97
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958568"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Обработка и возврат ошибок из пользовательской функции

Если во время выполнения пользовательской функции что-то пошло не так, сообщите пользователю об ошибке. Если у вас есть определенные требования к параметрам, например только положительные числа, проверьте параметры и выдайте ошибку, если они не верны. Блок также можно использовать для [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) перехвата любых ошибок, которые возникают во время выполнения пользовательской функции.

## <a name="detect-and-throw-an-error"></a>Обнаружение и возвращение ошибки

Рассмотрим случай, когда необходимо убедиться, что параметр почтового индекса имеет правильный формат для работы пользовательской функции. В следующей пользовательской функции используется регулярное выражение для проверки почтового индекса. Если формат почтового индекса правильный, он будет искать город с помощью другой функции и возвращать значение. Если формат не является допустимым, функция возвращает ошибку `#VALUE!` в ячейку.

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

Объект [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) используется для возврата ошибки в ячейку. При создании объекта укажите `ErrorCode` , какую ошибку вы хотите использовать, выбрав одно из следующих значений перечисления.

|Значение перечисления ErrorCode  |Значение ячейки Excel  |Описание  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | Функция пытается разделить на ноль. |
|`invalidName`    | `#NAME?`  | В имени функции есть опечатка. Обратите внимание, что эта ошибка поддерживается как пользовательская ошибка ввода функции, но не как пользовательская ошибка вывода функции. |
|`invalidNumber`  | `#NUM!`   | В формуле возникла проблема с числом. |
|`invalidReference` | `#REF!` | Функция ссылается на недопустимую ячейку. Обратите внимание, что эта ошибка поддерживается как пользовательская ошибка ввода функции, но не как пользовательская ошибка вывода функции.|
|`invalidValue`   | `#VALUE!` | Значение в формуле имеет неправильный тип. |
|`notAvailable`   | `#N/A`    | Функция или служба недоступны. |
|`nullReference`  | `#NULL!`  | Диапазоны в формуле не пересекаются. |

В следующем примере кода показано, как создать и вернуть ошибку для неверного числа (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

Ошибки `#VALUE!` и `#N/A` сообщения об ошибках также поддерживают пользовательские сообщения об ошибках. Пользовательские сообщения об ошибках отображаются в меню индикатора ошибок, к которому можно получить доступ, наведя указатель мыши на флаг ошибки в каждой ячейке с ошибкой. В следующем примере показано, как вернуть пользовательское сообщение об ошибке `#VALUE!` .

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### <a name="handle-errors-when-working-with-dynamic-arrays"></a>Обработка ошибок при работе с динамическими массивами

В дополнение к возврату одной ошибки пользовательская функция может выведите динамический массив, содержащий ошибку. Например, пользовательская функция может выведите массив `[1],[#NUM!],[3]`. В следующем примере кода показано, как ввести три параметра в пользовательскую функцию, `#NUM!` заменить один из входных параметров ошибкой, а затем вернуть двумерный массив с результатами обработки каждого входного параметра.

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
  const error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  const firstResult = first;
  const secondResult =  error;
  const thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### <a name="errors-as-custom-function-inputs"></a>Ошибки в качестве входных данных пользовательской функции

Пользовательская функция может вычислять, даже если входной диапазон содержит ошибку. Например, пользовательская функция может принимать диапазон **A2:A7** в качестве входных данных, даже если **A6:A7** содержит ошибку.

Для обработки входных данных, содержащих ошибки, пользовательской функции должно быть задано свойство метаданных `allowErrorForDataTypeAny` `true`JSON. Дополнительные сведения см. в статье [о создании метаданных JSON вручную для пользовательских](custom-functions-json.md#metadata-reference) функций.

> [!IMPORTANT]
> Свойство `allowErrorForDataTypeAny` можно использовать только с созданными [вручную метаданными JSON](custom-functions-json.md). Это свойство не работает с автоматически созданным процессом метаданных JSON.

## <a name="use-trycatch-blocks"></a>Использование блоков `try...catch`

Как правило, блоки [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) в пользовательской функции используются для перехвата возможных ошибок. Если не обрабатывать исключения в коде, они будут возвращены в Excel. По умолчанию Excel возвращает необработанные `#VALUE!` ошибки или исключения.

В следующем примере кода пользовательская функция создает запрос fetch в службу REST. Возможно, что вызов завершится сбоем (например, если служба REST возвращает ошибку или не работает сеть). В этом случае пользовательская функция вернет `#N/A` значение, указывающее, что веб-вызов завершился сбоем.

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
* [Наборы обязательных элементов пользовательских функций](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)

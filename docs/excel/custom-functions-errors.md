---
ms.date: 09/23/2020
description: 'Обработка и возврат таких ошибок, как #ПУСТО!, из настраиваемой функции.'
title: Обработка и возвращение ошибок из настраиваемой функции
localization_priority: Normal
ms.openlocfilehash: 2822b3e93f7e5f16410e49d4414110e37172f3569b8f3c5d7d4dd98d5c5ecf6a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079677"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Обработка и возвращение ошибок из настраиваемой функции

Если что-то пойдет не так во время работы настраиваемой функции, возвращайте ошибку для информирования пользователя. Если у вас есть определенные требования к параметрам, например только положительные номера, проверьте параметры и в случае их неправильной ошибки. Можно также использовать блок `try`-`catch`, чтобы отслеживать любые ошибки, возникающие при выполнении пользовательской функции.

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

## <a name="use-try-catch-blocks"></a>Использование блоков try-catch

В общем случае используйте `try` - `catch` блоки в настраиваемой функции, чтобы поймать возможные ошибки. Если в коде не обрабатываются исключения, они будут возвращаться в Excel. По умолчанию Excel `#VALUE!` возвращается за невыполнение ошибок или исключений.

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
* [Требования к настраиваемым функциям](custom-functions-requirement-sets.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)

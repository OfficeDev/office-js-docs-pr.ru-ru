---
ms.date: 09/21/2020
description: 'Обработка и возврат таких ошибок, как #ПУСТО!, из пользовательской функции.'
title: Обработка и возврат ошибок пользовательской функции
localization_priority: Normal
ms.openlocfilehash: 58c2ab432a4525f660e2d89735fd3add6e76fa7f
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175530"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Обработка и возврат ошибок пользовательской функции

Если при выполнении пользовательской функции возникла проблема, возвращайте ошибку, чтобы уведомить пользователя. Если у вас есть особые требования к параметрам, например, только положительные числа, проверьте параметры и вызовите ошибку, если они неправильные. Можно также использовать блок `try`-`catch`, чтобы отслеживать любые ошибки, возникающие при выполнении пользовательской функции.

## <a name="detect-and-throw-an-error"></a>Обнаружение и возвращение ошибки

Рассмотрим ситуацию, в которой необходимо убедиться, что параметр ZIP-кода имеет правильный формат, чтобы пользовательская функция работала. В следующей пользовательской функции используется регулярное выражение для проверки почтового индекса. Если формат ZIP-кода правильный, то он будет искать город с помощью другой функции и возвращать значение. Если формат не является допустимым, функция возвращает `#VALUE!` ошибку в ячейку.

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

Объект [CustomFunctions. Error](/javascript/api/custom-functions-runtime/customfunctions.error) используется для возврата к ячейке ошибки. При создании объекта укажите, какую ошибку следует использовать, выбрав одно из следующих `ErrorCode` значений перечисления.


|Значение перечисления ErrorCode  |Значение ячейки Excel  |Смысл  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | Обратите внимание, что JavaScript позволяет делить на нуль, поэтому при создании обработчика ошибок нужно внимательно определить это условие. |
|`invalidName`    | `#NAME?`  | В имени функции присутствует опечатка. Обратите внимание, что эта ошибка поддерживается как ошибка ввода пользовательской функции, но не в качестве ошибки вывода пользовательской функции. | 
|`invalidNumber`  | `#NUM!`   | Возникла проблема с числом в формуле. |
|`invalidReference` | `#REF!` | Функция ссылается на недопустимую ячейку. Обратите внимание, что эта ошибка поддерживается как ошибка ввода пользовательской функции, но не в качестве ошибки вывода пользовательской функции.|
|`invalidValue`   | `#VALUE!` | Недопустимый тип значения в формуле. |
|`notAvailable`   | `#N/A`    | Функция или служба недоступна. |
|`nullReference`  | `#NULL!`  | Диапазоны в формуле не пересекаются. |

В следующем примере кода показано, как создать и вернуть ошибку для неверного числа (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

`#VALUE!` `#N/A` Кроме того, ошибки также поддерживают настраиваемые сообщения об ошибках. Настраиваемые сообщения об ошибках отображаются в меню индикации ошибки, доступ к которому осуществляется при наведении курсора на флаг ошибки в каждой ячейке с ошибкой. В приведенном ниже примере показано, как вернуть настраиваемое сообщение об ошибке с `#VALUE!` ошибкой.

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>Использование блоков try-catch

В общем случае `try` - `catch` для перехвата возможных ошибок используйте блоки в пользовательской функции. Если в коде не обрабатываются исключения, они будут возвращаться в Excel. По умолчанию Excel возвращает `#VALUE!` для необработанных ошибок или исключений.

В следующем примере кода пользовательская функция создает запрос fetch в службу REST. Возможно, что вызов завершится сбоем (например, если служба REST возвращает ошибку или не работает сеть). В этом случае пользовательская функция вернется, `#N/A` чтобы указать, что веб-вызов завершился ошибкой.


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

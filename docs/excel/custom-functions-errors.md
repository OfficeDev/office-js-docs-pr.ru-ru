---
ms.date: 11/04/2019
description: 'Обработка и возврат таких ошибок, как #ПУСТО!, из пользовательской функции'
title: Обработка и возврат ошибок из пользовательской функции (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: 5c62b7ccfbc1f0b450e6f36a0fd32f76fe099716
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217073"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a>Обработка и возврат ошибок из пользовательской функции (предварительная версия)

> [!NOTE]
> Возможности, описанные в этой статье, в настоящее время доступны в предварительной версии и могут изменяться. В настоящее время их нельзя использовать в рабочих средах. Вам нужно быть [участником предварительной оценки Office](https://insider.office.com/join), чтобы ознакомиться с предварительными возможностями.  Хороший способ ознакомиться с такими возможностями — использование подписки на Office 365. Если у вас еще нет подписки на Office 365, вы можете оформить бесплатную возобновляемую подписку на Office 365 на 90 дней, присоединившись к [программе для разработчиков Office 365](https://developer.microsoft.com/office/dev-program).

Если при выполнении пользовательской функции возникает ошибка, потребуется возвратить сообщение об ошибке, чтобы уведомить пользователя. Если у вас есть конкретные требования к параметрам, например применение только положительных чисел, нужно протестировать параметры и вернуть ошибку, если они неверны. Можно также использовать блок `try`-`catch`, чтобы отслеживать любые ошибки, возникающие при выполнении пользовательской функции.

## <a name="detect-and-throw-an-error"></a>Обнаружение и возвращение ошибки

Рассмотрим случай, в котором нужно убедиться в правильном формате параметра почтового индекса для пользовательской функции. В следующей пользовательской функции используется регулярное выражение для проверки почтового индекса. Если он правильный, будет подставлен город (в другой функции) и вернется значение. В противном случае в ячейке возвращается ошибка `#VALUE!`.

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

Объект `CustomFunctions.Error` используется для возвращения ошибки в ячейку. При создании объекта укажите, какую ошибку нужно использовать, применив одно из следующих значений перечисления `ErrorCode`.


|Значение перечисления ErrorCode  |Значение ячейки Excel  |Смысл  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | В формуле используется значение неправильного типа. |
|`notAvailable`   | `#N/A`    | Функция или служба недоступна. |
|`divisionByZero` | `#DIV/0`  | Обратите внимание, что JavaScript позволяет делить на нуль, поэтому при создании обработчика ошибок нужно внимательно определить это условие. |
|`invalidNumber`  | `#NUM!`   | Обнаружена проблема с числом, используемым в формуле |
|`nullReference`  | `#NULL!`  | Диапазоны формулы не пересекаются. |

В следующем примере кода показано, как создать и вернуть ошибку для неверного числа (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

При возврате ошибки `#VALUE!` также можно включить настраиваемое сообщение, отображаемое во всплывающем окне, когда пользователь наводит на ячейку указатель мыши. В следующем примере показано, как вернуть настраиваемое сообщение об ошибке.

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, “The parameter can only contain lowercase characters.”);
throw error;
```

## <a name="use-try-catch-blocks"></a>Использование блоков try-catch

Как правило, для отслеживания любых возможных ошибок следует использовать блоки `try`-`catch` в пользовательской функции. Если в коде не обрабатываются исключения, они будут возвращаться в Excel. По умолчанию Excel возвращает `#VALUE!` для необработанного исключения.

В следующем примере кода пользовательская функция создает запрос fetch в службу REST. Возможно, что вызов завершится сбоем (например, если служба REST возвращает ошибку или не работает сеть). В этом случае пользовательская функция возвращает `#N/A`, чтобы указать на сбой веб-вызова.


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

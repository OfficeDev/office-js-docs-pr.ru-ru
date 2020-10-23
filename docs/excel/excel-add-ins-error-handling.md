---
title: Обработка ошибок с помощью API JavaScript для Excel
description: Изучите логику обработки ошибок API JavaScript для Excel, чтобы учитывать ошибки времени выполнения.
ms.date: 10/22/2020
localization_priority: Normal
ms.openlocfilehash: a3b1bbfa7daba1b856bce35aa075d5b625bd9769
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740821"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Обработка ошибок с помощью API JavaScript для Excel

При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Это очень важно из-за асинхронного характера API.

> [!NOTE]
> Дополнительные сведения о `sync()` методе и асинхронной природе API JavaScript для Excel можно найти [в статье объектная модель JavaScript для Excel в](excel-add-ins-core-concepts.md)надстройках Office.

## <a name="best-practices"></a>Рекомендации

В примерах кода в этой документации вы заметите, что каждый вызов `Excel.run` сопровождается оператором `catch`, что позволяет перехватывать все ошибки, возникающие в `Excel.run`. Мы рекомендуем использовать этот шаблон, когда вы будете создавать надстройки с использованием API JavaScript для Excel.

```js
Excel.run(function (context) {
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);
```

## <a name="api-errors"></a>Ошибки API

Если не удается выполнить запрос API JavaScript для Excel, API возвращает объект error, содержащий следующие свойства:

- **code**.  Свойство `code` сообщения об ошибке содержит строку, входящую в список `OfficeExtension.ErrorCodes` или `Excel.ErrorCodes`. Например, код ошибки InvalidReference указывает, что ссылка недопустима для указанной операции. Коды ошибок не локализованы.

- **message.** Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке. Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.

- **debugInfo.** Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки.

> [!NOTE]
> Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере. Конечные пользователи не будут видеть эти сообщения об ошибках в области задач надстройки или в любом месте приложения Office.

## <a name="error-messages"></a>Сообщения об ошибках

В таблице ниже перечислены ошибки, которые может возвращать API.

|Код ошибки | Сообщение об ошибке |
|:----------|:--------------|
|`AccessDenied` |Вы не можете выполнить запрашиваемую операцию.|
|`ActivityLimitReached`|Достигнут предел действий.|
|`ApiNotAvailable`|Запрашиваемый интерфейс API недоступен.|
|`ApiNotFound`|Не удалось найти API, который вы пытаетесь использовать. Она может быть доступна в более новой версии Excel. Дополнительные сведения см. в статье набор обязательных элементов [API JavaScript для Excel](../reference/requirement-sets/excel-api-requirement-sets.md) .|
|`BadPassword`|Введен недопустимый пароль.|
|`Conflict`|Запрос не удалось обработать из-за конфликта.|
|`ContentLengthRequired`|`Content-length`Отсутствует заголовок HTTP.|
|`GeneralException`|При обработке запроса возникла внутренняя ошибка.|
|`InsertDeleteConflict`|Операция вставки или удаления привела к конфликту.|
|`InvalidArgument` |Аргумент недопустим, отсутствует или имеет неправильный формат.|
|`InvalidBinding`  |Эта привязка объектов недопустима из-за предыдущих обновлений.|
|`InvalidOperation`|Выполняемая операция недопустима для этого объекта.|
|`InvalidReference`|Эта ссылка недопустима для текущей операции.|
|`InvalidRequest`  |Не удается обработать запрос.|
|`InvalidSelection`|Выбранный фрагмент недопустим для этой операции.|
|`ItemAlreadyExists`|Создаваемый ресурс уже существует.|
|`ItemNotFound` |Запрашиваемый ресурс не существует.|
|`NonBlankCellOffSheet`|Запрос на вставку новых ячеек невозможно выполнить, так как он переносит непустые ячейки из конца листа. Эти непустые ячейки могут быть пустыми, но иметь пустые значения, часть форматирования или формулу. Удалите достаточное количество строк или столбцов, чтобы освободить место для вставки, а затем повторите попытку.|
|`NotImplemented`|Запрашиваемая функция не реализована.|
|`RangeExceedsLimit`|Число ячеек в диапазоне превысило максимальное поддерживаемое число. Дополнительные сведения см. в статье [пределы ресурсов и оптимизация производительности для надстроек Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .|
|`RequestAborted`|Запрос прерван во время выполнения.|
|`RequestPayloadSizeLimitExceeded`|Размер полезных данных запроса превысил допустимое значение. Дополнительные сведения см. в статье [пределы ресурсов и оптимизация производительности для надстроек Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) . <br><br>Эта ошибка возникает только в Excel в Интернете.|
|`ResponsePayloadSizeLimitExceeded`|Размер полезных данных ответа превысил допустимое значение. Дополнительные сведения см. в статье [пределы ресурсов и оптимизация производительности для надстроек Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) .  <br><br>Эта ошибка возникает только в Excel в Интернете.|
|`ServiceNotAvailable`|Служба недоступна.|
|`Unauthenticated` |Требуемые сведения о проверке подлинности отсутствуют или недопустимы.|
|`UnsupportedOperation`|Выполняемая операция не поддерживается.|
|`UnsupportedSheet`|Этот тип листа не поддерживает эту операцию, так как он является макросом или листом диаграммы.|

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Объект OfficeExtension.Error (API JavaScript для Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)

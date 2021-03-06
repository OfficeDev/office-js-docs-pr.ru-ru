---
title: Обработка ошибок с Excel API JavaScript
description: Узнайте о Excel логике обработки ошибок API JavaScript для учета ошибок во время работы.
ms.date: 01/15/2021
localization_priority: Normal
ms.openlocfilehash: 42ef52b5d20a2c2d1284f57c7b4026ff2c71ebdd
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349513"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Обработка ошибок с Excel API JavaScript

При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Это очень важно из-за асинхронного характера API.

> [!NOTE]
> Дополнительные сведения о методе и асинхронном характере API JavaScript Excel см. в Excel объектной модели JavaScript в Office `sync()` [надстройки.](excel-add-ins-core-concepts.md)

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

Если Excel API JavaScript не удается успешно выполнить, API возвращает объект ошибки, содержащий следующие свойства.

- **code**.  Свойство `code` сообщения об ошибке содержит строку, входящую в список `OfficeExtension.ErrorCodes` или `Excel.ErrorCodes`. Например, код ошибки InvalidReference указывает, что ссылка недопустима для указанной операции. Коды ошибок не локализованы.

- **message.** Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке. Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.

- **debugInfo.** Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки.

> [!NOTE]
> Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере. Конечные пользователи не будут видеть эти сообщения об ошибке в области задач надстройки или в любом Office приложении.

## <a name="error-messages"></a>Сообщения об ошибках

В таблице ниже перечислены ошибки, которые может возвращать API.

|Код ошибки | Сообщение об ошибке |
|:----------|:--------------|
|`AccessDenied` |Вы не можете выполнить запрашиваемую операцию.|
|`ActivityLimitReached`|Достигнут предел действий.|
|`ApiNotAvailable`|Запрашиваемый интерфейс API недоступен.|
|`ApiNotFound`|API, который вы пытаетесь использовать, не удалось найти. Он может быть доступен в более новой версии Excel. Дополнительные [сведения см. в Excel API JavaScript.](../reference/requirement-sets/excel-api-requirement-sets.md)|
|`BadPassword`|Предоставленный пароль является неправильным.|
|`Conflict`|Запрос не удалось обработать из-за конфликта.|
|`ContentLengthRequired`|Отсутствует `Content-length` заглавная головка HTTP.|
|`GeneralException`|При обработке запроса возникла внутренняя ошибка.|
|`InactiveWorkbook`|Операция не удалась, так как несколько книг открыты, а вызванная этим API книга потеряла фокус.|
|`InsertDeleteConflict`|Операция вставки или удаления привела к конфликту.|
|`InvalidArgument` |Аргумент недопустим, отсутствует или имеет неправильный формат.|
|`InvalidBinding`  |Эта привязка объектов недопустима из-за предыдущих обновлений.|
|`InvalidOperation`|Выполняемая операция недопустима для этого объекта.|
|`InvalidReference`|Эта ссылка недопустима для текущей операции.|
|`InvalidRequest`  |Не удается обработать запрос.|
|`InvalidSelection`|Выбранный фрагмент недопустим для этой операции.|
|`ItemAlreadyExists`|Создаваемый ресурс уже существует.|
|`ItemNotFound` |Запрашиваемый ресурс не существует.|
|`NonBlankCellOffSheet`|Microsoft Excel не может вставить новые ячейки, так как это отодвигает непустые ячейки с конца таблицы. Эти непустые ячейки могут казаться пустыми, но имеют пустые значения, некоторое форматирование или формулу. Удалите достаточно строк или столбцов, чтобы сделать место для того, что вы хотите вставить, а затем попробуйте еще раз.|
|`NotImplemented`|Запрашиваемая функция не реализована.|
|`RangeExceedsLimit`|Количество ячейки в диапазоне превысило максимально поддерживаемый номер. Дополнительные сведения см. в статье Ограничения ресурсов и [оптимизация производительности для Office надстройки.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)|
|`RequestAborted`|Запрос прерван во время выполнения.|
|`RequestPayloadSizeLimitExceeded`|Размер полезной нагрузки запроса превысил предел. Дополнительные сведения см. в статье Ограничения ресурсов и [оптимизация производительности для Office надстройки.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) <br><br>Эта ошибка возникает только в Excel в Интернете.|
|`ResponsePayloadSizeLimitExceeded`|Размер полезной нагрузки отклика превысил предел. Дополнительные сведения см. в статье Ограничения ресурсов и [оптимизация производительности для Office надстройки.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)  <br><br>Эта ошибка возникает только в Excel в Интернете.|
|`ServiceNotAvailable`|Служба недоступна.|
|`Unauthenticated` |Требуемые сведения о проверке подлинности отсутствуют или недопустимы.|
|`UnsupportedOperation`|Выполняемая операция не поддерживается.|
|`UnsupportedSheet`|Этот тип листа не поддерживает эту операцию, так как это лист Макроса или Диаграммы.|

> [!NOTE]
> В предыдущей таблице перечислены сообщения об ошибках, с которыми вы можете столкнуться при использовании Excel API JavaScript. Если вы работаете с общим API вместо приложения Excel API JavaScript, см. Office общие коды ошибок [API,](../reference/javascript-api-for-office-error-codes.md) чтобы узнать о соответствующих сообщениях об ошибках.

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Объект OfficeExtension.Error (API JavaScript для Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Коды ошибок общего API для Office](../reference/javascript-api-for-office-error-codes.md)

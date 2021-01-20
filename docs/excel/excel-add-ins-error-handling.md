---
title: Обработка ошибок с помощью API JavaScript для Excel
description: Узнайте о логике обработки ошибок API JavaScript для Excel, чтобы учесть ошибки во время работы.
ms.date: 01/15/2021
localization_priority: Normal
ms.openlocfilehash: 00aa1ae1c8ed39b21146d86090df912a8804c8b3
ms.sourcegitcommit: 4fc5829d66cdd52f110d9a59dd7317b520807cbe
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/20/2021
ms.locfileid: "49908908"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Обработка ошибок с помощью API JavaScript для Excel

При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Это очень важно из-за асинхронного характера API.

> [!NOTE]
> Дополнительные сведения о методе и асинхронном характере API JavaScript для Excel см. в объектной модели JavaScript для Excel в `sync()` [надстройки Office.](excel-add-ins-core-concepts.md)

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
> Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере. Конечные пользователи не будут видеть эти сообщения об ошибках в области задач надстройки или где-либо в приложении Office.

## <a name="error-messages"></a>Сообщения об ошибках

В таблице ниже перечислены ошибки, которые может возвращать API.

|Код ошибки | Сообщение об ошибке |
|:----------|:--------------|
|`AccessDenied` |Вы не можете выполнить запрашиваемую операцию.|
|`ActivityLimitReached`|Достигнут предел действий.|
|`ApiNotAvailable`|Запрашиваемый интерфейс API недоступен.|
|`ApiNotFound`|Не удалось найти API, который вы пытаетесь использовать. Она может быть доступна в более новой версии Excel. Дополнительные сведения см. в статье наборов требований [API JavaScript](../reference/requirement-sets/excel-api-requirement-sets.md) для Excel.|
|`BadPassword`|Пароль, который вы предоставили, неверен.|
|`Conflict`|Запрос не удалось обработать из-за конфликта.|
|`ContentLengthRequired`|Отсутствует `Content-length` заголок HTTP.|
|`GeneralException`|При обработке запроса возникла внутренняя ошибка.|
|`InactiveWorkbook`|Операция не удалась из-за того, что открыто несколько книг, и книга, которая была вызвана этим API, теряет фокус.|
|`InsertDeleteConflict`|Операция вставки или удаления привела к конфликту.|
|`InvalidArgument` |Аргумент недопустим, отсутствует или имеет неправильный формат.|
|`InvalidBinding`  |Эта привязка объектов недопустима из-за предыдущих обновлений.|
|`InvalidOperation`|Выполняемая операция недопустима для этого объекта.|
|`InvalidReference`|Эта ссылка недопустима для текущей операции.|
|`InvalidRequest`  |Не удается обработать запрос.|
|`InvalidSelection`|Выбранный фрагмент недопустим для этой операции.|
|`ItemAlreadyExists`|Создаваемый ресурс уже существует.|
|`ItemNotFound` |Запрашиваемый ресурс не существует.|
|`NonBlankCellOffSheet`|Microsoft Excel не может вставлять новые ячейки, так как при этом непустые ячейки будут отставляться с конца листа. Эти непустые ячейки могут отображаться пустыми, но имеют пустые значения, некоторые форматирование или формулу. Удалите достаточно строк или столбцов, чтобы упустить место для вставки, а затем попробуйте еще раз.|
|`NotImplemented`|Запрашиваемая функция не реализована.|
|`RangeExceedsLimit`|Число ячеок в диапазоне превысило максимальное поддерживаемые числа. Дополнительные [сведения см.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) в статье об ограничениях ресурсов и оптимизации производительности надстройки Office.|
|`RequestAborted`|Запрос прерван во время выполнения.|
|`RequestPayloadSizeLimitExceeded`|Размер полезной нагрузки запроса превысил ограничение. Дополнительные [сведения см.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) в статье об ограничениях ресурсов и оптимизации производительности надстройки Office. <br><br>Эта ошибка возникает только в Excel в Интернете.|
|`ResponsePayloadSizeLimitExceeded`|Размер полезной нагрузки ответа превысил ограничение. Дополнительные [сведения см.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) в статье об ограничениях ресурсов и оптимизации производительности надстройки Office.  <br><br>Эта ошибка возникает только в Excel в Интернете.|
|`ServiceNotAvailable`|Служба недоступна.|
|`Unauthenticated` |Требуемые сведения о проверке подлинности отсутствуют или недопустимы.|
|`UnsupportedOperation`|Выполняемая операция не поддерживается.|
|`UnsupportedSheet`|Этот тип листа не поддерживает эту операцию, так как он является листом макроса или диаграммы.|

> [!NOTE]
> В предыдущей таблице перечислены сообщения об ошибках, которые могут возникнуть при использовании API JavaScript для Excel. Если вы работаете с общим API, а не С API JavaScript для Конкретных приложений, см. коды ошибок общего [API Office,](../reference/javascript-api-for-office-error-codes.md) чтобы узнать о соответствующих сообщениях об ошибках.

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Объект OfficeExtension.Error (API JavaScript для Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Коды ошибок общего API для Office](../reference/javascript-api-for-office-error-codes.md)

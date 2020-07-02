---
title: Обработка ошибок
description: Изучите логику обработки ошибок API JavaScript для Excel, чтобы учитывать ошибки времени выполнения.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 8d410ae7eea315e14383b5aa08373ede3768cace
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006446"
---
# <a name="error-handling"></a>Обработка ошибок

When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.

> [!NOTE]
> Для получения дополнительных сведений о `sync()` методе и асинхронной природе API JavaScript для Excel ознакомьтесь [с основными концепциями программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md).

## <a name="best-practices"></a>Рекомендации

Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.

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

- **code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.

- **message.** Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке. Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.

- **debugInfo.** Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки.

> [!NOTE]
> Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере. Эти сообщения об ошибках не будут отображаться для пользователей в области задач надстройки или в другом месте ведущего приложения.

## <a name="error-messages"></a>Сообщения об ошибках

В таблице ниже перечислены ошибки, которые может возвращать API.

|error.code | error.message |
|:----------|:--------------|
|`AccessDenied` |Вы не можете выполнить запрашиваемую операцию.|
|`ActivityLimitReached`|Достигнут предел действий.|
|`ApiNotAvailable`|Запрашиваемый интерфейс API недоступен.|
|`Conflict`|Запрос не удалось обработать из-за конфликта.|
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
|`NotImplemented`  |Запрашиваемая функция не реализована.|
|`RequestAborted`|Запрос прерван во время выполнения.|
|`ServiceNotAvailable`|Служба недоступна.|
|`Unauthenticated` |Требуемые сведения о проверке подлинности отсутствуют или недопустимы.|
|`UnsupportedOperation`|Выполняемая операция не поддерживается.|
|`UnsupportedSheet`|Этот тип листа не поддерживает эту операцию, так как он является макросом или листом диаграммы.|

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Объект OfficeExtension.Error (API JavaScript для Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview)

---
title: Обработка ошибок
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: caba29f7d6949cc6d9df1498ac0a3d4f5de6c4ee
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579816"
---
# <a name="error-handling"></a>Обработка ошибок

При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок выполнения. Это очень важно из-за асинхронного характера API.

> [!NOTE]
> Дополнительные сведения о методе **sync()** и асинхронном характере API JavaScript для Excel см. в статье [Фундаментальные понятия программирования с использованием интерфейса API JavaScript для Excel](excel-add-ins-core-concepts.md).

## <a name="best-practices"></a>Рекомендации

В примерах кода в этой документации вы заметите, что каждый вызов `Excel.run` сопровождается оператором `catch`, что позволяет перехватывать все ошибки, возникающие в `Excel.run`. Мы рекомендуем использовать этот шаблон при создании надстройки с использованием API JavaScript для Excel.

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

Если не удается выполнить запрос API JavaScript для Excel, API возвращает объект error, содержащий приведенные ниже свойства. 

- **code**. Свойство `code` сообщения об ошибке содержит строку, входящую в список `OfficeExtension.ErrorCodes` или `Excel.ErrorCodes`. Например, код ошибки InvalidReference указывает, что ссылка недопустима для указанной операции. Коды ошибок не локализованы. 

- **message**. Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке. Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.

- **debugInfo**. Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые можно использовать, чтобы понять первопричину ошибки. 

> [!NOTE]
> Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будут отображаться только на сервере. Эти сообщения об ошибках не будут отображаться для пользователей в области задач надстройки или в другом месте ведущего приложения.

## <a name="error-messages"></a>Сообщения об ошибках

В таблице ниже перечислены ошибки, которые может возвращать API.

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |Аргумент недопустим, отсутствует или имеет неправильный формат.|
|InvalidRequest  |Не удается обработать запрос.|
|InvalidReference|Эта ссылка недопустима для текущей операции.|
|InvalidBinding  |Эта привязка объектов недопустима из-за предыдущих обновлений.|
|InvalidSelection|Выбранный фрагмент недопустим для этой операции.|
|Unauthenticated |Требуемые сведения о проверке подлинности отсутствуют или недопустимы.|
|AccessDenied |Вы не можете выполнить запрашиваемую операцию.|
|ItemNotFound |Запрашиваемый ресурс не существует.|
|ActivityLimitReached|Достигнут предел действий.|
|GeneralException|При обработке запроса возникла внутренняя ошибка.|
|NotImplemented  |Запрашиваемая функция не реализована.|
|ServiceNotAvailable|Служба недоступна.|
|Conflict|Запрос не удалось обработать из-за конфликта.|
|ItemAlreadyExists|Создаваемый ресурс уже существует.|
|UnsupportedOperation|Выполняемая операция не поддерживается.|
|RequestAborted|Запрос прерван во время выполнения.|
|ApiNotAvailable|Запрашиваемый интерфейс API недоступен.|
|InsertDeleteConflict|Операция вставки или удаления привела к конфликту.|
|InvalidOperation|Выполняемая операция недопустима для этого объекта.|

## <a name="see-also"></a>См. также

- [Фундаментальные понятия программирования с использованием интерфейса API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Объект OfficeExtension.Error (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)

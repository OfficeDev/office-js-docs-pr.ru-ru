---
title: Обработка ошибок
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 59619ff4ccba1985f875d13a29ab691c1617d21b
ms.sourcegitcommit: 7ecc1dc24bf7488b53117d7a83ad60e952a6f7aa
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/23/2018
ms.locfileid: "19437208"
---
# <a name="error-handling"></a>Обработка ошибок

При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Это очень важно из-за асинхронного характера API.

> [!NOTE]
> Дополнительные сведения о методе **sync()** и асинхронном характере API JavaScript для Excel см. в статье [Основные понятия API JavaScript для Excel](excel-add-ins-core-concepts.md).

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

- **message**. Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке. Сообщение об ошибке не предназначено для пользователей; код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.

- **debugInfo**. Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки. 

> [!NOTE]
> Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере. Эти сообщения об ошибках не будут отображаться для пользователей в области задач надстройки или в другом месте ведущего приложения.

## <a name="see-also"></a>См. также

- [Основные понятия API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Объект OfficeExtension.Error (API JavaScript для Excel)](https://dev.office.com/reference/add-ins/excel/error)

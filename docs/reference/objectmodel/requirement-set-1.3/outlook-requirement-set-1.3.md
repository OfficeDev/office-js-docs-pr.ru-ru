---
title: Набор обязательных элементов API для надстройки Outlook 1.3
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 375fc5d7cce8592b8e4a270713c1f611129cc7d0
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165428"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Набор обязательных элементов API для надстройки Outlook 1.3

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).

## <a name="whats-new-in-13"></a>Новые возможности в версии 1.3

Набор обязательных элементов 1.3 включает все возможности [набора обязательных элементов версии 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). В нем добавлены перечисленные ниже возможности.

- Добавлена поддержка [команд надстроек](../../../outlook/add-in-commands-for-outlook.md).
- Добавлена возможность сохранять и закрывать создаваемый элемент.
- Расширенный объект [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3) , позволяющий надстройкам получать или задавать текст целиком.
- Добавлены методы для преобразования идентификаторов из формата EWS в формат REST и наоборот.
- Появилась возможность добавлять сообщения уведомления на информационную панель элементов.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-). Возвращает текущий текст в указанном формате.
- Добавлен метод [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#setasync-data--options--callback-). Заменяет весь текст указанным текстом.
- Добавлен объект [Event](/javascript/api/office/office.addincommands.event). Передается как параметр в функции команд, не требующих пользовательского интерфейса, в надстройке Outlook. Используется для уведомления о завершении обработки.
- Добавлен метод [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods). Закрывает текущий создаваемый элемент.
- Добавлен метод [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods). Асинхронно сохраняет элемент.
- Добавлено свойство [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties). Получает сообщения уведомления для элемента.
- Добавлен метод [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods). Преобразует идентификатор элемента из формата REST в формат EWS.
- Добавлен метод [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods). Преобразует идентификатор элемента из формата EWS в формат REST.
- Добавлено свойство [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3). Указывает тип сообщения уведомления для встречи или сообщения.
- Добавлено свойство [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3). Указывает версию REST API, которая соответствует идентификатору элемента в формате REST.
- Добавлен объект [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3). Предоставляет методы для доступа к сообщениям уведомления в надстройке Outlook.
- Добавлен тип [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3). Возвращается методом `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)

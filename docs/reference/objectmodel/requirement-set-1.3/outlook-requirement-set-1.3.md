---
title: Набор обязательных элементов API для надстройки Outlook 1.3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 276096870b128896e987bcb303b4cccdb77e0e50
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871278"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Набор обязательных элементов API для надстройки Outlook 1.3

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). 

## <a name="whats-new-in-13"></a>Новые возможности в версии 1.3

Набор обязательных элементов 1.3 включает все возможности [набора обязательных элементов версии 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). В нем добавлены перечисленные ниже возможности.

- Добавлена поддержка [команд надстроек](/outlook/add-ins/add-in-commands-for-outlook).
- Добавлена возможность сохранять и закрывать создаваемый элемент.
- Расширенный объект [Body](/javascript/api/outlook_1_3/office.body) позволяет надстройкам получать или задавать текст целиком.
- Добавлены методы для преобразования идентификаторов из формата EWS в формат REST и наоборот.
- Появилась возможность добавлять сообщения уведомления на информационную панель элементов.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-). Возвращает текущий текст в указанном формате.
- Добавлен метод [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-). Заменяет весь текст указанным текстом.
- Добавлено свойство [Office.context.officeTheme](office.context.md#officetheme-object). Предоставляет доступ к цветам темы Office.
- Добавлен объект [Event](/javascript/api/office/office.addincommands.event). Передается как параметр в функции команд, не требующих пользовательского интерфейса, в надстройке Outlook. Используется для уведомления о завершении обработки.
- Добавлен метод [Office.context.mailbox.item.close](office.context.mailbox.item.md#close). Закрывает текущий создаваемый элемент.
- Добавлен метод [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback). Асинхронно сохраняет элемент.
- Добавлено свойство [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessages). Получает сообщения уведомления для элемента.
- Добавлен метод [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string). Преобразует идентификатор элемента из формата REST в формат EWS.
- Добавлен метод [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Преобразует идентификатор элемента из формата EWS в формат REST.
- Добавлено свойство [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype). Указывает тип сообщения уведомления для встречи или сообщения.
- Добавлено свойство [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion). Указывает версию REST API, которая соответствует идентификатору элемента в формате REST.
- Добавлен объект [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages). Предоставляет методы для доступа к сообщениям уведомления в надстройке Outlook.
- Добавлен тип [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails). Возвращается методом `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/outlook/add-ins/quick-start)

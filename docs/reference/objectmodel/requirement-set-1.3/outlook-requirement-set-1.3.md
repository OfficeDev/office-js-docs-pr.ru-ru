---
title: Набор обязательных элементов API для надстройки Outlook 1.3
description: Функции и API, которые были Outlook надстройки и Office API JavaScript в рамках API почтовых ящиков 1.3.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 85924d181ee494a8caa5e18a5bcf53c3f116ee3e
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671900"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Набор обязательных элементов API для надстройки Outlook 1.3

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).

## <a name="whats-new-in-13"></a>Новые возможности в версии 1.3

Набор требований 1.3 включает все функции набора [требований 1.2.](../requirement-set-1.2/outlook-requirement-set-1.2.md) В нем добавлены перечисленные ниже возможности.

- Добавлена поддержка [команд надстроек](../../../outlook/add-in-commands-for-outlook.md).
- Добавлена возможность сохранять и закрывать создаваемый элемент.
- Расширенный [объект Body,](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true) позволяющий надстройки получать или устанавливать все тело.
- Добавлены методы для преобразования идентификаторов из формата EWS в формат REST и наоборот.
- Появилась возможность добавлять сообщения уведомления на информационную панель элементов.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true#getAsync_coercionType__options__callback_). Возвращает текущий текст в указанном формате.
- Добавлен метод [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true#setAsync_data__options__callback_). Заменяет весь текст указанным текстом.
- Добавлен объект [Event](/javascript/api/office/office.addincommands.event). Передается как параметр в функции команд, не требующих пользовательского интерфейса, в надстройке Outlook. Используется для уведомления о завершении обработки.
- Добавлен метод [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods). Закрывает текущий создаваемый элемент.
- Добавлен метод [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods). Асинхронно сохраняет элемент.
- Добавлено свойство [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties). Получает сообщения уведомления для элемента.
- Добавлен метод [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods). Преобразует идентификатор элемента из формата REST в формат EWS.
- Добавлен метод [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods). Преобразует идентификатор элемента из формата EWS в формат REST.
- Добавлено свойство [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3&preserve-view=true). Указывает тип сообщения уведомления для встречи или сообщения.
- Добавлено свойство [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3&preserve-view=true). Указывает версию REST API, которая соответствует идентификатору элемента в формате REST.
- Добавлен объект [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3&preserve-view=true). Предоставляет методы для доступа к сообщениям уведомления в надстройке Outlook.
- Добавлен тип [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3&preserve-view=true). Возвращается методом `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)

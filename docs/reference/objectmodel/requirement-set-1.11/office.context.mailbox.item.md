---
title: Office.context.mailbox.item — набор требований 1.11
description: Outlook Требования К API почтовых ящиков устанавливают версию 1.11 объектной модели Item.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0f2b51efe510f60810140ccddc604a2d2ef85b51
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2021
ms.locfileid: "60682439"
---
# <a name="item-mailbox-requirement-set-111"></a>элемент (набор требований к почтовым ящикам 1.11)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` используется для доступа к выбранному в настоящее время сообщению, собранию или встрече. Тип элемента можно определить с помощью `itemType` свойства.

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)|С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)|Организатор встречи, участник встречи,<br>Композит сообщения или чтение сообщений|

> [!IMPORTANT]
> Android и iOS. Существуют ограничения, когда надстройки активируются и какие API доступны. Дополнительные сведения см. в статье [Добавление поддержки мобильных устройств в надстройку Outlook](../../../outlook/add-mobile-support.md#compose-mode-and-appointments).

## <a name="properties"></a>Свойства

| Свойство | Minimum<br>уровень разрешения | Сведения по режиму | Тип возвращаемых данных | Minimum<br>набор требований |
|---|---|---|---|:---:|
| attachments | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.11&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.11&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| СК. | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#bcc) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#body) | [Основной текст](/javascript/api/outlook/office.body?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#body) | [Основной текст](/javascript/api/outlook/office.body?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#body) | [Основной текст](/javascript/api/outlook/office.body?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#body) | [Основной текст](/javascript/api/outlook/office.body?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| копия; | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#cc) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.11&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#conversationId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#conversationId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#dateTimeCreated) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#dateTimeCreated) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#dateTimeModified) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#dateTimeModified) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#end) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#end)<br>(Запрос собрания) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#from) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.11&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#internetHeaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#internetMessageId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| расположение | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#location) | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#location)<br>(Запрос собрания) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#normalizedSubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#normalizedSubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.11&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.11&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.11&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.11&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#optionalAttendees) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#optionalAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.11&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#organizer) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.11&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#recurrence) | [Повторение](/javascript/api/outlook/office.recurrence?view=outlook-js-1.11&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#recurrence) | [Повторение](/javascript/api/outlook/office.recurrence?view=outlook-js-1.11&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#recurrence)<br>(Запрос собрания) | [Повторение](/javascript/api/outlook/office.recurrence?view=outlook-js-1.11&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#requiredAttendees) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#requiredAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.11&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| sessionData | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#sessionData) | [SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true) | [1.11](outlook-requirement-set-1.11.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#sessionData) | [SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true) | [1.11](outlook-requirement-set-1.11.md) |
| start | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#start) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#start)<br>(Запрос собрания) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#subject) | [Тема](/javascript/api/outlook/office.subject?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#subject) | [Тема](/javascript/api/outlook/office.subject?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| на | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#to) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.11&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Minimum<br>уровень разрешения | Сведения по режиму | Minimum<br>набор требований |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| disableClientSignatureAsync ([options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#disableClientSignatureAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#disableClientSignatureAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| displayReplyAllForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllFormAsync(formData, [options], [callback]) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| displayReplyForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyFormAsync(formData, [options], [callback]) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| getAllInternetHeadersAsync ([options], [callback]) | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getAllInternetHeadersAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync(attachmentId, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync ([options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getComposeTypeAsync ([options], callback) | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#getComposeTypeAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| getEntities() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getItemIdAsync ([options], callback) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync ([options], callback) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| isClientSignatureEnabledAsync ([options], callback) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#isClientSignatureEnabledAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#isClientSignatureEnabledAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.11&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.11&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.11&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>События

Вы можете подписаться и отписаться от следующих событий с помощью `addHandlerAsync` и `removeHandlerAsync` соответственно.

> [!IMPORTANT]
> События доступны только с реализацией области задач.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.11&preserve-view=true) | Описание | Minimum<br>набор требований |
|---|---|:---:|
|`AppointmentTimeChanged`| Изменилась дата или время выбранной встречи или серии. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| Вложение было добавлено или удалено из элемента. | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| Расположение выбранного назначения изменилось. | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`RecipientsChanged`| Список получателей выбранного элемента или расположения встречи изменен. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| Изменился шаблон повторяемости выбранной серии. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

## <a name="example"></a>Пример

В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```
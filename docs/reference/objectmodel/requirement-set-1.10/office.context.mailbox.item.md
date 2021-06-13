---
title: Office.context.mailbox.item — набор требований 1.10
description: Outlook Требования К API почтовых ящиков устанавливают версию 1.10 объектной модели Item.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f7de10591476de9cd83721656ef1d005549fe482
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893654"
---
# <a name="item-mailbox-requirement-set-110"></a>элемент (набор требований к почтовым ящикам 1.10)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` используется для доступа к выбранному в настоящее время сообщению, собранию или встрече. Тип элемента можно определить с помощью `itemType` свойства.

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)|С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)|Организатор встречи, участник встречи,<br>Композит сообщения или чтение сообщений|

## <a name="properties"></a>Свойства

| Свойство | Minimum<br>уровень разрешения | Сведения по режиму | Тип возвращаемых данных | Minimum<br>набор требований |
|---|---|---|---|:---:|
| вложения | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| СК. | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#bcc) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| копия; | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#cc) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#end)<br>(Запрос собрания) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#from) | [From](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| расположение; | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#location) | [Расположение](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#location)<br>(Запрос собрания) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#optionalattendees) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#optionalattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#recurrence)<br>(Запрос собрания) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#requiredattendees) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#requiredattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| начать | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#start)<br>(Запрос собрания) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#subject) | [Тема](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#subject) | [Тема](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| на | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#to) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Minimum<br>уровень разрешения | Сведения по режиму | Minimum<br>набор требований |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| disableClientSignatureAsync ([options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#disableclientsignatureasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#disableclientsignatureasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| displayReplyAllForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllFormAsync(formData, [options], [callback]) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayreplyallformasync-formdata--options--callback-) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayreplyallformasync-formdata--options--callback-) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| displayReplyForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyFormAsync(formData, [options], [callback]) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayreplyformasync-formdata--options--callback-) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayreplyformasync-formdata--options--callback-) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| getAllInternetHeadersAsync ([options], [callback]) | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getallinternetheadersasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync(attachmentId, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync ([options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getattachmentsasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getattachmentsasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getComposeTypeAsync ([options], callback) | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getcomposetypeasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| getEntities() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getItemIdAsync ([options], callback) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getitemidasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getitemidasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync ([options], callback) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| isClientSignatureEnabledAsync ([options], callback) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#isclientsignatureenabledasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#isclientsignatureenabledasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>События

Вы можете подписаться и отписаться от следующих событий с помощью `addHandlerAsync` и `removeHandlerAsync` соответственно.

> [!IMPORTANT]
> События доступны только с реализацией области задач.

| [Event](/javascript/api/office/office.eventtype) | Описание | Minimum<br>набор требований |
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

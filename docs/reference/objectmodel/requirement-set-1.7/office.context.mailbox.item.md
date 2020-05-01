---
title: Office. Context. Mailbox. Item — набор требований 1,7
description: Набор требований API почтового ящика Outlook 1,7 версия объектной модели элемента.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: ce5d86f2bf525169d26c4d6a7cf7a66e87b98ee9
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891301"
---
# <a name="item-mailbox-requirement-set-17"></a>Item (набор требований для почтового ящика 1,7)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`используется для доступа к текущему выбранному сообщению, приглашению на собрание или встрече. Тип элемента можно определить с помощью `itemType` свойства.

##### <a name="requirements"></a>Requirements

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)|С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)|Организатор встречи, участник встречи,<br>Создание сообщения или чтение сообщения|

## <a name="properties"></a>Свойства

| Свойство | Минимальные<br>уровень разрешения | Сведения по режиму | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|---|:---:|
| attachments | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| СК. | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#bcc) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#body) | [Основной текст](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#body) | [Основной текст](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#body) | [Основной текст](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#body) | [Основной текст](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| копия; | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#cc) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#cc) | Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#conversationid) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#conversationid) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#datetimecreated) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#datetimecreated) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#datetimemodified) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#datetimemodified) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#end) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#end)<br>(Приглашение на собрание) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadWriteItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#from) | [From](/javascript/api/outlook/office.from) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#internetmessageid) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemclass) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemclass) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Идентификатор | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemid) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemid) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#location) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#location)<br>(Приглашение на собрание) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#normalizedsubject) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#normalizedsubject) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#optionalattendees) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#optionalattendees) | Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#recurrence) | [Повторения](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#recurrence) | [Повторения](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#recurrence)<br>(Приглашение на собрание) | [Повторения](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#requiredattendees) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#requiredattendees) | Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| отправитель; | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#seriesid) | Строка | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#seriesid) | Строка | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#seriesid) | Строка | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#seriesid) | Строка | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#start) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#start)<br>(Приглашение на собрание) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#subject) | [Тема](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#subject) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#subject) | [Тема](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#subject) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| to | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#to) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#to) | Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Минимальные<br>уровень разрешения | Сведения по режиму | Минимальные<br>набор требований |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Entities () | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restricted | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches () | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType, [параметры], обратный вызов) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| Жетселектедентитиес () | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| Жетселектедрежексматчес () | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>События

Вы можете подписаться на следующие события и отписаться `addHandlerAsync` на `removeHandlerAsync` них, используя и соответственно.

| Событие | Описание | Минимальные<br>набор требований |
|---|---|:---:|
|`AppointmentTimeChanged`| Дата или время выбранной встречи или ряда изменились. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecipientsChanged`| Список получателей выбранного элемента или места встречи изменился. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| Шаблон повторения выбранного ряда изменился. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

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

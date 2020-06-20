---
title: Office. Context. Mailbox. Item — Предварительная версия набора требований
description: Предварительная версия набора обязательных элементов API почтового ящика Outlook для объектной модели элемента.
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: c9a384bcbac05b5a6f30b73fadcc5db19642ccb6
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778685"
---
# <a name="item-mailbox-preview-requirement-set"></a>элемент (набор требований Preview для почтового ящика)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`используется для доступа к текущему выбранному сообщению, приглашению на собрание или встрече. Тип элемента можно определить с помощью `itemType` Свойства.

##### <a name="requirements"></a>Requirements

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)|С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)|Организатор встречи, участник встречи,<br>Создание сообщения или чтение сообщения|

## <a name="properties"></a>Свойства

| Свойство | Минимальные<br>уровень разрешения | Сведения по режиму | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|---|:---:|
| attachments | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| СК. | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#bcc) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#categories) | [Категории](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#categories) | [Категории](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#categories) | [Категории](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#categories) | [Категории](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| копия; | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#cc) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#cc) | Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimecreated) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimecreated) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimemodified) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimemodified) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#end) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#end)<br>(Приглашение на собрание) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| енханцедлокатион | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#from) | [From](/javascript/api/outlook/office.from) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Internetheaders: | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| исаллдайевент | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#isalldayevent) | [IsAllDayEvent](/javascript/api/outlook/office.isalldayevent) | [Предварительная версия](outlook-requirement-set-preview.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#isalldayevent) | Boolean | [Предварительная версия](outlook-requirement-set-preview.md) |
| itemClass | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Идентификатор | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#location)<br>(Приглашение на собрание) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#optionalattendees) | Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#recurrence) | [Повторения](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#recurrence) | [Повторения](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#recurrence)<br>(Приглашение на собрание) | [Повторения](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#requiredattendees) | Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| отправитель; | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sensitivity | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#sensitivity) | [Sensitivity](/javascript/api/outlook/office.sensitivity) | [Предварительная версия](outlook-requirement-set-preview.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#sensitivity) | [MailboxEnums. Аппоинтментсенситивититипе](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype) | [Предварительная версия](outlook-requirement-set-preview.md) |
| seriesId | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| начать | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#start) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#start)<br>(Приглашение на собрание) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) | [Тема](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#subject) | [Тема](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| на | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#to) | [Получатели](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#to) | Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Минимальные<br>уровень разрешения | Сведения по режиму | Минимальные<br>набор требований |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| Дисаблеклиентсигнатуреасинк ([параметры], [обратный вызов]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#disableclientsignatureasync-options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#disableclientsignatureasync-options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| displayReplyAllForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Дисплайрепляллформасинк (Формдата, [параметры], [обратный вызов]) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyallformasync-formdata--options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyallformasync-formdata--options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| displayReplyForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Дисплайреплиформасинк (Формдата, [параметры], [обратный вызов]) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyformasync-formdata--options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyformasync-formdata--options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| Жеталлинтернесеадерсасинк ([параметры], [обратный вызов]) | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getallinternetheadersasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| Жетаттачментконтентасинк (attachmentId, [параметры], [обратный вызов]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| Жетаттачментсасинк ([параметры], [обратный вызов]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| Жеткомпосетипеасинк ([параметры], обратный вызов) | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| Entities () | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restricted | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getInitializationContextAsync ([параметры], [обратный вызов]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Предварительная версия](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Предварительная версия](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Предварительная версия](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Предварительная версия](../preview-requirement-set/outlook-requirement-set-preview.md) |
| Жетитемидасинк ([параметры], обратный вызов) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches () | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType, [параметры], обратный вызов) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| Жетселектедентитиес () | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| Жетселектедрежексматчес () | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| Жетшаредпропертиесасинк ([параметры], обратный вызов) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| Исклиентсигнатуринабледасинк ([параметры], обратный вызов) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#isclientsignatureenabledasync-options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#isclientsignatureenabledasync-options--callback-) | [Предварительная версия](outlook-requirement-set-preview.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>События

Вы можете подписаться на следующие события и отписаться на них, используя `addHandlerAsync` и `removeHandlerAsync` соответственно.

| Событие | Описание | Минимальные<br>набор требований |
|---|---|:---:|
|`AppointmentTimeChanged`| Дата или время выбранной встречи или ряда изменились. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| Вложение было добавлено или удалено из элемента. | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| Расположение выбранной встречи изменилось. | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
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

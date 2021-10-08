---
title: Office.context.mailbox.item — набор требований 1.1
description: Outlook Требования К API почтовых ящиков устанавливают версию 1.1 объектной модели Item.
ms.date: 07/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 52c778fab1687351d6cc1a55b2f8b2d6af24c527
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237204"
---
# <a name="item-mailbox-requirement-set-11"></a>элемент (набор требований к почтовым ящикам 1.1)

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
| attachments | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| СК. | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#bcc) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#body) | [Основной текст](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#body) | [Основной текст](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#body) | [Основной текст](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#body) | [Основной текст](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| копия; | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#cc) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#conversationId) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#conversationId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#dateTimeCreated) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#dateTimeCreated) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#dateTimeModified) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#dateTimeModified) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#end) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#end)<br>(Запрос собрания) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#internetMessageId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#itemClass) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#itemClass) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#itemId) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#location) | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#location) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#location)<br>(Запрос собрания) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#normalizedSubject) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#normalizedSubject) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| optionalAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#optionalAttendees) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#optionalAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#requiredAttendees) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#requiredAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| start | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#start) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#start)<br>(Запрос собрания) | Дата | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#subject) | [Тема](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#subject) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#subject) | [Тема](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#subject) | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| на | ReadItem | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#to) | [Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Minimum<br>уровень разрешения | Сведения по режиму | Minimum<br>набор требований |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches() | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (имя) | ReadItem | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Участник встречи](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.1&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Чтение сообщения](/javascript/api/outlook/office.messageread?view=outlook-js-1.1&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Организатор встречи](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.1&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Создание сообщения](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.1&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

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

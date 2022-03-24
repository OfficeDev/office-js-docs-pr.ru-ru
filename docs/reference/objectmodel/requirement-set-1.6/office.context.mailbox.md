---
title: Office.context.mailbox — набор требований 1.6
description: Outlook API почтовых ящиков установлена версия 1.6 объектной модели почтовых ящиков.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6254ec331b007f1535d2999a03be1e426d3d8e51
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745262"
---
# <a name="mailbox-requirement-set-16"></a>почтовый ящик (набор требований 1.6)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)| С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

## <a name="properties"></a>Свойства

| Свойство | Minimum<br>уровень разрешения | Режимы | Тип возвращаемых данных | Minimum<br>набор требований |
|---|---|---|---|:---:|
| [диагностика](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-diagnostics-member) | ReadItem | Создание<br>Чтение | [Диагностика](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-ewsurl-member) | ReadItem | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [элемента](office.context.mailbox.item.md) | Restricted | Создание<br>Чтение | [Элемент](/javascript/api/outlook/office.item?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-resturl-member) | ReadItem | Создание<br>Чтение | Строка | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-userprofile-member) | ReadItem | Создание<br>Чтение | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Minimum<br>уровень разрешения | Режимы | Minimum<br>набор требований |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-converttoewsid-member(1)) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-converttolocalclienttime-member(1)) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-converttorestid-member(1)) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (вход)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-converttoutcclienttime-member(1)) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-displayappointmentform-member(1)) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-displaymessageform-member(1)) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-displaynewappointmentform-member(1)) | ReadItem | Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm (параметры)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-displaynewmessageform-member(1)) | ReadItem | Чтение | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(1)) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(2)) | ReadItem | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-getuseridentitytokenasync-member(1)) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-makeewsrequestasync-member(1)) | ReadWriteMailbox | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>События

Вы можете подписаться и отписаться от следующих событий с помощью [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) и [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) соответственно.

> [!IMPORTANT]
> События доступны только с реализацией области задач.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.6&preserve-view=true) | Описание | Minimum<br>набор требований |
|---|---|:---:|
|`ItemChanged`| Другой элемент Outlook для просмотра при закреплении области задач. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

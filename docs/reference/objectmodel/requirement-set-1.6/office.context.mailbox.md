---
title: Office.context.mailbox — набор требований 1.6
description: Outlook Требования К API почтовых ящиков устанавливают версию 1.6 объектной модели почтовых ящиков.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 38fde5ae321676659e8f9c43678a04416be10351
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893570"
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
| [диагностика](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#diagnostics) | ReadItem | Создание<br>Чтение | [Диагностика](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#ewsurl) | ReadItem | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [элемента](office.context.mailbox.item.md) | Restricted | Создание<br>Чтение | [Элемент](/javascript/api/outlook/office.item?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#resturl) | ReadItem | Создание<br>Чтение | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#userprofile) | ReadItem | Создание<br>Чтение | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Minimum<br>уровень разрешения | Режимы | Minimum<br>набор требований |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#converttoewsid-itemid--restversion-) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#converttorestid-itemid--restversion-) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (вход)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#converttoutcclienttime-input-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displayappointmentform-itemid-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displaymessageform-itemid-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm (параметры)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displaynewmessageform-parameters-) | ReadItem | Чтение | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>События

Вы можете подписаться и отписаться от следующих событий с помощью [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) и [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#removehandlerasync-eventtype--options--callback-) соответственно.

> [!IMPORTANT]
> События доступны только с реализацией области задач.

| [Event](/javascript/api/office/office.eventtype) | Описание | Minimum<br>набор требований |
|---|---|:---:|
|`ItemChanged`| Другой элемент Outlook для просмотра при закреплении области задач. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

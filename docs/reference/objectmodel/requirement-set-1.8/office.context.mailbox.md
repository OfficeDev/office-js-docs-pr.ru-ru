---
title: Office. Context. Mailbox — набор обязательных элементов 1,8
description: Набор требований API почтового ящика Outlook 1,8 версия объектной модели почтового ящика.
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: e22ad060b5cea75fd1686687d43be813e5c9ca71
ms.sourcegitcommit: fd110305c2be8660ab8a47c1da3e3969bd1ede86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/23/2020
ms.locfileid: "48214563"
---
# <a name="mailbox-requirement-set-18"></a>Mailbox (набор требований 1,8)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)| С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

## <a name="properties"></a>Свойства

| Свойство | Minimum<br>уровень разрешения | Способов | Тип возвращаемых данных | Minimum<br>набор требований |
|---|---|---|---|:---:|
| [диагностики](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#diagnostics) | ReadItem | Создание<br>Чтение | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#ewsurl) | ReadItem | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [элемента](office.context.mailbox.item.md) | Restricted | Создание<br>Чтение | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [мастеркатегориес](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#mastercategories) | ReadWriteMailbox | Создание<br>Чтение | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8&preserve-view=true) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#resturl) | ReadItem | Создание<br>Чтение | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#userprofile) | ReadItem | Создание<br>Чтение | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Minimum<br>уровень разрешения | Способов | Minimum<br>набор требований |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Создание<br>Чтение | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [Конверттоевсид (itemId, Рестверсион)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#converttoewsid-itemid--restversion-) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, Рестверсион)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#converttorestid-itemid--restversion-) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (вход)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#converttoutcclienttime-input-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#displayappointmentform-itemid-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#displaymessageform-itemid-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [Дисплайневмессажеформ (Parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#displaynewmessageform-parameters-) | ReadItem | Чтение | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | Создание<br>Чтение | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | Создание<br>Чтение | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>События

Вы можете подписаться на следующие события и отписаться на них, используя [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) и [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#removehandlerasync-eventtype--options--callback-) соответственно.

| Событие | Описание | Minimum<br>набор требований |
|---|---|:---:|
|`ItemChanged`| Для просмотра выбран другой элемент Outlook, когда область задач закреплена. | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

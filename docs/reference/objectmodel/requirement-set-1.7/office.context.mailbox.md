---
title: Office. Context. Mailbox — набор обязательных элементов 1,7
description: Набор требований API почтового ящика Outlook 1,7 версия объектной модели почтового ящика.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: ad48175f94d592e825d48b0403c272ff44f3ee49
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612215"
---
# <a name="mailbox-requirement-set-17"></a>Mailbox (набор требований 1,7)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)| С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

## <a name="properties"></a>Свойства

| Свойство | Минимальные<br>уровень разрешения | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|---|:---:|
| [диагностики](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#diagnostics) | ReadItem | Создание<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#ewsurl) | ReadItem | Создание<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [элемента](office.context.mailbox.item.md) | Restricted | Создание<br>Read | [Ресурс](/javascript/api/outlook/office.item?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#resturl) | ReadItem | Создание<br>Read | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#userprofile) | ReadItem | Создание<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Минимальные<br>уровень разрешения | Способов | Минимальные<br>набор требований |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Создание<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [Конверттоевсид (itemId, Рестверсион)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#converttoewsid-itemid--restversion-) | Restricted | Создание<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#converttolocalclienttime-timevalue-) | ReadItem | Создание<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, Рестверсион)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#converttorestid-itemid--restversion-) | Restricted | Создание<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (вход)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#converttoutcclienttime-input-) | ReadItem | Создание<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#displayappointmentform-itemid-) | ReadItem | Создание<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#displaymessageform-itemid-) | ReadItem | Создание<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#displaynewappointmentform-parameters-) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [Дисплайневмессажеформ (Parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#displaynewmessageform-parameters-) | ReadItem | Создание<br>Read | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#getcallbacktokenasync-options--callback-) | ReadItem | Создание<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#getcallbacktokenasync-callback--usercontext-) | ReadItem | Создание<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Создание<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Создание<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | ReadItem | Создание<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>События

Вы можете подписаться на следующие события и отписаться на них, используя [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) и [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) соответственно.

| Событие | Описание | Минимальные<br>набор требований |
|---|---|:---:|
|`ItemChanged`| Для просмотра выбран другой элемент Outlook, когда область задач закреплена. | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

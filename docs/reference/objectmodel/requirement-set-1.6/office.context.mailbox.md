---
title: Office. Context. Mailbox — набор обязательных элементов 1,6
description: ''
ms.date: 03/06/2020
localization_priority: Normal
ms.openlocfilehash: 13d4021baded0203b9e94cf38dced4e6afec796b
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/07/2020
ms.locfileid: "42562028"
---
# <a name="mailbox"></a>mailbox

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#diagnostics) | ReadItem | Создание<br>Чтение | [Диагностики](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#ewsurl) | ReadItem | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [элемента](office.context.mailbox.item.md) | Restricted | Создание<br>Чтение | [Элемент](/javascript/api/outlook/office.item?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#resturl) | ReadItem | Создание<br>Чтение | Строка | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#userprofile) | ReadItem | Создание<br>Чтение | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Минимальные<br>уровень разрешения | Способов | Минимальные<br>набор требований |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Создание<br>Чтение | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [Конверттоевсид (itemId, Рестверсион)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttoewsid-itemid--restversion-) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttolocalclienttime-timevalue-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, Рестверсион)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttorestid-itemid--restversion-) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (вход)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttoutcclienttime-input-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displayappointmentform-itemid-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaymessageform-itemid-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaynewappointmentform-parameters-) | ReadItem | Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [Дисплайневмессажеформ (Parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaynewmessageform-parameters-) | ReadItem | Создание<br>Чтение | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getcallbacktokenasync-options--callback-) | ReadItem | Создание<br>Чтение | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getcallbacktokenasync-callback--usercontext-) | ReadItem | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#removehandlerasync-eventtype--options--callback-) | ReadItem | Создание<br>Чтение | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>События

Вы можете подписаться на следующие события и отписаться на них, используя [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#addhandlerasync-eventtype--handler--options--callback-) и [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#removehandlerasync-eventtype--options--callback-) соответственно.

| Событие | Описание | Минимальные<br>набор требований |
|---|---|:---:|
|`ItemChanged`| Для просмотра выбран другой элемент Outlook, когда область задач закреплена. | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

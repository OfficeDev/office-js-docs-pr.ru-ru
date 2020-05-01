---
title: Office. Context. Mailbox — набор обязательных элементов 1,4
description: Набор требований API почтового ящика Outlook 1,4 версия объектной модели почтового ящика.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 683f9d8418d57c40e79269c36fd28319bd301f1e
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890534"
---
# <a name="mailbox-requirement-set-14"></a>Mailbox (набор требований 1,4)

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#diagnostics) | ReadItem | Создание<br>Чтение | [Диагностики](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#ewsurl) | ReadItem | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [элемента](office.context.mailbox.item.md) | Restricted | Создание<br>Чтение | [Элемент](/javascript/api/outlook/office.item?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#userprofile) | ReadItem | Создание<br>Чтение | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Минимальные<br>уровень разрешения | Способов | Минимальные<br>набор требований |
|---|---|---|:---:|
| [Конверттоевсид (itemId, Рестверсион)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#converttoewsid-itemid--restversion-) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#converttolocalclienttime-timevalue-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, Рестверсион)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#converttorestid-itemid--restversion-) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (вход)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#converttoutcclienttime-input-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#displayappointmentform-itemid-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#displaymessageform-itemid-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#displaynewappointmentform-parameters-) | ReadItem | Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#getcallbacktokenasync-callback--usercontext-) | ReadItem | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

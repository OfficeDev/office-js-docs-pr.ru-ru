---
title: Office.context.mailbox — набор требований 1.6
description: Outlook Требования К API почтовых ящиков устанавливают версию 1.6 объектной модели почтовых ящиков.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: a1583e326e8a38cf6f0144638785221b9b7b44f3
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237646"
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
| [диагностика](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#diagnostics) | ReadItem | Создание<br>Чтение | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#ewsUrl) | ReadItem | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [элемента](office.context.mailbox.item.md) | Restricted | Создание<br>Чтение | [Элемент](/javascript/api/outlook/office.item?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#restUrl) | ReadItem | Создание<br>Чтение | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#userProfile) | ReadItem | Создание<br>Чтение | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Методы

| Метод | Minimum<br>уровень разрешения | Режимы | Minimum<br>набор требований |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#convertToEwsId_itemId__restVersion_) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#convertToLocalClientTime_timeValue_) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#convertToRestId_itemId__restVersion_) | Restricted | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (вход)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#convertToUtcClientTime_input_) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displayAppointmentForm_itemId_) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displayMessageForm_itemId_) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displayNewAppointmentForm_parameters_) | ReadItem | Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm (параметры)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displayNewMessageForm_parameters_) | ReadItem | Чтение | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getCallbackTokenAsync_options__callback_) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getCallbackTokenAsync_callback__userContext_) | ReadItem | Создание<br>Чтение | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getUserIdentityTokenAsync_callback__userContext_) | ReadItem | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#makeEwsRequestAsync_data__callback__userContext_) | ReadWriteMailbox | Создание<br>Чтение | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | ReadItem | Создание<br>Чтение | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>События

Вы можете подписаться и отписаться от следующих событий с помощью [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) и [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#removeHandlerAsync_eventType__options__callback_) соответственно.

> [!IMPORTANT]
> События доступны только с реализацией области задач.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.6&preserve-view=true) | Описание | Minimum<br>набор требований |
|---|---|:---:|
|`ItemChanged`| Другой элемент Outlook для просмотра при закреплении области задач. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

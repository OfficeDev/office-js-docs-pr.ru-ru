---
title: Просмотр и изменение данных элемента в форме создания элементов Outlook
description: Просматривайте и устанавливайте различные свойства элемента в надстройке Outlook при сценарии создания, такие как сведения о получателях, тема, текст, а также место и время встречи.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2ae4b6a30d08199207faf89079c57fbff46d6a0e
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467240"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Просмотр и изменение данных элемента в форме создания элементов Outlook

Сведения о том, как получать и задавать различные свойства элемента в надстройке Outlook в сценарии создания, такие как сведения о получателях, тема, текст, а также место и время встречи.

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>Получение и установка свойств элемента для надстройки создания

В формах создания можно получить доступ к большинству свойств, предоставляемых таким типом элемента в форме чтения (например, участники, получатели, тема и текст), а несколько дополнительных свойств доступны только в форме создания (текст, СК).

For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The  [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [item.conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.

Кроме доступа к свойствам элементов в API JavaScript для Office, доступ к свойствам на уровне элементов можно получить с помощью веб-служб Exchange (EWS). С разрешением на чтение и запись почтового ящика можно использовать метод [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) для доступа к операциям EWS, [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) и [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation), чтобы получить и задать дополнительные свойства элемента или элементов в почтовом ящике пользователя.

Метод `makeEwsRequestAsync` доступен как в формах создания, так и для чтения. Дополнительные сведения о разрешении на чтение и запись почтового ящика и доступе к веб-службам EWS через платформу надстроек Office см. в статье "Общие сведения о разрешениях надстроек [Outlook](understanding-outlook-add-in-permissions.md) и вызове веб-служб из надстройки [Outlook"](web-services.md).

**Таблица 1. Асинхронные методы для просмотра или изменения свойств элемента в форме создания**

| Свойство | Тип свойства | Асинхронный метод для получения свойства | Асинхронные методы для задания |
|:-----|:-----|:-----|:-----|
|[bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Recipients](/javascript/api/outlook/office.recipients)|[Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))|[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))|
|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Body](/javascript/api/outlook/office.body)|[Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)), [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))|
|[cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Time](/javascript/api/outlook/office.time)|[Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1))|[Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))|
|[location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Location](/javascript/api/outlook/office.location)|[Location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1))|[Location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))|
|[optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Time|Time.getAsync|Time.setAsync|
|[subject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Subject](/javascript/api/outlook/office.subject)|[Subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1))|[Subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))|
|[to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>См. также

- [Создание надстроек Outlook для форм создания](compose-scenario.md)
- [Общие сведения о разрешениях для надстройки Outlook](understanding-outlook-add-in-permissions.md)
- [Вызов веб-служб из надстройки Outlook](web-services.md)
- [Считывание и запись данных элемента Outlook в формах просмотра и создания](item-data.md)

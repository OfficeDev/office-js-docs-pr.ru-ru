---
title: Просмотр и изменение данных элемента в форме создания элементов Outlook
description: Просматривайте и устанавливайте различные свойства элемента в надстройке Outlook при сценарии создания, такие как сведения о получателях, тема, текст, а также место и время встречи.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: f888e0f5a9d1d3c3ab64a174064f3b2984111229
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53670288"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Просмотр и изменение данных элемента в форме создания элементов Outlook

Сведения о том, как получать и задавать различные свойства элемента в надстройке Outlook в сценарии создания, такие как сведения о получателях, тема, текст, а также место и время встречи.

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>Получение и установка свойств элемента для надстройки создания

В формах создания можно получить доступ к большинству свойств, предоставляемых таким типом элемента в форме чтения (например, участники, получатели, тема и текст), а несколько дополнительных свойств доступны только в форме создания (текст, СК).

Методы получения и задания большинства этих свойств асинхронные, так как надстройка Outlook и пользователь могут изменять одно свойство в пользовательском интерфейсе одновременно. В таблице 1 перечислены свойства уровня элемента и соответствующие асинхронные методы, позволяющие их получить и задать в форме создания. Исключение составляют свойства [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [item.conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), потому что пользователи не могут их менять. Их можно получить программно в форме создания так же, как и в форме чтения, напрямую из родительского объекта.

Кроме доступа к свойствам элементов Office API JavaScript, вы можете получить доступ к свойствам уровня элементов с помощью Exchange веб-служб (EWS). Имея разрешение **ReadWriteMailbox**, вы можете получать и задавать дополнительные свойства элементов в почтовом ящике пользователя, используя метод [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) для доступа к операциям EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) и [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation).

Функция `makeEwsRequestAsync` доступна как в формах создания, так и в формах чтения. Дополнительные сведения о разрешении **ReadWriteMailbox** и доступе к EWS с помощью платформы надстроек Office см. в статьях [Общие сведения о разрешениях для надстроек Outlook](understanding-outlook-add-in-permissions.md) и [Вызов веб-служб из надстройки Outlook](web-services.md).

**Таблица 1. Асинхронные методы для просмотра или изменения свойств элемента в форме создания**

<br/>

| Свойство | Тип свойства | Асинхронный метод для получения свойства | Асинхронные методы для установки свойства |
|:-----|:-----|:-----|:-----|
|[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Recipients](/javascript/api/outlook/office.Recipients)|[Recipients.getAsync](/javascript/api/outlook/office.Recipients#getAsync_options__callback_)|[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addAsync_recipients__options__callback_), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setAsync_recipients__options__callback_)|
|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Body](/javascript/api/outlook/office.Body)|[Body.getAsync](/javascript/api/outlook/office.Body#getAsync_coercionType__options__callback_)|[Body.prependAsync](/javascript/api/outlook/office.Body#prependAsync_data__options__callback_), [Body.setAsync](/javascript/api/outlook/office.Body#setAsync_data__options__callback_), [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setSelectedDataAsync_data__options__callback_)|
|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Time](/javascript/api/outlook/office.Time)|[Time.getAsync](/javascript/api/outlook/office.Time#getAsync_options__callback_)|[Time.setAsync](/javascript/api/outlook/office.Time#setAsync_dateTime__options__callback_)|
|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Location](/javascript/api/outlook/office.Location)|[Location.getAsync](/javascript/api/outlook/office.Location#getAsync_options__callback_)|[Location.setAsync](/javascript/api/outlook/office.Location#setAsync_location__options__callback_)|
|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Time|Time.getAsync|Time.setAsync|
|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Subject](/javascript/api/outlook/office.Subject)|[Subject.getAsync](/javascript/api/outlook/office.Subject#getAsync_options__callback_)|[Subject.setAsync](/javascript/api/outlook/office.Subject#setAsync_subject__options__callback_)|
|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>См. также

- [Создание надстроек Outlook для форм создания](compose-scenario.md)
- [Общие сведения о разрешениях для надстройки Outlook](understanding-outlook-add-in-permissions.md)
- [Вызов веб-служб из надстройки Outlook](web-services.md)
- [Считывание и запись данных элемента Outlook в формах просмотра и создания](item-data.md)

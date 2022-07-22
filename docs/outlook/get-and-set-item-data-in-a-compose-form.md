---
title: Просмотр и изменение данных элемента в форме создания элементов Outlook
description: Просматривайте и устанавливайте различные свойства элемента в надстройке Outlook при сценарии создания, такие как сведения о получателях, тема, текст, а также место и время встречи.
ms.date: 12/10/2019
ms.localizationpriority: medium
ms.openlocfilehash: ddc6cd0011060bc49d1fd5cd8e6c9ceebb2a8c08
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958981"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>Просмотр и изменение данных элемента в форме создания элементов Outlook

Сведения о том, как получать и задавать различные свойства элемента в надстройке Outlook в сценарии создания, такие как сведения о получателях, тема, текст, а также место и время встречи.

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>Получение и установка свойств элемента для надстройки создания

В формах создания можно получить доступ к большинству свойств, предоставляемых таким типом элемента в форме чтения (например, участники, получатели, тема и текст), а несколько дополнительных свойств доступны только в форме создания (текст, СК).

Методы получения и задания большинства этих свойств асинхронные, так как надстройка Outlook и пользователь могут изменять одно свойство в пользовательском интерфейсе одновременно. В таблице 1 перечислены свойства уровня элемента и соответствующие асинхронные методы, позволяющие их получить и задать в форме создания. Исключение составляют свойства [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) и [item.conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), потому что пользователи не могут их менять. Их можно получить программно в форме создания так же, как и в форме чтения, напрямую из родительского объекта.

Кроме доступа к свойствам элементов в API JavaScript для Office, доступ к свойствам на уровне элементов можно получить с помощью веб-служб Exchange (EWS). Имея разрешение **ReadWriteMailbox**, вы можете получать и задавать дополнительные свойства элементов в почтовом ящике пользователя, используя метод [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) для доступа к операциям EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) и [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation).

Метод `makeEwsRequestAsync` доступен как в формах создания, так и для чтения. Дополнительные сведения о разрешении **ReadWriteMailbox** и доступе к EWS с помощью платформы надстроек Office см. в статьях [Общие сведения о разрешениях для надстроек Outlook](understanding-outlook-add-in-permissions.md) и [Вызов веб-служб из надстройки Outlook](web-services.md).

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

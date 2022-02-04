---
title: Просмотр и изменение данных элемента в форме создания элементов Outlook
description: 'Просматривайте и устанавливайте различные свойства элемента в надстройке Outlook при сценарии создания, такие как сведения о получателях, тема, текст, а также место и время встречи.'
ms.date: 12/10/2019
ms.localizationpriority: medium
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
|[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Recipients](/javascript/api/outlook/office.recipients)|[Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))|[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))|
|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Body](/javascript/api/outlook/office.body)|[Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)), [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))|
|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Time](/javascript/api/outlook/office.time)|[Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1))|[Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))|
|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Location](/javascript/api/outlook/office.location)|[Location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1))|[Location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))|
|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Time|Time.getAsync|Time.setAsync|
|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Subject](/javascript/api/outlook/office.subject)|[Subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1))|[Subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))|
|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>См. также

- [Создание надстроек Outlook для форм создания](compose-scenario.md)
- [Общие сведения о разрешениях для надстройки Outlook](understanding-outlook-add-in-permissions.md)
- [Вызов веб-служб из надстройки Outlook](web-services.md)
- [Считывание и запись данных элемента Outlook в формах просмотра и создания](item-data.md)

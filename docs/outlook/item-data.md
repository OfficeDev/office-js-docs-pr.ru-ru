---
title: Просмотр и изменение данных элемента в надстройке Outlook
description: В зависимости от активации надстройки в форме чтения или создания элемента, свойства, доступные надстройке для элемента, отличаются.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 925dcaff524be3f4e3d76213c0ee3413a11f43d6
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325356"
---
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>Просмотр и изменение данных элемента Outlook в формах чтения и создания

Начиная с версии 1.1 схемы манифестов для надстроек Office, Outlook может активировать надстройки, когда пользователь просматривает или создает элемент. В зависимости от активации надстройки в форме чтения или создания элемента, свойства, доступные надстройке для элемента, так же отличаются.

Например, свойства [dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) определены только для элемента, который уже был отправлен (элемент затем просматривается в форме чтения), но не для элемента, который создается (в форме создания). Другим примером является свойство [bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), которое имеет смысл, если сообщение создается (в форме создания), и недоступно пользователю в форме чтения.

## <a name="item-properties-available-in-compose-and-read-forms"></a>Свойства элементов, доступные в формах создания и чтения элементов

В таблице 1 показаны свойства на уровне элемента в API JavaScript для Office, доступные в каждом режиме (чтение и создание) почтовых надстроек. обычно эти свойства, доступные в формах чтения, доступны только для чтения, и они доступны в формах создания для чтения и записи, за исключением свойств [ItemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), [conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)и [ItemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) , которые всегда доступны только для чтения.

Для остальных свойств на уровне элемента, доступных в формах создания (поскольку надстройка и пользователь могут просматривать или записывать одно и то же свойство одновременно), применяются асинхронные методы просмотра или изменения в режиме создания, поэтому типы объектов, возвращаемых этими свойствами, также могут отличаться в формах создания и чтения. Дополнительные сведения об использовании асинхронных методов просмотра или изменения свойств на уровне элементов в режиме создания см. статью [Просмотр и изменение данных элемента в форме создания элементов Outlook](get-and-set-item-data-in-a-compose-form.md).


**Таблица 1. Свойства элементов, доступные в формах создания и просмотра элементов**

<br/>

|**Тип элемента**|**Свойство**|**Тип свойства в формах просмотра элементов**|**Тип свойства в формах создания элементов**|
|:-----|:-----|:-----|:-----|
|Встречи и сообщения|[dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Объект JavaScript **Date**|Свойство недоступно|
|Встречи и сообщения|[dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Объект JavaScript **Date**|Свойство недоступно|
|Встречи и сообщения|[itemClass](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Строка|Свойство недоступно|
|Встречи и сообщения|[itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Строка|Свойство недоступно|
|Встречи и сообщения|[itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Строка в перечислении [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)|Строка в перечислении [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) (только для чтения)|
|Встречи и сообщения|[attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|Свойство недоступно|
|Встречи и сообщения|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Body](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|Встречи и сообщения|[normalizedSubject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Строка|Свойство недоступно|
|Встречи и сообщения|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Строка|[Subject](/javascript/api/outlook/office.subject)|
|Встречи|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Объект JavaScript **Date**|[Time](/javascript/api/outlook/office.time)|
|Встречи|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Строка|[Location](/javascript/api/outlook/office.location)|
|Встречи|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Встречи|[organizer](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)|
|Встречи|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Встречи|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Объект JavaScript **Date**|[Time](/javascript/api/outlook/office.time)|
|Сообщения|[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Свойство недоступно|[Получатели](/javascript/api/outlook/office.recipients)|
|Сообщения|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Получатели](/javascript/api/outlook/office.recipients)|
|Сообщения|[conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Строка|String (только для чтения)|
|Сообщения|[from](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)|
|Сообщения|[internetMessageId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Целое число|Свойство недоступно|
|Сообщения|[sender](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|Свойство недоступно|
|Сообщения|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Получатели](/javascript/api/outlook/office.recipients)|

## <a name="use-exchange-server-callback-tokens-from-a-read-add-in"></a>Использование маркеров обратного вызова Exchange Server из надстройки для просмотра элементов

Если надстройка Outlook активирована в формах просмотра элементов, вы можете получить маркер обратного вызова Exchange. Этот маркер можно использовать в серверном коде для доступа ко всему элементу через веб-службы Exchange (EWS).

Указывая разрешение **ReadItem** в манифесте надстройки, вы можете применить метод [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) для получения маркера обратного вызова Exchange, а также свойство [mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) для получения URL-адреса конечной точки EWS для почтового ящика пользователя и [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), чтобы получить идентификатор EWS для выбранного элемента. Затем передайте маркер обратного вызова, URL-адрес конечной точки EWS и идентификатор элемента EWS в серверный код для доступа к операции [GetItem](/exchange/client-developer/web-service-reference/getitem-operation), что позволить получить больше свойств для элемента.


## <a name="access-ews-from-a-read-or-compose-add-in"></a>Доступ к веб-службам EWS из надстройки для просмотра или создания элементов

Вы также можете использовать метод [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods), чтобы получить доступ к операциям веб-служб Exchange (EWS) [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) и [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) непосредственно из надстройки. Эти операции можно использовать для просмотра и изменения многих свойств заданного элемента. Этот метод доступен для надстроек Outlook независимо от активации надстройки в форме чтения или создания элемента, при условии указания разрешения **ReadWriteMailbox** в манифесте надстройки.

Дополнительные сведения об использовании метода **makeEwsRequestAsync** для получения доступа к операциям EWS см. в статье [Вызов веб-служб из надстройки Outlook](web-services.md).


## <a name="see-also"></a>См. также

- [Просмотр и изменение данных элемента в форме создания элементов Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Вызов веб-служб из надстройки Outlook](web-services.md)

---
title: Office.context.mailbox.item — набор обязательных элементов 1.5
description: ''
ms.date: 05/30/2019
localization_priority: Priority
ms.openlocfilehash: aba1590d570a55ed512f1eb223b7d927c953dbaf
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706339"
---
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [attachments](#attachments-arrayattachmentdetails) | Элемент |
| [bcc](#bcc-recipients) | Элемент |
| [body](#body-body) | Элемент |
| [cc](#cc-arrayemailaddressdetailsrecipients) | Элемент |
| [conversationId](#nullable-conversationid-string) | Элемент |
| [dateTimeCreated](#datetimecreated-date) | Элемент |
| [dateTimeModified](#datetimemodified-date) | Элемент |
| [end](#end-datetime) | Элемент |
| [from](#from-emailaddressdetails) | Элемент |
| [internetMessageId](#internetmessageid-string) | Элемент |
| [itemClass](#itemclass-string) | Элемент |
| [itemId](#nullable-itemid-string) | Элемент |
| [itemType](#itemtype-officemailboxenumsitemtype) | Элемент |
| [location](#location-stringlocation) | Элемент |
| [normalizedSubject](#normalizedsubject-string) | Элемент |
| [notificationMessages](#notificationmessages-notificationmessages) | Элемент |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsrecipients) | Элемент |
| [organizer](#organizer-emailaddressdetails) | Элемент |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsrecipients) | Member |
| [sender](#sender-emailaddressdetails) | Элемент |
| [start](#start-datetime) | Элемент |
| [subject](#subject-stringsubject) | Элемент |
| [to](#to-arrayemailaddressdetailsrecipients) | Элемент |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | Метод |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | Метод |
| [close](#close) | Метод |
| [displayReplyAllForm](#displayreplyallformformdata-callback) | Метод |
| [displayReplyForm](#displayreplyformformdata-callback) | Метод |
| [getEntities](#getentities--entities) | Метод |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | Метод |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | Метод |
| [getRegExMatches](#getregexmatches--object) | Метод |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | Метод |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | Метод |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | Метод |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | Метод |
| [saveAsync](#saveasyncoptions-callback) | Метод |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | Метод |

### <a name="example"></a>Пример

В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.

```javascript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

### <a name="members"></a>Members

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a>attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)>

Получает массив вложений для элемента. Только в режиме чтения.

> [!NOTE]
> Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются. Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Тип

*   Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)>

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.

```javascript
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

#### <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a>bcc: [Recipients](/javascript/api/outlook_1_5/office.recipients)

Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения. Только в режиме создания.

##### <a name="type"></a>Тип

*   [Получатели](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание|

##### <a name="example"></a>Пример

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook15officebody"></a>body: [Body](/javascript/api/outlook_1_5/office.body)

Получает объект, предоставляющий методы для работы с основным текстом элемента.

##### <a name="type"></a>Тип

*   [Body](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

В этом примере возвращается текст сообщения в формате обычного текста.

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

Ниже приведен пример итогового параметра, переданного функции обратного вызова.

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a>cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)

Предоставляет доступ к получателям копии сообщения. Тип объекта и уровень доступа зависят от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a>Режим создания

Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Тип

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

#### <a name="nullable-conversationid-string"></a>(nullable) conversationId: String

Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.

Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.

Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a>dateTimeCreated: Date

Получает дату и время создания элемента. Только в режиме чтения.

##### <a name="type"></a>Тип

*   Дата

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a>dateTimeModified: Date

Получает дату и время последнего изменения элемента. Только в режиме чтения.

> [!NOTE]
> Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="type"></a>Тип

*   Дата

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook15officetime"></a>end: Date|[Time](/javascript/api/outlook_1_5/office.time)

Получает или задает дату и время окончания встречи.

Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).

##### <a name="read-mode"></a>Режим чтения

Свойство `end` возвращает объект `Date`.

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a>Режим создания

Свойство `end` возвращает объект `Time`.

Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a>Тип

*   Date | [Time](/javascript/api/outlook_1_5/office.time)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a>from: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

Получает электронный адрес отправителя сообщения. Только в режиме чтения.

Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

> [!NOTE]
> Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.

##### <a name="type"></a>Тип

*   [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a>internetMessageId: String

Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass: String

Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.

Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.

| Тип | Описание | Класс элемента |
| --- | --- | --- |
| Элементы встречи | Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| Элементы сообщения | Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId: String

Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.

> [!NOTE]
> Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange. Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook. Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).

Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a>itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

Получает тип элемента, который представляет экземпляр.

Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.

##### <a name="type"></a>Тип

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook15officelocation"></a>location: String|[Location](/javascript/api/outlook_1_5/office.location)

Получает или задает место встречи.

##### <a name="read-mode"></a>Режим чтения

Свойство `location` возвращает строку, содержащую сведения о месте встречи.

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a>Режим создания

Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a>Тип

*   String | [Location](/javascript/api/outlook_1_5/office.location)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

#### <a name="normalizedsubject-string"></a>normalizedSubject: String

Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.

Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a>notificationMessages: [NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)

Получает сообщения уведомления для элемента.

##### <a name="type"></a>Тип

*   [NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a>optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)

Предоставляет доступ к необязательным участникам события. Тип объекта и уровень доступа зависят от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a>Режим создания

Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Тип

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a>organizer: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

Получает электронный адрес организатора указанного собрания. Только в режиме чтения.

##### <a name="type"></a>Тип

*   [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a>requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)

Предоставляет доступ к обязательным участникам события. Тип объекта и уровень доступа зависят от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a>Режим создания

Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a>Тип

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a>sender: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.

Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

> [!NOTE]
> Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.

##### <a name="type"></a>Тип

*   [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook15officetime"></a>start: Date|[Time](/javascript/api/outlook_1_5/office.time)

Получает или задает дату и время начала встречи.

Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).

##### <a name="read-mode"></a>Режим чтения

Свойство `start` возвращает объект `Date`.

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a>Режим создания

Свойство `start` возвращает объект `Time`.

Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a>Тип

*   Date | [Time](/javascript/api/outlook_1_5/office.time)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

#### <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a>subject: String|[Subject](/javascript/api/outlook_1_5/office.subject)

Получает или задает описание, которое отображается в поле темы элемента.

Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.

##### <a name="read-mode"></a>Режим чтения

Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a>Режим создания

Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a>Тип

*   String | [Subject](/javascript/api/outlook_1_5/office.subject)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a>to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)

Предоставляет доступ к получателям, указанным в строке **Кому** сообщения. Тип объекта и уровень доступа зависят от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a>Режим создания

Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Тип

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="methods"></a>Методы

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Добавляет файл в сообщение или встречу в качестве вложения.

Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.

Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`uri`| String||Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.|
|`attachmentName`| String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.|
|`options`| Object| &lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
| `options.asyncContext` | Object | &lt;необязательно&gt; | В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ. |
| `options.isInline` | Boolean | &lt;Необязательный&gt; | Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений. |
|`callback`| function| &lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

| Код ошибки | Описание |
|------------|-------------|
| `AttachmentSizeExceeded` | Вложение превышает максимальный размер. |
| `FileTypeNotSupported` | Расширение вложения не поддерживается. |
| `NumberOfAttachmentsExceeded` | Сообщение или встреча содержат слишком много вложений. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание|

##### <a name="examples"></a>Примеры

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.

```javascript
Office.context.mailbox.item.addFileAttachmentAsync(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.

С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.

Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`itemId`| String||Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.|
|`attachmentName`| String||Тема вкладываемого элемента. Максимальная длина: 255 символов.|
|`options`| Object| &lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Object| &lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`| функция| &lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

| Код ошибки | Описание |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | Сообщение или встреча содержат слишком много вложений. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание|

##### <a name="example"></a>Пример

В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="close"></a>close()

Закрывает текущий создаваемый элемент.

Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.

> [!NOTE]
> Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.

Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание|

#### <a name="displayreplyallformformdata-callback"></a>displayReplyAllForm(formData, [callback])

Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### <a name="parameters"></a>Параметры

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
|`formData`| String &#124; Object| |Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта. |
| `formData.htmlBody` | String | &lt;необязательно&gt; | Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;необязательно&gt; | Массив объектов JSON, представляющих собой вложенные файлы или элементы. |
| `formData.attachments.type` | String | | Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента. |
| `formData.attachments.name` | String | | Строка, содержащая имя вложения, длиной до 255 символов.|
| `formData.attachments.url` | String | | Используется, только если свойству `type` задано значение `file`. URI расположения файла. |
| `formData.attachments.isInline` | Логический | | Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений. |
| `formData.attachments.itemId` | String | | Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов. |
| `callback` | function | &lt;Необязательный&gt; | По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="examples"></a>Примеры

Приведенный ниже код передает строку в функцию `displayReplyAllForm`.

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

Ответ только с текстом сообщения.

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Ответ с текстом сообщения и вложенным элементом.

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a>displayReplyForm(formData, [callback])

Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### <a name="parameters"></a>Параметры

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
|`formData`| String &#124; Object| | Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта. |
| `formData.htmlBody` | String | &lt;необязательно&gt; | Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;необязательно&gt; | Массив объектов JSON, представляющих собой вложенные файлы или элементы. |
| `formData.attachments.type` | String | | Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента. |
| `formData.attachments.name` | String | | Строка, содержащая имя вложения, длиной до 255 символов.|
| `formData.attachments.url` | String | | Используется, только если свойству `type` задано значение `file`. URI расположения файла. |
| `formData.attachments.isInline` | Логический | | Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений. |
| `formData.attachments.itemId` | String | | Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов. |
| `callback` | function | &lt;Необязательный&gt; | По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="examples"></a>Примеры

Приведенный ниже код передает строку в функцию `displayReplyForm`.

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

Ответ только с текстом сообщения.

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Ответ с текстом сообщения и вложенным элементом.

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a>getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}

Получает сущности, обнаруженные в теле выбранного элемента.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [Entities](/javascript/api/outlook_1_5/office.entities)

##### <a name="example"></a>Пример

Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}

Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|Одно из значений перечисления EntityType.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL. Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив. В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.

Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.

| Значение параметра `entityType` | Тип объектов в возвращаемом массиве | Необходимый уровень разрешений |
| --- | --- | --- |
| `Address` | String | **Restricted** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restricted** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **Restricted** |

Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>

##### <a name="example"></a>Пример

В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.

```javascript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}

Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.

Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.

Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) для этого.

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.

<dl class="param-type">

<dt>Тип</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>Пример

В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (nullable) {Array.< String >}

Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.

<dl class="param-type">

<dt>Тип</dt>

<dd>Array.< String ></dd>

</dl>

##### <a name="example"></a>Пример

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Асинхронно возвращает данные, выбранные в теме или тексте сообщения.

Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](office.md#coerciontype-string)||Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).|
|`options`| Object| &lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Object| &lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`| функция||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`. Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание|

##### <a name="returns"></a>Возвращаемое значение:

Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.

<dl class="param-type">

<dt>Тип</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>Пример

```javascript
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.

Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) в свойстве `asyncResult.value`. Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.|
|`userContext`| Объект| &lt;Необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова. Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

#### <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Удаляет вложение из сообщения или встречи.

Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`attachmentId`| String||Идентификатор удаляемого вложения.|
|`options`| Object| &lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Object| &lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`| функция| &lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.|

##### <a name="errors"></a>Ошибки

| Код ошибки | Описание |
|------------|-------------|
| `InvalidAttachmentId` | Идентификатор вложения не существует. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание|

##### <a name="example"></a>Пример

Указанный ниже код удаляет вложение с идентификатором "0".

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

#### <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Асинхронно сохраняет элемент.

При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.

> [!NOTE]
> Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных. До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.

Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.

> [!NOTE]
> Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:
>
> - Outlook для Mac не поддерживает сохранение собрания. Метод `saveAsync` не работает при вызове из собрания в режиме создания. Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).
> - Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`options`| Object| &lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Object| &lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание|

##### <a name="examples"></a>Примеры

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Асинхронно вставляет данные в текст или тему сообщения.

Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.|
|`options`| Object| &lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Object| &lt;необязательно&gt;|В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.|
|`options.coercionType`| [Office.CoercionType](office.md#coerciontype-string)| &lt;необязательный&gt;|Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.<br/><br/>Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.<br/><br/>Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание|

##### <a name="example"></a>Пример

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```

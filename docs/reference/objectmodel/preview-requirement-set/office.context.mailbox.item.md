---
title: Office. Context. Mailbox. Item — Предварительная версия набора требований
description: ''
ms.date: 06/11/2019
localization_priority: Normal
ms.openlocfilehash: 9844fdeba274271a226501d6c0a8694e164f6ac7
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910275"
---
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).

##### <a name="requirements"></a>Requirements

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [attachments](#attachments-arrayattachmentdetails) | Элемент |
| [bcc](#bcc-recipients) | Элемент |
| [body](#body-body) | Элемент |
| [разделов](#categories-categories) | Элемент |
| [cc](#cc-arrayemailaddressdetailsrecipients) | Элемент |
| [conversationId](#nullable-conversationid-string) | Элемент |
| [dateTimeCreated](#datetimecreated-date) | Элемент |
| [dateTimeModified](#datetimemodified-date) | Элемент |
| [end](#end-datetime) | Элемент |
| [Енханцедлокатион](#enhancedlocation-enhancedlocation) | Элемент |
| [from](#from-emailaddressdetailsfrom) | Элемент |
| [Internetheaders:](#internetheaders-internetheaders) | Элемент |
| [internetMessageId](#internetmessageid-string) | Элемент |
| [itemClass](#itemclass-string) | Элемент |
| [itemId](#nullable-itemid-string) | Элемент |
| [itemType](#itemtype-officemailboxenumsitemtype) | Элемент |
| [location](#location-stringlocation) | Элемент |
| [normalizedSubject](#normalizedsubject-string) | Элемент |
| [notificationMessages](#notificationmessages-notificationmessages) | Member |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsrecipients) | Элемент |
| [organizer](#organizer-emailaddressdetailsorganizer) | Элемент |
| [recurrence](#nullable-recurrence-recurrence) | Элемент |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsrecipients) | Элемент |
| [sender](#sender-emailaddressdetails) | Member |
| [seriesId](#nullable-seriesid-string) | Member |
| [start](#start-datetime) | Member |
| [subject](#subject-stringsubject) | Элемент |
| [to](#to-arrayemailaddressdetailsrecipients) | Элемент |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | Метод |
| [addFileAttachmentFromBase64Async](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | Метод |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Метод |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | Метод |
| [close](#close) | Метод |
| [displayReplyAllForm](#displayreplyallformformdata-callback) | Метод |
| [displayReplyForm](#displayreplyformformdata-callback) | Метод |
| [Жетаттачментконтентасинк](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | Метод |
| [Жетаттачментсасинк](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | Метод |
| [getEntities](#getentities--entities) | Метод |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | Метод |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | Метод |
| [getInitializationContextAsync](#getinitializationcontextasyncoptions-callback) | Метод |
| [Жетитемидасинк](#getitemidasyncoptions-callback) | Метод |
| [getRegExMatches](#getregexmatches--object) | Метод |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | Метод |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | Метод |
| [Жетселектедентитиес](#getselectedentities--entities) | Метод |
| [Жетселектедрежексматчес](#getselectedregexmatches--object) | Метод |
| [Жетшаредпропертиесасинк](#getsharedpropertiesasyncoptions-callback) | Метод |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | Метод |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | Метод |
| [removeHandlerAsync](#removehandlerasynceventtype-options-callback) | Метод |
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

### <a name="members"></a>Элементы

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

Получает вложения элемента в виде массива. Только в режиме чтения.

> [!NOTE]
> Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются. Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Тип

*   Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

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

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a>СК: [получатели](/javascript/api/outlook/office.recipients)

Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения. Только в режиме создания.

##### <a name="type"></a>Тип

*   [Получатели](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

##### <a name="example"></a>Пример

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a>основной текст: [Body](/javascript/api/outlook/office.body)

Получает объект, предоставляющий методы для работы с основным текстом элемента.

##### <a name="type"></a>Тип

*   [Body](/javascript/api/outlook/office.body)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

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

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a>Категории: [категории](/javascript/api/outlook/office.categories)

Получает объект, предоставляющий методы для управления категориями элемента.

> [!NOTE]
> Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="type"></a>Тип

*   [Categories](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

В этом примере возвращаются категории элемента.

```javascript
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)

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

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

---
---

#### <a name="nullable-conversationid-string"></a>(Nullable) conversationId: строка

Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.

Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.

Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a>dateTimeCreated: Дата

Получает дату и время создания элемента. Только в режиме чтения.

##### <a name="type"></a>Тип

*   Дата

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a>dateTimeModified: Дата

Получает дату и время последнего изменения элемента. Только в режиме чтения.

> [!NOTE]
> Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="type"></a>Тип

*   Дата

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a>конец: Дата | [Time (время](/javascript/api/outlook/office.time) )

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

Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.

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

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>Енханцедлокатион: [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation)

Получает или задает расположение встречи.

##### <a name="read-mode"></a>Режим чтения

Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который позволяет получить набор расположений (каждый, представленный объектом локатиондетаилс), связанный с встречей. [](/javascript/api/outlook/office.locationdetails) `enhancedLocation`

##### <a name="compose-mode"></a>Режим создания

`enhancedLocation` Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который предоставляет методы для получения, удаления или добавления расположений для встречи.

##### <a name="type"></a>Тип

*   [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

В следующем примере показано получение текущих расположений, связанных с встречей.

```javascript
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a>от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[](/javascript/api/outlook/office.from)

Получает электронный адрес отправителя сообщения.

Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

> [!NOTE]
> Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.

##### <a name="read-mode"></a>Режим чтения

`from` Свойство возвращает `EmailAddressDetails` объект.

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a>Режим создания

`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a>Тип

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)

##### <a name="requirements"></a>Требования

|Требование|||
|---|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|Создание|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)

Возвращает или задает заголовки Интернета сообщения.

##### <a name="type"></a>Тип

*   [InternetHeaders](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a>internetMessageId: строка

Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a>itemClass: строка

Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.

Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.

|Тип|Описание|Класс элемента|
|---|---|---|
|Элементы встречи|Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|Элементы сообщения|Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a>(Nullable) itemId: строка

Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.

> [!NOTE]
> Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange. Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook. Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).

Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

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

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a>itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

Получает тип элемента, который представляет экземпляр.

Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.

##### <a name="type"></a>Тип

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a>Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location) )

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

*   String | [Location](/javascript/api/outlook/office.location)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

---
---

#### <a name="normalizedsubject-string"></a>normalizedSubject: строка

Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.

Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a>notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)

Получает сообщения уведомления для элемента.

##### <a name="type"></a>Тип

*   [NotificationMessages](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)

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

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a>Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Организатор](/javascript/api/outlook/office.organizer)

Получает адрес электронной почты организатора для указанного собрания.

##### <a name="read-mode"></a>Режим чтения

`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , представляющий организатора собрания.

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a>Режим создания

Свойство возвращает объект организатора, который предоставляет метод для получения значения организатора. [](/javascript/api/outlook/office.organizer) `organizer`

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a>Тип

*   [](/javascript/api/outlook/office.emailaddressdetails) | [Организатор](/javascript/api/outlook/office.organizer) EmailAddressDetails

##### <a name="requirements"></a>Требования

|Требование|||
|---|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|Создание|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a>(Nullable) повторение [](/javascript/api/outlook/office.recurrence) : повторение

Получает или задает шаблон повторения встречи. Получает шаблон повторения приглашения на собрание. Режимы чтения и создания для элементов встречи. Режим чтения для элементов приглашения на собрания.

`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду. `null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч. `undefined`возвращается для сообщений, которые не являются приглашениями на собрания.

> Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.

> Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.

##### <a name="read-mode"></a>Режим чтения

`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , представляющий повторение встречи. Оно доступно для встреч и приглашений на собрания.

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a>Режим создания

`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , который предоставляет методы для управления повторением встречи. Оно доступно для встреч.

```javascript
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a>Тип

* [Повторения](/javascript/api/outlook/office.recurrence)

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.7|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)

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

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a>Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.

Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

> [!NOTE]
> Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.

##### <a name="type"></a>Тип

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a>(Nullable) seriesId: строка

Получает идентификатор ряда, к которому принадлежит экземпляр.

В OWA и Outlook `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент. Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.

> [!NOTE]
> Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange. `seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook. Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).

`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.

##### <a name="type"></a>Тип

* String

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.7|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a>Начало: Дата | [Time (время](/javascript/api/outlook/office.time) )

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

Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.

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

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a>Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject) )

Получает или задает описание, которое отображается в поле темы элемента.

Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.

##### <a name="read-mode"></a>Режим чтения

Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.

В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.

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

*   String | [Subject](/javascript/api/outlook/office.subject)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)

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

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

### <a name="methods"></a>Методы

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Добавляет файл в сообщение или встречу в качестве вложения.

Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.

Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

##### <a name="parameters"></a>Параметры
|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`uri`|String||Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.|
|`attachmentName`|String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.|
|`options`|Объект|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;Необязательно&gt;|В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.|
|`options.isInline`|Boolean|&lt;необязательно&gt;|Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.|
|`callback`|function|&lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`AttachmentSizeExceeded`|Вложение превышает максимальный размер.|
|`FileTypeNotSupported`|Расширение вложения не поддерживается.|
|`NumberOfAttachmentsExceeded`|Сообщение или встреча содержат слишком много вложений.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

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

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов])

Добавляет файл из кодировки Base64 в сообщение или встречу в качестве вложения.

`addFileAttachmentFromBase64Async` Метод передает файл из кодировки Base64 и вкладывает его в элемент в форме создания. Этот метод возвращает идентификатор вложения в объекте AsyncResult. Value.

Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`base64File`|String||Содержимое изображения или файла в кодировке Base64, которое добавляется в сообщение электронной почты или событие.|
|`attachmentName`|String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.|
|`options`|Объект|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;Необязательно&gt;|В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.|
|`options.isInline`|Boolean|&lt;необязательно&gt;|Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.|
|`callback`|function|&lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`AttachmentSizeExceeded`|Вложение превышает максимальный размер.|
|`FileTypeNotSupported`|Расширение вложения не поддерживается.|
|`NumberOfAttachmentsExceeded`|Сообщение или встреча содержат слишком много вложений.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

##### <a name="examples"></a>Примеры

```javascript
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

Добавляет обработчик для поддерживаемого события.

В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.

##### <a name="parameters"></a>Параметры

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Событие, которое должно вызвать обработчик. |
| `handler` | Function || Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`. |
| `options` | Объект | &lt;Необязательно&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.asyncContext` | Объект | &lt;Необязательно&gt; | Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова. |
| `callback` | функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение |

##### <a name="example"></a>Пример

```javascript
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.

С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.

Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`itemId`|String||Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.|
|`attachmentName`|String||Тема вкладываемого элемента. Максимальная длина: 255 символов.|
|`options`|Object|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;Необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция|&lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|Сообщение или встреча содержат слишком много вложений.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

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

---
---

#### <a name="close"></a>close()

Закрывает текущий создаваемый элемент.

Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.

> [!NOTE]
> Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.

Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

---
---

#### <a name="displayreplyallformformdata-callback"></a>displayReplyAllForm(formData, [callback])

Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`formData`|String &#124; Object||Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.|
|`formData.htmlBody`|String|&lt;необязательно&gt;|Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.
|`formData.attachments`|Array.&lt;Object&gt;|&lt;необязательно&gt;|Массив объектов JSON, представляющих собой вложенные файлы или элементы.|
|`formData.attachments.type`|String||Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.|
|`formData.attachments.name`|Строка||Строка, содержащая имя вложения, длиной до 255 символов.|
|`formData.attachments.url`|String||Используется, только если свойству `type` задано значение `file`. URI расположения файла.|
|`formData.attachments.isInline`|Логический||Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.|
|`formData.attachments.itemId`|String||Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.|
|`callback`|function|&lt;Необязательный&gt;|По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

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

---
---

#### <a name="displayreplyformformdata-callback"></a>displayReplyForm(formData, [callback])

Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`formData`|String &#124; Object||Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.|
|`formData.htmlBody`|String|&lt;необязательно&gt;|Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.
|`formData.attachments`|Array.&lt;Object&gt;|&lt;необязательно&gt;|Массив объектов JSON, представляющих собой вложенные файлы или элементы.|
|`formData.attachments.type`|String||Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.|
|`formData.attachments.name`|Строка||Строка, содержащая имя вложения, длиной до 255 символов.|
|`formData.attachments.url`|String||Используется, только если свойству `type` задано значение `file`. URI расположения файла.|
|`formData.attachments.isInline`|Логический||Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.|
|`formData.attachments.itemId`|String||Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.|
|`callback`|function|&lt;Необязательный&gt;|По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

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

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>Жетаттачментконтентасинк (attachmentId, [параметры], [callback]) → [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)

Получает указанное вложение из сообщения или встречи и возвращает его в виде `AttachmentContent` объекта.

`getAttachmentContentAsync` Метод получает вложение с указанным идентификатором из элемента. Рекомендуется использовать идентификатор для получения вложения в том же сеансе, когда Аттачментидс был получен с помощью вызова `getAttachmentsAsync` или. `item.attachments` В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`attachmentId`|String||Идентификатор вложения, которое требуется получить.|
|`options`|Объект|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция|&lt;Необязательный&gt;|По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)

##### <a name="example"></a>Пример

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>Жетаттачментсасинк ([параметры], [обратный вызов]) → массив. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

Получает вложения элемента в виде массива. Только в режиме создания.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`options`|Object|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция|&lt;Необязательный&gt;|По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

##### <a name="returns"></a>Возвращаемое значение:

Тип: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="example"></a>Пример

В приведенном ниже примере создается строка HTML со сведениями обо всех вложениях в текущем элементе.

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a>getEntities() → {[Entities](/javascript/api/outlook/office.entities)}

Получает сущности, обнаруженные в теле выбранного элемента.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>Пример

Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Описание|
|---|---|---|
|`entityType`|[Office.MailboxEnums.EntityType](/javascript/api/outlook/office.mailboxenums.entitytype)|Одно из значений перечисления EntityType.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL. Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив. В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.

Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.

|Значение параметра `entityType`|Тип объектов в возвращаемом массиве|Необходимый уровень разрешений|
|---|---|---|
|`Address`|String|**Restricted**|
|`Contact`|Contact|**ReadItem**|
|`EmailAddress`|String|**ReadItem**|
|`MeetingSuggestion`|MeetingSuggestion|**ReadItem**|
|`PhoneNumber`|PhoneNumber|**Restricted**|
|`TaskSuggestion`|TaskSuggestion|**ReadItem**|
|`URL`|String|**Restricted**|

Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

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
};
```

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Описание|
|---|---|---|
|`name`|String|Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.

Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a>getInitializationContextAsync ([параметры], [обратный вызов])

Получает данные инициализации, передаваемые при активации надстройки [сообщением с действиями](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

> [!NOTE]
> Этот метод поддерживается только в Outlook 2016 или более поздней версии для Windows ("нажми и работай" более поздней версии, чем 16.0.8413.1000) и Outlook в Интернете для Office 365.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`options`|Object|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;Необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция|&lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>При успешном выполнении данные инициализации предоставляются в `asyncResult.value` свойстве в виде строки.<br/>Если `asyncResult` контекст инициализации отсутствует, объект будет содержать `Error` объект со `code` свойством, `9020` `name` для свойства которого задано значение. `GenericResponseError`|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```javascript
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

---
---

#### <a name="getitemidasyncoptions-callback"></a>Жетитемидасинк ([параметры], обратный вызов)

Асинхронно получает идентификатор сохраненного элемента. Только в режиме создания.

При вызове этот метод возвращает идентификатор элемента с помощью метода обратного вызова.

> [!NOTE]
> Если надстройка вызывает `getItemIdAsync` элемент в режиме создания (например, чтобы получить доступ `itemId` к использованию с помощью EWS или REST API), имейте в виду, что если Outlook находится в режиме кэширования, может потребоваться некоторое время до синхронизации элемента с сервером. Пока элемент не будет синхронизирован, он не `itemId` распознается и не будет использоваться, возвращается сообщение об ошибке.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`options`|Object|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;Необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`ItemNotSaved`|Идентификатор невозможно извлечь, пока не будет сохранен элемент.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

##### <a name="examples"></a>Примеры

```javascript
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

В следующем примере показана структура `result` параметра, переданного функции обратного вызова. `value` Свойство содержит идентификатор элемента.

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

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

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.

##### <a name="requirements"></a>Requirements

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.

<dl class="param-type">Тип:<dd>Object</dd>

</dl>

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (nullable) {Array.< String >}

Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Описание|
|---|---|---|
|`name`|String|Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

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

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Асинхронно возвращает данные, выбранные в теме или тексте сообщения.

Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`coercionType`|[Office.CoercionType](office.md#coerciontype-string)||Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).|
|`options`|Объект|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`. Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.2|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

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

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a>getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}

Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.6|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.

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

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.

##### <a name="requirements"></a>Requirements

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.6|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a>Жетшаредпропертиесасинк ([параметры], обратный вызов)

Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`options`|Object|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;Необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Общие свойства предоставляются в виде [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объекта в `asyncResult.value` свойстве. Этот объект можно использовать для получения общих свойств элемента.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.

Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`callback`|function||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`. Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.|
|`userContext`|Объект|&lt;Необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова. Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание или чтение|

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

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Удаляет вложение из сообщения или встречи.

Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`attachmentId`|String||Идентификатор удаляемого вложения.|
|`options`|Объект|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;Необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция|&lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`InvalidAttachmentId`|Идентификатор вложения не существует.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

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

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a>removeHandlerAsync(eventType, [options], [callback])

Удаляет обработчиков для поддерживаемого типа события.

В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.

##### <a name="parameters"></a>Параметры

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Событие, которое должно отменить обработчик. |
| `options` | Объект | &lt;Необязательно&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.asyncContext` | Объект | &lt;Необязательно&gt; | Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова. |
| `callback` | функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение |

---
---

#### <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Асинхронно сохраняет элемент.

При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.

> [!NOTE]
> Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных. До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.

Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.

> [!NOTE]
> Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:
>
> - Outlook для Mac не поддерживает сохранение собраний. `saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания. Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.
> - Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`options`|Объект|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;Необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|функция||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

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

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Асинхронно вставляет данные в текст или тему сообщения.

Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`data`|String||Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.|
|`options`|Object|&lt;Необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Объект|&lt;Необязательно&gt;|В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.|
|`options.coercionType`|[Office.CoercionType](office.md#coerciontype-string)|&lt;необязательный&gt;|Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.<br/><br/>Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.<br/><br/>Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.|
|`callback`|функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.2|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)|Создание|

##### <a name="example"></a>Пример

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```

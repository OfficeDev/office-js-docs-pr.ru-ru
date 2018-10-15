
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="members-and-methods"></a>Члены и методы

| Член | Тип |
|--------|------|
| [attachments](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | Член |
| [bcc](#bcc-recipientsjavascriptapioutlook16officerecipients) | Член |
| [body](#body-bodyjavascriptapioutlook16officebody) | Член |
| [cc](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | Член |
| [conversationId](#nullable-conversationid-string) | Член |
| [dateTimeCreated](#datetimecreated-date) | Член |
| [dateTimeModified](#datetimemodified-date) | Член |
| [end](#end-datetimejavascriptapioutlook16officetime) | Член |
| [from](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | Член |
| [internetMessageId](#internetmessageid-string) | Член |
| [itemClass](#itemclass-string) | Член |
| [itemId](#nullable-itemid-string) | Член |
| [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | Член |
| [location](#location-stringlocationjavascriptapioutlook16officelocation) | Член |
| [normalizedSubject](#normalizedsubject-string) | Член |
| [notificationMessages](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | Член |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | Член |
| [organizer](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | Член |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | Член |
| [sender](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | Член |
| [start](#start-datetimejavascriptapioutlook16officetime) | Член |
| [subject](#subject-stringsubjectjavascriptapioutlook16officesubject) | Член |
| [to](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | Член |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | Метод |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | Метод |
| [close](#close) | Метод |
| [displayReplyAllForm](#displayreplyallformformdata) | Метод |
| [displayReplyForm](#displayreplyformformdata) | Метод |
| [getEntities](#getentities--entitiesjavascriptapioutlook16officeentities) | Метод |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | Метод |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | Метод |
| [getRegExMatches](#getregexmatches--object) | Метод |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | Метод |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | Метод |
| [getSelectedEntities](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | Метод |
| [getSelectedRegExMatches](#getselectedregexmatches--object) | Метод |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | Метод |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | Метод |
| [saveAsync](#saveasyncoptions-callback) | Метод |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | Метод |

### <a name="example"></a>Пример

В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.

```
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
}
```

### <a name="members"></a>Члены

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a>attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)>

Получает массив вложений для элемента. Только в режиме чтения.

> [!NOTE]
> Некоторые типы файлов блокируются Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются. Дополнительные сведения см. в статье [Блокированные вложения в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Тип:

*   Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)>

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a>bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)

Получает объект, который предоставляет методы для получения или обновления получателей в строке Bcc (скрытой копии) сообщения. Только в режиме создания.

##### <a name="type"></a>Тип:

*   [Recipients](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a>body :[Body](/javascript/api/outlook_1_6/office.body)

Получает объект, предоставляющий методы для работы с основным текстом элемента.

##### <a name="type"></a>Тип:

*   [Body](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a>cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)

Предоставляет доступ к получателям Cc (копии) сообщения. Уровень доступа и тип объекта, зависит от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 членов.

##### <a name="compose-mode"></a>Режим создания

Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Cc (копия)** сообщения.

##### <a name="type"></a>Тип:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.

Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.

Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

Получает дату и время создания элемента. Только в режиме чтения.

##### <a name="type"></a>Тип:

*   Date

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

Получает дату и время последнего изменения элемента. Только в режиме чтения.

> [!NOTE]
> Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="type"></a>Тип:

*   Date

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a>end :Date|[Time](/javascript/api/outlook_1_6/office.time)

Получает или задает дату и время окончания встречи.

Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime).

##### <a name="read-mode"></a>Режим чтения

Свойство `end` возвращает объект `Date`.

##### <a name="compose-mode"></a>Режим создания

Свойство `end` возвращает объект `Time`.

Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

##### <a name="type"></a>Тип:

*   Date | [Time](/javascript/api/outlook_1_6/office.time)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a>from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

Получает электронный адрес отправителя сообщения. Только в режиме чтения.

Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

> [!NOTE]
> Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `from` — `undefined`.

##### <a name="type"></a>Тип:

*   [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

#### <a name="internetmessageid-string"></a>internetMessageId :String

Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.

Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.

| Тип | Описание | item class |
| --- | --- | --- |
| Элементы встречи | Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Элементы сообщения | Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например, настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.

> [!NOTE]
> Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange. Свойство  `itemId` не совпадает с идентификатором записи Outlook или идентификатором, используемым API-Интерфейсом REST Outlook. Прежде чем осуществлять вызовы API-Интерфейса REST с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Дополнительные сведения см. в статье [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).

Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a>itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

Получает тип элемента, который представляет экземпляр.

Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.

##### <a name="type"></a>Тип:

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a>location :String|[Location](/javascript/api/outlook_1_6/office.location)

Получает или задает место встречи.

##### <a name="read-mode"></a>Режим чтения

Свойство `location` возвращает строку, содержащую сведения о месте встречи.

##### <a name="compose-mode"></a>Режим создания

Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.

##### <a name="type"></a>Тип:

*   String | [Location](/javascript/api/outlook_1_6/office.location)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.

Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a>notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)

Получает сообщения уведомления для элемента.

##### <a name="type"></a>Тип:

*   [NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a>optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)

Предоставляет доступ к необязательным участникам события. Уровень доступа и тип объекта, зависит от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.

##### <a name="compose-mode"></a>Режим создания

Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления необязательных участников собрания.

##### <a name="type"></a>Тип:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a>organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

Получает электронный адрес организатора указанного собрания. Только в режиме чтения.

##### <a name="type"></a>Тип:

*   [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a>requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)

Предоставляет доступ к обязательным участникам события. Уровень доступа и тип объекта, зависит от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.

##### <a name="compose-mode"></a>Режим создания

Свойство `requiredAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления обязательных участников собрания.

##### <a name="type"></a>Тип:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a>sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.

Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

> [!NOTE]
> Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `sender` — `undefined`.

##### <a name="type"></a>Тип:

*   [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a>start :Date|[Time](/javascript/api/outlook_1_6/office.time)

Получает или задает дату и время начала встречи.

Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime).

##### <a name="read-mode"></a>Режим чтения

Свойство `start` возвращает объект `Date`.

##### <a name="compose-mode"></a>Режим создания

Свойство `start` возвращает объект `Time`.

Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

##### <a name="type"></a>Тип:

*   Date | [Time](/javascript/api/outlook_1_6/office.time)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a>subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)

Получает или задает описание, которое отображается в поле темы элемента.

Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.

##### <a name="read-mode"></a>Режим чтения

Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>Режим создания

Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>Тип:

*   String | [Subject](/javascript/api/outlook_1_6/office.subject)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a>to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)

Предоставляет доступ получателей к строке **To (Кому)** в сообщении. Уровень доступа и тип объекта, зависит от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **To (Кому)** сообщения. Коллекция может включать не более 100 элементов.

##### <a name="compose-mode"></a>Режим создания

Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **To**(Кому) сообщения.

##### <a name="type"></a>Тип:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>Методы

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Добавляет файл в сообщение или встречу в качестве вложения.

Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.

Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`uri`| String||Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.|
|`attachmentName`| String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.|
|`options`| Объект| &lt;необязательный&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
| `options.asyncContext` | Объект | &lt;необязательный&gt; | Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова. |
| `options.isInline` | Логический | &lt;необязательный&gt; | Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений. |
|`callback`| function| &lt;необязательный&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

| Код ошибки | Описание |
|------------|-------------|
| `AttachmentSizeExceeded` | Вложение превышает максимальный размер. |
| `FileTypeNotSupported` | Расширение вложения не поддерживается. |
| `NumberOfAttachmentsExceeded` | Сообщение или встреча содержат слишком много вложений. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание|

##### <a name="examples"></a>Примеры

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
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
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Добавляет к сообщению элемент Exchange, например, сообщение, в виде вложения.

С помощью метода `addItemAttachmentAsync` в элемент формы создания можно вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии в метод обратного вызова.

Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

Если ваша надстройка Office выполняется в веб-приложении Outlook, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`itemId`| String||Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.|
|`attachmentName`| String||Тема вкладываемого элемента. Максимальная длина — 255 символов.|
|`options`| Объект| &lt;необязательный&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Объект| &lt;необязательный&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`| function| &lt;необязательный&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

| Код ошибки | Описание |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | Сообщение или встреча содержат слишком много вложений. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание|

##### <a name="example"></a>Пример

В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a>close()

Закрывает текущий создаваемый элемент.

Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.

> [!NOTE]
> В Outlook в Интернете, если элемент является встречей, и он был ранее сохранен с помощью `saveAsync`, пользователю предлагается сохранить, отменить или удалить, даже если не произошло каких-либо изменений, поскольку элемент был последним сохраненным.

Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание|

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строчный параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### <a name="parameters"></a>Параметры:

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
|`formData`| String | Object| |Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта. |
| `formData.htmlBody` | String | &lt;необязательный&gt; | Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;необязательный&gt; | Массив объектов JSON, представляющих собой вложенные файлы или элементы. |
| `formData.attachments.type` | String | | Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента. |
| `formData.attachments.name` | String | | Строка, содержащая имя вложения, длиной до 255 символов.|
| `formData.attachments.url` | String | | Используется, только если свойству `type` задано значение `file`. URI расположения файла. |
| `formData.attachments.isInline` | Логический | | Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений. |
| `formData.attachments.itemId` | String | | Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов. |
| `callback` | function | &lt;необязательный&gt; | По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="examples"></a>Примеры

Приведенный ниже код передает строку в функцию `displayReplyAllForm`.

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```
Office.context.mailbox.item.displayReplyAllForm({});
```

Ответ только с текстом сообщения.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```
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

```
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

```
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

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строчный параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### <a name="parameters"></a>Параметры:

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
|`formData`| String | Object| | Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта. |
| `formData.htmlBody` | String | &lt;необязательный&gt; | Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;необязательный&gt; | Массив объектов JSON, представляющих собой вложенные файлы или элементы. |
| `formData.attachments.type` | String | | Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента. |
| `formData.attachments.name` | String | | Строка, содержащая имя вложения, длиной до 255 символов.|
| `formData.attachments.url` | String | | Используется, только если свойству `type` задано значение `file`. URI расположения файла. |
| `formData.attachments.isInline` | Логический | | Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений. |
| `formData.attachments.itemId` | String | | Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов. |
| `callback` | function | &lt;необязательный&gt; | По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="examples"></a>Примеры

Приведенный ниже код передает строку в функцию `displayReplyForm`.

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```
Office.context.mailbox.item.displayReplyForm({});
```

Ответ только с текстом сообщения.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```
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

```
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

```
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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a>getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}

Получает сущности, обнаруженные в выбранном элементе.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [Entities](/javascript/api/outlook_1_6/office.entities)

##### <a name="example"></a>Пример

Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a>getEntitiesByType(entityType) → (допускающий значение NULL) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}

Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|Одно из значений перечисления EntityType.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL. Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив. В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.

Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.

| Значение параметра `entityType` | Тип объектов в возвращаемом массиве | Необходимый уровень разрешений |
| --- | --- | --- |
| `Address` | String | **Ограниченный доступ** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Ограниченный доступ** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **Ограниченный доступ** |

Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>

##### <a name="example"></a>Пример

В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a>getFilteredEntitiesByName(name) → (допускающий значение NULL) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}

Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.

Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Возвращает строчные значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.

Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) для этого.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.

<dl class="param-type">

<dt>Тип</dt>

<dd>Объект</dd>

</dl>

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (nullable) {Array.< String >}

Возвращает строчные значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.

<dl class="param-type">

<dt>Тип</dt>

<dd>Array.< String ></dd>

</dl>

##### <a name="example"></a>Пример

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Асинхронно возвращает данные, выбранные в теме или тексте сообщения.

Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](office.md#coerciontype-string)||Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).|
|`options`| Объект| &lt;необязательный&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Объект| &lt;необязательный&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`| function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`. Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите  `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание|

##### <a name="returns"></a>Возвращаемое значение:

Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.

<dl class="param-type">

<dt>Тип</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>Пример

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a>getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}

Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [Entities](/javascript/api/outlook_1_6/office.entities)

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.

Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) для этого.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Асинхронно загружает настраиваемые свойства для надстройки выбранного элемента.

Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) в свойстве `asyncResult.value`. Этот объект можно использовать для получения, задания и удаления настраиваемых свойств из элемента и сохранения изменений настраиваемого свойства на сервере.|
|`userContext`| Объект| &lt;необязательный&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова. Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Cоздание или чтение|

##### <a name="example"></a>Пример

В приведенном ниже примере кода показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. В этом примере кода, после того как выполнена загрузка настраиваемых свойств, метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Удаляет вложение из сообщения или встречи.

Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В веб-приложении Outlook и веб-приложении Outlook для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`attachmentId`| String||Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.|
|`options`| Объект| &lt;необязательный&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Объект| &lt;необязательный&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`| function| &lt;необязательный&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.|

##### <a name="errors"></a>Ошибки

| Код ошибки | Описание |
|------------|-------------|
| `InvalidAttachmentId` | Идентификатор вложения не существует. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание|

##### <a name="example"></a>Пример

Указанный ниже код удаляет вложение с идентификатором "0".

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Асинхронно сохраняет элемент.

При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.

> [!NOTE]
> Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных. До окончания синхронизации применение параметра `itemId`  будет приводить к ошибке.

Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.

> [!NOTE]
> Следующие клиенты имеют разную реакцию на событие для `saveAsync` для встреч в режиме создания:
>
> - Mac Outlook не поддерживает `saveAsync` на собрании в режиме создания. Вызов `saveAsync` на собрании в Mac Outlook возвращает ошибку.
> - Outlook в Интернете всегда отправляет приглашение или обновления при вызове `saveAsync` на встрече в режиме создания.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`options`| Объект| &lt;необязательный&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Объект| &lt;необязательный&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`| function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание|

##### <a name="examples"></a>Примеры

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Асинхронно вставляет данные в текст или тему сообщения.

Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.|
|`options`| Объект| &lt;необязательный&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`| Объект| &lt;необязательный&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`options.coercionType`| [Office.CoercionType](office.md#coerciontype-string)| &lt;необязательный&gt;|Если задано значение `text`, текущий стиль применяется в Outlook и веб-приложении Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.<br/><br/>Если `html` и поле поддерживают HTML (тема не поддерживает), в веб-приложении Outlook применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.<br/><br/>Если тип `coercionType` не установлен, результат зависит от поля: если поле имеет формат HTML, то используется HTML; если поле является текстовым, то используется обычный текст.|
|`callback`| function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
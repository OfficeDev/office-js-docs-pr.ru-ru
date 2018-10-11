
# <a name="mailbox"></a>mailbox

### [Office](Office.md)[.context](Office.context.md). mailbox

Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Члены и методы

| Член | Тип |
|--------|------|
| [ewsUrl](#ewsurl-string) | Член |
| [restUrl](#resturl-string) | Член |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Метод |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | Метод |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) | Метод |
| [convertToRestId](#converttorestiditemid-restversion--string) | Метод |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Метод |
| [displayAppointmentForm](#displayappointmentformitemid) | Метод |
| [displayMessageForm](#displaymessageformitemid) | Метод |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | Метод |
| [displayNewMessageForm](#displaynewmessageformparameters) | Метод |
| [getCallbackTokenAsync](#getcallbacktokenasyncoptions-callback) | Метод |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Метод |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Метод |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Метод |

### <a name="namespaces"></a>Пространства имен

[diagnostics](Office.context.mailbox.diagnostics.md): предоставляет надстройке Outlook диагностические сведения.

[item](Office.context.mailbox.item.md): предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.

[userProfile](Office.context.mailbox.userProfile.md): предоставляет сведения о пользователе в надстройке Outlook.

### <a name="members"></a>Члены

#### <a name="ewsurl-string"></a>ewsUrl :String

Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.

> [!NOTE]
> Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.

Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

В манифесте приложения должно быть указано разрешение **ReadItem** для вызова метода `ewsUrl` в режиме чтения.

Перед использованием члена `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

#### <a name="resturl-string"></a>restUrl :String

Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.

С помощью значения `restUrl` можно выполнять вызовы [REST API](https://docs.microsoft.com/outlook/rest/) для почтового ящика пользователя.

В манифесте приложения должно быть указано разрешение **ReadItem** для вызова метода `restUrl` в режиме чтения.

Перед использованием члена `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="methods"></a>Методы

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

Добавляет обработчик для поддерживаемого события.

В настоящее время поддерживаемые типы событий: `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.

##### <a name="parameters"></a>Параметры:

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Событие, которое должно вызвать обработчик. |
| `handler` | Функция || Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`. |
| `options` | Объект | &lt;необязательный&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.asyncContext` | Объект | &lt;необязательный&gt; | Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова. |
| `callback` | function| &lt;необязательный&gt;|Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

Преобразует идентификатор элемента из формата REST в формат EWS.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Формат идентификаторов, извлекаемых через API REST (например [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразует идентификатор из формата REST в формат EWS.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате REST API для Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion)|Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: String

##### <a name="example"></a>Пример

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}

Получает словарь, содержащий информацию о времени в локальном времени клиента.

Для даты и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.

Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`timeValue`| Date|Объект Date|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)

####  <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

Преобразует идентификатор элемента из формата EWS в формат REST.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Формат идентификаторов, извлекаемых через EWS или через свойство `itemId`, отличается от формата API REST (таких как [API почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)). Метод `convertToRestId` преобразует идентификатор из формата EWS в формат REST.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате EWS|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion)|Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: String

##### <a name="example"></a>Пример

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

Получает объект Date из словаря, содержащего сведения о времени.

Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)|Значение локального времени для преобразования.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект Date со временем в формате UTC.

<dl class="param-type">

<dt>Тип</dt>

<dd>Date</dd>

</dl>

####  <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

Отображает имеющуюся встречу из календаря.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.

В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).

В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит не более 32 КБ символов.

Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор веб-служб Exchange для существующей встречи в календаре.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

Отображает имеющееся сообщение.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.

В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит не более 32 КБ символов.

Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.

Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор веб-служб Exchange (EWS) для существующего сообщения.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

Отображает форму для создания новой встречи в календаре.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.

В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.

Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.

Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.

##### <a name="parameters"></a>Параметры:

> [!NOTE]
> Примечание. Все параметры являются необязательными.

|Имя| Тип| Описание|
|---|---|---|
| `parameters` | Объект | Словарь параметров, описывающий новую встречу. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей. |
| `parameters.start` | Date | Объект `Date`, указывающий дату и время начала встречи. |
| `parameters.end` | Date | Объект `Date`, указывающий дату и время окончания встречи. |
| `parameters.location` | String | Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255. |
| `parameters.resources` | Array.&lt;String&gt; | Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей. |
| `parameters.subject` | String | Строка с темой встречи. Максимальное количество символов в строке — 255. |
| `parameters.body` | String | Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Read|

##### <a name="example"></a>Пример

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="displaynewmessageformparameters"></a>displayNewMessageForm(parameters)

Отображает форму для создания сообщения.

Метод `displayNewMessageForm` открывает форму, в которой пользователь может создать сообщение. Если параметры заданы, поля формы сообщения автоматически заполняются их содержимым.

Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.

##### <a name="parameters"></a>Параметры:

> [!NOTE]
> Примечание. Все параметры являются необязательными.

|Имя| Тип| Описание|
|---|---|---|
| `parameters` | Объект | Словарь параметров, описывающий новое сообщение. |
| `parameters.toRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Кому". Массив может включать не более 100 записей. |
| `parameters.ccRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Копия". Массив может включать не более 100 записей. |
| `parameters.bccRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Скрытая копия". Массив может включать не более 100 записей. |
| `parameters.subject` | String | Строка с темой сообщения. Максимальное количество символов в строке — 255. |
| `parameters.htmlBody` | String | Текст сообщения в формате HTML. Максимальный размер содержимого сообщения — 32 КБ. |
| `parameters.attachments` | Array.&lt;Object&gt; | Массив объектов JSON, представляющих собой вложенные файлы или элементы. |
| `parameters.attachments.type` | String | Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента. |
| `parameters.attachments.name` | String | Строка, содержащая имя вложения, длиной до 255 символов.|
| `parameters.attachments.url` | String | Используется, только если свойству `type` задано значение `file`. URI расположения файла. |
| `parameters.attachments.isInline` | Логический | Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений. |
| `parameters.attachments.itemId` | String | Используется только в том случае, если свойству `type` присвоено значение `item`. Идентификатор элемента веб-служб Exchange существующего сообщения электронной почты, которые необходимо присоединить к новому сообщению. Это строка длиной до 100 символов. |


##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync([options], callback)

Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.

Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.

> [!NOTE]
> Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange. 

**Маркеры REST**

Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.

С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.

**Маркеры EWS**

Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.

С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
| `options` | Объект | &lt;необязательный&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.isRest` | Логический |  &lt;необязательный&gt; | Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию — `false`. |
| `options.asyncContext` | Объект |  &lt;необязательный&gt; | Данные о состоянии, передаваемые в асинхронный метод. |
|`callback`| function||Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание и чтение|

##### <a name="example"></a>Пример

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.

Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.

Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

В манифесте приложения должно быть указано разрешение **ReadItem** для вызова метода `getCallbackTokenAsync` в режиме чтения.

Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function||Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.|
|`userContext`| Объект| &lt;необязательный&gt;|Данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание и чтение|

##### <a name="example"></a>Пример

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

Получает маркер, идентифицирующий пользователя и надстройку Office.

Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Маркер указывается в виде строки в свойстве `asyncResult.value`.|
|`userContext`| Объект| &lt;необязательный&gt;|Данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.

> [!NOTE]
> Этот метод не поддерживается в следующих сценариях.
> - В Outlook для iOS или Outlook для Android.
> - Когда надстройка загружается в почтовом ящике Gmail
> 
> Вместо этого надстройкам следует использовать [API-интерфейсы REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.

Метод `makeEwsRequestAsync` отправляет к Exchange EWS-запрос от имени надстройки. См. [Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) для ознакомления с информацией о списке поддерживаемых операций веб-служб Exchange.

С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.

XML-запрос должен указывать кодировку UTF-8.

```
<?xml version="1.0" encoding="utf-8"?>
```

У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> Администратор сервера должен установить для `OAuthAuthentication` значение true в каталоге сервера клиентского доступа EWS, чтобы включить метод  `makeEwsRequestAsync` для запросов служб EWS.

##### <a name="version-differences"></a>Различия версий

Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии, предшествующей 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Запрос EWS.|
|`callback`| function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`. Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.|
|`userContext`| Объект| &lt;необязательный&gt;|Данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

В следующем примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```
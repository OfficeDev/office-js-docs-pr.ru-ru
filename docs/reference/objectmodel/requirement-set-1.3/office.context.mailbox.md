
# <a name="mailbox"></a>mailbox

### [Office](Office.md)[.context](Office.context.md). mailbox

Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="namespaces"></a>Пространства имен

[diagnostics](Office.context.mailbox.diagnostics.md): предоставляет надстройке Outlook диагностические сведения.

[item](Office.context.mailbox.item.md): предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.

[userProfile](Office.context.mailbox.userProfile.md): предоставляет сведения о пользователе в надстройке Outlook.

### <a name="members"></a>Члены

#### <a name="ewsurl-string"></a>ewsUrl: String

Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.

> [!NOTE]
> Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.

Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

Чтобы вызвать член `ewsUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.

Перед использованием члена `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="methods"></a>Методы

####  <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

Преобразует идентификатор элемента из формата REST в формат EWS.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Формат идентификаторов, извлекаемых через API REST (например [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразует идентификатор из формата REST в формат EWS.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате REST API для Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
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

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}

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
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)

####  <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

Преобразует идентификатор элемента из формата EWS в формат REST.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Формат идентификаторов, извлекаемых через EWS или через свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)). Метод `convertToRestId` преобразует идентификатор из формата EWS в формат REST.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате EWS|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
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
|`input`| [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)|Значение локального времени для преобразования.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
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

В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или образец встречи из повторяющегося ряда, но не экземпляр ряда, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).

В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит не более 32 КБ символов.

Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор веб-служб Exchange (EWS) для существующей встречи в календаре.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

Отображает существующее сообщение.

> [!NOTE]
> Этот метод не поддерживается в Outlook для iOS или Outlook для Android.

Метод `displayMessageForm` открывает существующее сообщение в новом окне на компьютере или в диалоговом окне на мобильных устройствах.

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
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
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

Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, создается исключение.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Описание|
|---|---|---|
| `parameters` | Object | Словарь параметров, описывающий новую встречу. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 элементов. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 элементов. |
| `parameters.start` | Date | Объект `Date`, указывающий дату и время начала встречи. |
| `parameters.end` | Date | Объект `Date`, указывающий дату и время окончания встречи. |
| `parameters.location` | String | Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255. |
| `parameters.resources` | Array.&lt;String&gt; | Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 элементов. |
| `parameters.subject` | String | Строка с темой встречи. Максимальное количество символов в строке — 255. |
| `parameters.body` | String | Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Чтение|

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.

Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.

Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

Для вызова метода `getCallbackTokenAsync` в режиме чтения в манифесте приложения должно быть указано разрешение **ReadItem**.

Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function||Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.|
|`userContext`| Object| &lt;необязательный&gt;|Любые данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
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
|`callback`| function||Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Маркер указывается в виде строки в свойстве `asyncResult.value`.|
|`userContext`| Object| &lt;необязательный&gt;|Любые данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
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

Выполняет асинхронный запрос к веб-службе Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.

> [!NOTE]
> Этот метод не поддерживается в следующих сценариях.
> - В Outlook для iOS или Outlook для Android.
> - Когда надстройка загружается в почтовом ящике Gmail
> 
> Вместо этого надстройкам для доступа к почтовому ящику пользователя следует использовать [API-интерфейсы REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api).

Метод `makeEwsRequestAsync` отправляет к Exchange EWS-запрос от имени надстройки. Список поддерживаемых операций веб-служб Exchange см. в статье [Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).

С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.

XML-запрос должен указывать кодировку UTF-8.

```
<?xml version="1.0" encoding="utf-8"?>
```

У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> Чтобы включить метод  `makeEwsRequestAsync` для запросов служб EWS, администратор сервера должен установить для `OAuthAuthentication` значение true в каталоге сервера клиентского доступа EWS.

##### <a name="version-differences"></a>Различия версий

Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в версиях Outlook, предшествующих 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Запрос EWS.|
|`callback`| function||Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`. Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.|
|`userContext`| Object| &lt;необязательный&gt;|Любые данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
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
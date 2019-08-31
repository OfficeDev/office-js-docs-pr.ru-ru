---
title: Office. Context. Mailbox — набор обязательных элементов 1,4
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 66ae7cb05ac56224fd7461c5c29587e21a24020a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696213"
---
# <a name="mailbox"></a>mailbox

### <a name="officeofficemdcontextofficecontextmdmailbox"></a>[Office](Office.md)[.context](Office.context.md).mailbox

Предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [ewsUrl](#ewsurl-string) | Элемент |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | Метод |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Метод |
| [convertToRestId](#converttorestiditemid-restversion--string) | Метод |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Метод |
| [displayAppointmentForm](#displayappointmentformitemid) | Метод |
| [displayMessageForm](#displaymessageformitemid) | Метод |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | Метод |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Метод |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Метод |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Метод |

### <a name="namespaces"></a>Пространства имен

[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.

[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.

[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.

### <a name="members"></a>Элементы

#### <a name="ewsurl-string"></a>ewsUrl: строка

Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.

> [!NOTE]
> Этот элемент не поддерживается в Outlook на iOS или Android.

Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).

Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.

Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="methods"></a>Методы

#### <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

Преобразовывает идентификатор элемента из формата REST в формат EWS.

> [!NOTE]
> Этот метод не поддерживается в Outlook на iOS или Android.

Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате REST API для Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: String

##### <a name="example"></a>Пример

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}

Получает словарь, содержащий сведения о локальном времени клиента.

Почтовое приложение для Outlook на настольном компьютере или в Интернете может использовать разные часовые пояса для дат и времени. Outlook на рабочем столе использует часовой пояс клиентского компьютера; В Outlook в Интернете используется часовой пояс, установленный в центре администрирования Exchange. Значения даты и времени должны обрабатываться таким образом, чтобы значения, отображаемые в интерфейсе пользователя, всегда согласовывались с часовым поясом, ожидаемым пользователем.

Если почтовое приложение запущено в Outlook на настольном клиенте `convertToLocalClientTime` , метод возвратит объект Dictionary со значениями, заданными для часового пояса клиентского компьютера. Если почтовое приложение запущено в Outlook в Интернете, `convertToLocalClientTime` метод возвратит объект Dictionary со значениями, заданными в часовом поясе, заданном в центре администрирования Exchange.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`timeValue`| Date|Объект Date|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

Преобразовывает идентификатор элемента в формате EWS в формат REST.

> [!NOTE]
> Этот метод не поддерживается в Outlook на iOS или Android.

Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате EWS|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: String

##### <a name="example"></a>Пример

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

Получает объект Date из словаря, содержащего сведения о времени.

Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)|Значение локального времени для преобразования.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект Date со временем в формате UTC.

Тип: Date

##### <a name="example"></a>Пример

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

Отображает имеющуюся встречу из календаря.

> [!NOTE]
> Этот метод не поддерживается в Outlook на iOS или Android.

Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.

В Outlook на Mac Этот метод можно использовать для отображения одной встречи, которая не является частью повторяющегося ряда, или главной встречи повторяющейся серии, но невозможно отобразить экземпляр ряда. Это связано с тем, что в Outlook на Mac-адресе невозможно получить доступ к свойствам (включая идентификатор элемента) повторяющихся рядов.

В Outlook в Интернете этот метод открывает указанную форму, только если текст формы меньше или равен 32 КБ числу символов.

Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор веб-служб Exchange для существующей встречи в календаре.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

Отображает имеющееся сообщение.

> [!NOTE]
> Этот метод не поддерживается в Outlook на iOS или Android.

Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.

В Outlook в Интернете этот метод открывает указанную форму только в том случае, если размер текста формы меньше или равен 32 КБ.

Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.

Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор веб-служб Exchange для существующего сообщения.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

Отображает форму для создания новой встречи в календаре.

> [!NOTE]
> Этот метод не поддерживается в Outlook на iOS или Android.

Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.

В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.

Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.

Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
| `parameters` | Object | Словарь параметров, описывающий новую встречу. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей. |
| `parameters.start` | Date | Объект `Date`, указывающий дату и время начала встречи. |
| `parameters.end` | Date | Объект `Date`, указывающий дату и время окончания встречи. |
| `parameters.location` | Строка | Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255. |
| `parameters.resources` | Array.&lt;String&gt; | Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей. |
| `parameters.subject` | String | Строка с темой встречи. Максимальное количество символов в строке — 255. |
| `parameters.body` | String | Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```js
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

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.

Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.

Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).

Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.

Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Маркер указывается в виде строки в свойстве `asyncResult.value`.<br><br>При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.|
|`userContext`| Объект| &lt;необязательно&gt;|Данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`HTTPRequestFailure`|Запрос не выполнен. Просмотрите объект Diagnostics для кода ошибки HTTP.|
|`InternalServerError`|Сервер Exchange возвратил ошибку. Дополнительные сведения можно найти в объекте диагностики.|
|`NetworkError`|Пользователь больше не подключен к сети. Проверьте сетевое подключение и повторите попытку.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание и чтение|

##### <a name="example"></a>Пример

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

Получает маркер, идентифицирующий пользователя и надстройку Office.

Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Маркер указывается в виде строки в свойстве `asyncResult.value`.<br><br>При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.|
|`userContext`| Объект| &lt;необязательно&gt;|Данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`HTTPRequestFailure`|Запрос не выполнен. Просмотрите объект Diagnostics для кода ошибки HTTP.|
|`InternalServerError`|Сервер Exchange возвратил ошибку. Дополнительные сведения можно найти в объекте диагностики.|
|`NetworkError`|Пользователь больше не подключен к сети. Проверьте сетевое подключение и повторите попытку.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.

> [!NOTE]
> Этот метод не поддерживается в следующих сценариях:
> - В Outlook на iOS или Android
> - Если надстройка загружается в почтовый ящик Gmail.
> 
> В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.

Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange. Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).

С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.

В запросе XML должна быть указана кодировка UTF-8.

```xml
<?xml version="1.0" encoding="utf-8"?>
```

У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.

##### <a name="version-differences"></a>Различия версий

Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Запрос EWS.|
|`callback`| function||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`. Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.|
|`userContext`| Объект| &lt;необязательно&gt;|Данные о состоянии, передаваемые в асинхронный метод.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.

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

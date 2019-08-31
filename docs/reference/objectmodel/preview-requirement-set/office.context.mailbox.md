---
title: Office. Context. Mailbox — Предварительная версия набора обязательных элементов
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 951bb4ff338507f369f23e7c095debdb6e7a945e
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696479"
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
| [мастеркатегориес](#mastercategories-mastercategories) | Элемент |
| [restUrl](#resturl-string) | Элемент |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Метод |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | Метод |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Метод |
| [convertToRestId](#converttorestiditemid-restversion--string) | Метод |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Метод |
| [displayAppointmentForm](#displayappointmentformitemid) | Метод |
| [displayMessageForm](#displaymessageformitemid) | Метод |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | Метод |
| [дисплайневмессажеформ](#displaynewmessageformparameters) | Метод |
| [getCallbackTokenAsync](#getcallbacktokenasyncoptions-callback) | Метод |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Метод |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Метод |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Метод |
| [removeHandlerAsync](#removehandlerasynceventtype-options-callback) | Метод |

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

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a>Мастеркатегориес: [мастеркатегориес](/javascript/api/outlook/office.mastercategories)

Получает объект, предоставляющий методы для управления главным списком категорий в этом почтовом ящике.

> [!NOTE]
> Этот элемент не поддерживается в Outlook на iOS или Android.

##### <a name="type"></a>Тип

*   [MasterCategories](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| Предварительная версия |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox |
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение |

##### <a name="example"></a>Пример

В этом примере показано получение сводного списка категорий для этого почтового ящика.

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a>Рестурл: строка

Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.

С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.

Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.

Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="methods"></a>Методы

#### <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

Добавляет обработчик для поддерживаемого события.

В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.

##### <a name="parameters"></a>Параметры

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Событие, которое должно вызвать обработчик. |
| `handler` | Function || Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`. |
| `options` | Объект | &lt;Необязательно&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.asyncContext` | Object | &lt;Необязательно&gt; | Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова. |
| `callback` | функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

Преобразовывает идентификатор элемента из формата REST в формат EWS.

> [!NOTE]
> Этот метод не поддерживается в Outlook на iOS или Android.

Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате REST API для Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion)|Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.|

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

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}

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

Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)

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
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion)|Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.|

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
|`input`| [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)|Значение локального времени для преобразования.|

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

> [!NOTE]
> Все параметры являются необязательными.

|Имя| Тип| Описание|
|---|---|---|
| `parameters` | Object | Словарь параметров, описывающий новую встречу. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей. |
| `parameters.start` | Date | Объект `Date`, указывающий дату и время начала встречи. |
| `parameters.end` | Date | Объект `Date`, указывающий дату и время окончания встречи. |
| `parameters.location` | String. | Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255. |
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

#### <a name="displaynewmessageformparameters"></a>Дисплайневмессажеформ (Parameters)

Отображает форму для создания нового сообщения.

`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение. Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.

Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.

##### <a name="parameters"></a>Параметры

> [!NOTE]
> Все параметры являются необязательными.

|Имя| Тип| Описание|
|---|---|---|
| `parameters` | Object | Словарь параметров, описывающих новое сообщение. |
| `parameters.toRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому". Массив может включать не более 100 записей. |
| `parameters.ccRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия". Массив может включать не более 100 записей. |
| `parameters.bccRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК". Массив может включать не более 100 записей. |
| `parameters.subject` | String. | Строка, содержащая тему сообщения. Максимальное количество символов в строке — 255. |
| `parameters.htmlBody` | String | Текст сообщения в формате HTML. Максимальный размер содержимого сообщения — 32 КБ. |
| `parameters.attachments` | Array.&lt;Object&gt; | Массив объектов JSON, представляющих собой вложенные файлы или элементы. |
| `parameters.attachments.type` | String. | Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента. |
| `parameters.attachments.name` | Строка | Строка, содержащая имя вложения, длиной до 255 символов.|
| `parameters.attachments.url` | String | Используется, только если свойству `type` задано значение `file`. URI расположения файла. |
| `parameters.attachments.isInline` | Логический | Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений. |
| `parameters.attachments.itemId` | String | Используется, только если свойству `type` присвоено значение `item`. Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению. Это строка длиной до 100 символов. |


##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Чтение|

##### <a name="example"></a>Пример

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync([options], callback)

Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.

Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.

> [!NOTE]
> Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.

**Маркеры REST**

Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.

С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.

**Маркеры EWS**

Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.

С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.

##### <a name="parameters"></a>Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
| `options` | Object | &lt;Необязательно&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.isRest` | Boolean |  &lt;необязательно&gt; | Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`. |
| `options.asyncContext` | Объект |  &lt;необязательно&gt; | Данные о состоянии, передаваемые в асинхронный метод. |
|`callback`| function||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Маркер указывается в виде строки в свойстве `asyncResult.value`.<br><br>При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`HTTPRequestFailure`|Запрос не выполнен. Просмотрите объект Diagnostics для кода ошибки HTTP.|
|`InternalServerError`|Сервер Exchange возвратил ошибку. Дополнительные сведения можно найти в объекте диагностики.|
|`NetworkError`|Пользователь больше не подключен к сети. Проверьте сетевое подключение и повторите попытку.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание и чтение|

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
|`callback`| функция||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Маркер указывается в виде строки в свойстве `asyncResult.value`.<br><br>При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.|
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
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
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
|`callback`| функция||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Маркер указывается в виде строки в свойстве `asyncResult.value`.<br><br>При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.|
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

Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Вы можете определить, работает ли почтовое приложение в Outlook в Интернете или на настольном клиенте с помощью свойства Mailbox. Diagnostics. hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.

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

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a>removeHandlerAsync(eventType, [options], [callback])

Удаляет обработчиков для поддерживаемого типа события.

В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.

##### <a name="parameters"></a>Параметры

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Событие, которое должно отменить обработчик. |
| `options` | Объект | &lt;Необязательно&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.asyncContext` | Object | &lt;Необязательно&gt; | Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова. |
| `callback` | функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

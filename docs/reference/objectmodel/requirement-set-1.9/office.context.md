---
title: Office.context — набор требований 1.9
description: Office. Участники объектов Context, доступные для Outlook надстройки с помощью API почтовых ящиков, устанавливают 1.9.
ms.date: 12/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: 76371d3b19e87380581bbfeebd46a41eec79272f
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746241"
---
# <a name="context-mailbox-requirement-set-19"></a>контекст (набор требований к почтовым ящикам 1.9)

### <a name="officecontext"></a>[Office](office.md).context

Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях. Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

## <a name="properties"></a>Свойства

| Свойство | Режимы | Тип возвращаемых данных | Minimum<br>набор требований |
|---|---|---|:---:|
| [auth](#auth-auth) | Создание<br>Чтение | [Auth](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [IdentityAPI 1.3](../../requirement-sets/identity-api-requirement-sets.md) |
| [contentLanguage](#contentlanguage-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [диагностика](#diagnostics-contextinformation) | Создание<br>Чтение | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [хост](#host-hosttype) | Создание<br>Чтение | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [mailbox](office.context.mailbox.md) | Создание<br>Чтение | [Mailbox](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [платформа](#platform-platformtype) | Создание<br>Чтение | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [требования](#requirements-requirementsetsupport) | Создание<br>Чтение | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Создание<br>Чтение | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Создание<br>Чтение | [UI](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>Сведения о свойстве

#### <a name="auth-auth"></a>Auth: [Auth](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true)

Поддерживает один вход [(SSO),](../../../outlook/authenticate-a-user-with-an-sso-token.md) предоставляя метод, который позволяет Office приложению получить маркер доступа к веб-приложению надстройки. Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему. См. набор требований [IdentityAPI 1.3](../../requirement-sets/identity-api-requirement-sets.md).

##### <a name="type"></a>Type

*   [Auth](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](../../requirement-sets/outlook-api-requirement-sets.md)| Недоступно|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a>contentLanguage: String

Получает локализ (язык), указанный пользователем для редактирования элемента.

Это `contentLanguage` значение отражает текущий параметр **Язык** редактирования, указанный в файле > **Параметры > язык** в клиентском приложении Office.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a>диагностика: [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true)

Получает сведения об среде, в которой работает надстройка.

##### <a name="type"></a>Type

*   [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage: String

Получает локализ (язык) в формате языковых тегов RFC 1766, заданный пользователем для пользовательского интерфейса Office клиентского приложения.

Это `displayLanguage` значение отражает текущий параметр **Язык** отображения, указанный в файле > **Параметры > язык** в клиентском приложении Office.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a>host: [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true)

Получает Office приложение, в которое размещена надстройка.

> [!NOTE]
> Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics](#diagnostics-contextinformation).

##### <a name="type"></a>Type

*   [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a>платформа: [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true)

Предоставляет платформу, на которой запущена надстройка.

> [!NOTE]
> Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics](#diagnostics-contextinformation).

##### <a name="type"></a>Type

*   [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a>требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true)

Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.

##### <a name="type"></a>Type

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true)

Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.

`RoamingSettings` Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому почтовому ящику.

##### <a name="type"></a>Type

*   [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)| С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="ui-ui"></a>ui: [пользовательский интерфейс](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true)

Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.

##### <a name="type"></a>Type

*   [UI](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

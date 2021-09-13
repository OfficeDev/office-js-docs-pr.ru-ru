---
title: Office.context — предварительная версия набора обязательных элементов
description: Office. Участники объектов Context, доступные для Outlook надстройки с помощью набора требований к API API почтовых ящиков.
ms.date: 12/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: f09e84062bc12f3d69adbbbedd6d02679362a728
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151629"
---
# <a name="context-mailbox-preview-requirement-set"></a>контекст (набор требований предварительного просмотра почтовых ящиков)

### <a name="officecontext"></a>[Office](office.md).context

Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях. Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API.](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

## <a name="properties"></a>Свойства

| Свойство | Режимы | Тип возвращаемых данных | Minimum<br>набор требований |
|---|---|---|:---:|
| [auth](#auth-auth) | Создание<br>Чтение | [Auth](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [IdentityAPI 1.3](../../requirement-sets/identity-api-requirement-sets.md) |
| [contentLanguage](#contentlanguage-string) | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [диагностика](#diagnostics-contextinformation) | Создание<br>Чтение | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [хост](#host-hosttype) | Создание<br>Чтение | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [mailbox](office.context.mailbox.md) | Создание<br>Чтение | [Mailbox](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [officeTheme](#officetheme-officetheme) | Создание<br>Чтение | [OfficeTheme](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [Предварительная версия](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [платформа](#platform-platformtype) | Создание<br>Чтение | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [требования](#requirements-requirementsetsupport) | Создание<br>Чтение | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Создание<br>Чтение | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Создание<br>Чтение | [UI](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>Сведения о свойстве

#### <a name="auth-auth"></a>Auth: [Auth](/javascript/api/office/office.auth)

Поддерживает один вход [(SSO),](../../../outlook/authenticate-a-user-with-an-sso-token.md) предоставляя метод, который позволяет Office приложению получить маркер доступа к веб-приложению надстройки. Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.

##### <a name="type"></a>Тип

*   [Auth](/javascript/api/office/office.auth)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](../../requirement-sets/outlook-api-requirement-sets.md)| Предварительная версия|
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

Это значение отражает текущий параметр Язык редактирования, указанный в файле > `contentLanguage` **Параметры > язык** в клиентском приложении Office. 

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

#### <a name="diagnostics-contextinformation"></a>диагностика: [ContextInformation](/javascript/api/office/office.contextinformation)

Получает сведения об среде, в которой работает надстройка.

##### <a name="type"></a>Тип

*   [ContextInformation](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>Требования

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

Это значение отражает текущий параметр Язык отображения, указанный в файле > `displayLanguage` **Параметры > язык** в клиентском приложении Office. 

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

#### <a name="host-hosttype"></a>host: [HostType](/javascript/api/office/office.hosttype)

Получает Office приложение, в которое размещена надстройка.

> [!NOTE]
> Кроме того, для получения хоста можно использовать [свойство Office.context.diagnostics.](#diagnostics-contextinformation)

##### <a name="type"></a>Тип

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>Требования

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

#### <a name="officetheme-officetheme"></a>officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)

Предоставляет доступ к свойствам цветов темы Office.

> [!NOTE]
> Этот член поддерживается только в Outlook на Windows.

Использование Office тем позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с пользовательским интерфейсом **File > Office Account > Office Theme,** который применяется во всех Office клиентских приложениях. Using Office theme colors is appropriate for mail and task pane add-ins.

##### <a name="type"></a>Тип

*   [OfficeTheme](/javascript/api/office/office.officetheme)

##### <a name="properties"></a>Свойства

|Имя| Тип| Описание|
|---|---|---|
|`bodyBackgroundColor`| String|Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.|
|`bodyForegroundColor`| String|Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.|
|`controlBackgroundColor`| String|Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.|
|`controlForegroundColor`| String|Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](../../requirement-sets/outlook-api-requirement-sets.md)| Предварительная версия|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtype"></a>платформа: [PlatformType](/javascript/api/office/office.platformtype)

Предоставляет платформу, на которой запущена надстройка.

> [!NOTE]
> Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics.](#diagnostics-contextinformation)

##### <a name="type"></a>Тип

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>Требования

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

#### <a name="requirements-requirementsetsupport"></a>требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.

##### <a name="type"></a>Тип

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a>Требования

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

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)

Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.

Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому `RoamingSettings` почтовому ящику.

##### <a name="type"></a>Тип

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)| С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="ui-ui"></a>ui: [пользовательский интерфейс](/javascript/api/office/office.ui)

Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.

##### <a name="type"></a>Тип

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

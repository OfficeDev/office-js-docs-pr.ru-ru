---
title: Office.context — предварительная версия набора обязательных элементов
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 5c34a7a0db5880a94ba5519059a93010a5243978
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629190"
---
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="properties"></a>Свойства

| Свойство | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|---|
| [contentLanguage](#contentlanguage-string) | Создание<br>Чтение | String | 1.0 |
| [diagnostics](#diagnostics-contextinformation) | Создание<br>Чтение | [контекстинформатион](/javascript/api/office/office.contextinformation) | 1.0 |
| [displayLanguage](#displaylanguage-string) | Создание<br>Чтение | String | 1.0 |
| [размещать](#host-hosttype) | Создание<br>Чтение | [HostType](/javascript/api/office/office.hosttype) | 1.0 |
| [officeTheme](#officetheme-officetheme) | Создание<br>Чтение | [OfficeTheme](/javascript/api/office/office.officetheme) | Предварительная версия |
| [управляем](#platform-platformtype) | Создание<br>Чтение | [PlatformType](/javascript/api/office/office.platformtype) | 1.0 |
| [потребность](#requirements-requirementsetsupport) | Создание<br>Чтение | [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport) | 1.0 |
| [roamingSettings](#roamingsettings-roamingsettings) | Создание<br>Чтение | [RoamingSettings](/javascript/api/outlook/office.roamingsettings) | 1.0 |
| [ui](#ui-ui) | Создание<br>Чтение | [UI](/javascript/api/office/office.ui) | 1.0 |

### <a name="namespaces"></a>Пространства имен

[Проверка подлинности](/javascript/api/office/office.auth): предоставляет поддержку [единого входа (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).

[почтовый ящик](office.context.mailbox.md): предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.

## <a name="property-details"></a>Сведения о свойстве

#### <a name="contentlanguage-string"></a>contentLanguage: строка

Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.

`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a>Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)

Получает сведения о среде, в которой выполняется надстройка.

##### <a name="type"></a>Тип

*   [контекстинформатион](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage: строка

Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.

The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a>узел: [HostType](/javascript/api/office/office.hosttype)

Получает узел приложений Office, в котором работает надстройка.

##### <a name="type"></a>Тип

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a>officeTheme: [officeTheme](/javascript/api/office/office.officetheme)

Предоставляет доступ к свойствам цветов темы Office.

> [!NOTE]
> Этот элемент поддерживается только в Outlook для Windows.

Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы**Office, которая применяется ко всем ведущим приложениям Office. Using Office theme colors is appropriate for mail and task pane add-ins.

##### <a name="type"></a>Тип

*   [OfficeTheme](/javascript/api/office/office.officetheme)

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`bodyBackgroundColor`| String|Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.|
|`bodyForegroundColor`| String|Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.|
|`controlBackgroundColor`| String|Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.|
|`controlForegroundColor`| String|Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.|

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| Предварительная версия|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

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

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a>Платформа: [PlatformType](/javascript/api/office/office.platformtype)

Предоставляет платформу, на которой работает надстройка.

##### <a name="type"></a>Тип

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a>требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)

Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.

##### <a name="type"></a>Тип

*   [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a>roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)

Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.

Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.

##### <a name="type"></a>Тип

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a>Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)

Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.

##### <a name="type"></a>Тип

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

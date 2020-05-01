---
title: Office. Context — набор обязательных элементов 1,7
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,7.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 357f8f6923730eeddeb9379279bb90fbbc64143b
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891273"
---
# <a name="context-mailbox-requirement-set-17"></a>контекст (набор требований для почтового ящика 1,7)

### <a name="officecontext"></a>[Office](office.md).context

Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office. В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.7).

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="properties"></a>Свойства

| Свойство | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|:---:|
| [contentLanguage](#contentlanguage-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [diagnostics](#diagnostics-contextinformation) | Создание<br>Чтение | [контекстинформатион](/javascript/api/office/office.contextinformation?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [размещать](#host-hosttype) | Создание<br>Чтение | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [mailbox](office.context.mailbox.md) | Создание<br>Чтение | [Mailbox](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [управляем](#platform-platformtype) | Создание<br>Чтение | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [потребность](#requirements-requirementsetsupport) | Создание<br>Чтение | [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Создание<br>Чтение | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Создание<br>Чтение | [UI](/javascript/api/office/office.ui?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>Сведения о свойстве

#### <a name="contentlanguage-string"></a>contentLanguage: строка

Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.

`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.

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

#### <a name="diagnostics-contextinformation"></a>Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)

Получает сведения о среде, в которой выполняется надстройка.

##### <a name="type"></a>Type

*   [контекстинформатион](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

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

#### <a name="host-hosttype"></a>узел: [HostType](/javascript/api/office/office.hosttype)

Получает узел приложений Office, в котором работает надстройка.

##### <a name="type"></a>Type

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a>Платформа: [PlatformType](/javascript/api/office/office.platformtype)

Предоставляет платформу, на которой работает надстройка.

##### <a name="type"></a>Type

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a>требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)

Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.

##### <a name="type"></a>Type

*   [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)

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

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)

Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.

Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.

##### <a name="type"></a>Type

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../outlook/understanding-outlook-add-in-permissions.md)| С ограничениями|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="ui-ui"></a>Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)

Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.

##### <a name="type"></a>Type

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

---
title: Office. Context — набор обязательных элементов 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c7c33f31a63fc2a57b9457fc788a80c925d3eef1
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450291"
---
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).


##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="namespaces"></a>Пространства имен

[mailbox](office.context.mailbox.md). Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.

### <a name="members"></a>Элементы

####  <a name="displaylanguage-string"></a>displayLanguage :String

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

```javascript
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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook11officeroamingsettings"></a>roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)

Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.

Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.

##### <a name="type"></a>Type

*   [RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| С ограничениями|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

---
title: Office. Context. Mailbox. userProfile — Предварительная версия набора требований
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 29111314f16bb9c6518b350254a3036ffa125796
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451964"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [accountType](#accounttype-string) | Member |
| [displayName](#displayname-string) | Member |
| [emailAddress](#emailaddress-string) | Member |
| [timeZone](#timezone-string) | Member |

### <a name="members"></a>Members

####  <a name="accounttype-string"></a>accountType: строка

> [!NOTE]
> В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии для Mac (сборка 16.9.1212 или более поздняя).

Возвращает тип учетной записи пользователя, связанного с почтовым ящиком. Возможные значения перечислены в приведенной ниже таблице.

| Значение | Описание |
|-------|-------------|
| `enterprise` | Почтовый ящик находится на локальном сервере Exchange. |
| `gmail` | Почтовый ящик связан с учетной записью Gmail. |
| `office365` | Почтовый ящик связан с рабочей или учебной учетной записью Office 365. |
| `outlookCom` | Почтовый ящик связан с личной учетной записью Outlook.com. |

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a>displayName :String

Получает отображаемое имя пользователя.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a>emailAddress :String

Получает адрес электронной почты SMTP пользователя.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a>timeZone :String

Получает часовой пояс пользователя по умолчанию.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```

---
title: Office. Context. Mailbox. Diagnostics — Предварительная версия набора требований
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 86ee4093dbfb35a5306938da27b61eb6e8936792
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127578"
---
# <a name="diagnostics"></a>diagnostics

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics

Предоставляет надстройке Outlook диагностические сведения.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [Сайту](#hostname-string) | Member |
| [hostVersion](#hostversion-string) | Member |
| [OWAView](#owaview-string) | Member |

### <a name="members"></a>Members

#### <a name="hostname-string"></a>Имя узла: строка

Получает строку, представляющую имя ведущего приложения.

Строка, которая может иметь одно из следующих значений: `Outlook`, `Mac Outlook`, `OutlookIOS` или `OutlookWebApp`.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

---
---

#### <a name="hostversion-string"></a>hostVersion: строка

Получает строку, которая представляет версию ведущего приложения или Exchange Server.

Если почтовая надстройка работает в клиенте Outlook для настольного компьютера или в iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook. В Outlook в Интернете свойство возвращает версию сервера Exchange. Пример — строка `15.0.468.0`.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

---
---

#### <a name="owaview-string"></a>OWAView: строка

Получает строку, представляющую текущее представление Outlook в Интернете.

Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.

Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.

В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:

*   `OneColumn`, который отображается, когда экран сужается. В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.
*   `TwoColumns`, который отображается, когда экран расширяется. Outlook в Интернете использует это представление на большинстве планшетов.
*   `ThreeColumns` используется для полноразмерных экранов. Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

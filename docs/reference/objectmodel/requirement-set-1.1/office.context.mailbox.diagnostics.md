---
title: Office.context.mailbox.diagnostics — набор обязательных элементов 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8ce3242a9e007e9aa686ddd098ab23ee922276e4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451999"
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

### <a name="members"></a>Элементы

####  <a name="hostname-string"></a>hostName :String

Получает строку, представляющую имя ведущего приложения.

Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

####  <a name="hostversion-string"></a>hostVersion :String

Получает строку, которая представляет версию ведущего приложения или Exchange Server.

Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения, Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Пример — строка `15.0.468.0`.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

####  <a name="owaview-string"></a>OWAView :String

Получает строку, отображающую текущее представление Outlook Web App.

Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.

Если Outlook Web App — не ведущее приложение, при получении доступа к этому свойству будет выдаваться значение `undefined`.

Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов.

*   `OneColumn` используется в случае узкого экрана: Outlook Web App использует этот макет размером в один столбец на экране смартфона.
*   `TwoColumns` используется при более широком экране: Outlook Web App использует это представление на большинстве планшетных ПК.
*   `ThreeColumns` используется для полноразмерных экранов. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.

##### <a name="type"></a>Тип

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

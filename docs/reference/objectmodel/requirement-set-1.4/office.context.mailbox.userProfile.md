---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0532a9971a05412d37334f4c5a4b6b12654f61f3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950994"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile

Предоставляет сведения о пользователе в надстройке Outlook.

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

## <a name="properties"></a>Свойства

| Свойство | Минимальные<br>уровень разрешения | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|---|:---:|
| [displayName](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | ReadItem | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [emailAddress](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | ReadItem | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | ReadItem | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

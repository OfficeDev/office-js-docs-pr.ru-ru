---
title: Включен элемент в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937302"
---
# <a name="enabled-element"></a>Элемент Включен

Указывает, включено ли управление [кнопкой](control.md#button-control) или [меню](control.md#menu-dropdown-button-controls) при запуске надстройки. Элемент **Включен** — это детский элемент [Управления.](control.md) Если он опущен, по умолчанию `true` .

Этот элемент действителен только в Excel; то есть, когда `Name` атрибутом элемента [Host](host.md) является "Книга".

Родительский контроль также может быть включен программным образом и отключен. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```

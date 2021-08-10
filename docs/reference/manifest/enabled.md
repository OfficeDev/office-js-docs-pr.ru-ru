---
title: Включен элемент в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 54d28839a274ff41bab0b1e2cdd2d169e76c5815095950dec67ce2564eade601
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093909"
---
# <a name="enabled-element"></a>Элемент Включен

Указывает, включено ли управление [кнопкой](control.md#button-control) или [меню](control.md#menu-dropdown-button-controls) при запуске надстройки. Элемент **Включен** — это детский элемент [Управления.](control.md) Если он опущен, по умолчанию `true` .

Этот элемент действителен только в Excel; то есть, когда `Name` атрибутом элемента [Host](host.md) является "Книга".

Родительский контроль также может быть включен программным образом и отключен. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```

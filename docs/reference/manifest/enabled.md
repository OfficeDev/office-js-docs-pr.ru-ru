---
title: Включен элемент в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: a14385f7114eb3d35845b5d9873bdd718b46c0e9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154002"
---
# <a name="enabled-element"></a>Элемент Включен

Указывает, включено ли управление [кнопкой](control.md#button-control) или [меню](control.md#menu-dropdown-button-controls) при запуске надстройки. Элемент **Включен** — это детский элемент [Управления.](control.md) Если он опущен, по умолчанию `true` .

Этот элемент действителен только в Excel; то есть, когда `Name` атрибутом элемента [Host](host.md) является "Книга".

Родительский контроль также может быть включен программным образом и отключен. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```

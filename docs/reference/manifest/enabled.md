---
title: Элемент Enabled в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771399"
---
# <a name="enabled-element"></a>Элемент Enabled

Указывает, включен ли [](control.md#menu-dropdown-button-controls) [при](control.md#button-control) запуске надстройки пункт "Кнопка" или "Меню". Элемент **Enabled** — это элемент [control.](control.md) Если он опущен, значение по умолчанию `true` : .

Этот элемент действителен только в Excel; то есть, если `Name` атрибутом элемента [Host](host.md) является "Workbook".

Родительский контроль также можно включить программным образом и отключить. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```

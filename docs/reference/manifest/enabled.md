---
title: Элемент Enabled в файле манифеста
description: Сведения о том, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 4c2c013c8e55966ba2678755536ce04ae3014ed0
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596902"
---
# <a name="enabled-element"></a>Элемент Enabled

Указывает, включен ли элемент управления ["Кнопка" или "](control.md#button-control) [меню](control.md#menu-dropdown-button-controls) " при запуске надстройки. Элемент **Enabled** является дочерним элементом [элемента управления](control.md). Если он не указан, используется `true`значение по умолчанию.

Родительский элемент управления также может быть включен и отключен программным способом. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```

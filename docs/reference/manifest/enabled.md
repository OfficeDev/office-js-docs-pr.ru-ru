---
title: Элемент Enabled в файле манифеста
description: Сведения о том, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566215"
---
# <a name="enabled-element"></a>Элемент Enabled

Указывает, включен ли элемент управления ["Кнопка" или "](control.md#button-control) [меню](control.md#menu-dropdown-button-controls) " при запуске надстройки. Элемент **Enabled** является дочерним элементом [элемента управления](control.md). Если он не указан, используется `true`значение по умолчанию. 

Родительский элемент управления также может быть включен и отключен программным способом. Дополнительные сведения можно найти [в статье Включение и отключение команд надстроек](/office/dev/add-ins/design/disable-add-in-commands).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```


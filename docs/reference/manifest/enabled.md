---
title: Включен элемент в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 75a2664143e29c86a05aaf039b0ea7bce659cef9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744758"
---
# <a name="enabled-element"></a>Элемент Включен

Указывает, включено ли управление [кнопкой](control-button.md) или [меню](control-menu.md) при запуске надстройки. Элемент **Включен** — это детский элемент [Управления](control.md). Если он опущен, по умолчанию .`true`

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

Этот элемент действителен только в Excel, PowerPoint и Word; то есть, `Name` когда атрибутом элемента [Host](host.md) являются "Книга", "Презентация" или "Документ".

Родительский контроль также может быть включен программным образом и отключен. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```

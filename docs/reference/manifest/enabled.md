---
title: Включен элемент в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: a3d83a6d117c498cc4d54dfbe73ae6d800995cb6
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467852"
---
# <a name="enabled-element"></a>Элемент Включен

Указывает, включено ли управление [кнопкой](control-button.md) или [меню](control-menu.md) при запуске надстройки. Элемент **Включен** — это детский элемент [Управления](control.md). Если он опущен, по умолчанию .`true`

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

Этот элемент действителен только в Excel, то есть, `Name` когда атрибутом элемента [Host](host.md) является "Книга".

Родительский контроль также может быть включен программным образом и отключен. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```

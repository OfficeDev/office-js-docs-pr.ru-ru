---
title: Включен элемент в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4c0107daaf73aee6ba116553a8d01250e9c7d981
ms.sourcegitcommit: 997a20f9fb011b96a50ceb04a4b9943d92d6ecf4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/19/2021
ms.locfileid: "61081437"
---
# <a name="enabled-element"></a>Элемент Включен

Указывает, включено ли управление [кнопкой](control.md#button-control) или [меню](control.md#menu-dropdown-button-controls) при запуске надстройки. Элемент **Включен** — это детский элемент [Управления.](control.md) Если он опущен, по умолчанию `true` .

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides:**

- Область задач 1.0

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

Этот элемент действителен только в Excel, то есть, когда атрибутом элемента `Name` [Host](host.md) является "Книга".

Родительский контроль также может быть включен программным образом и отключен. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```

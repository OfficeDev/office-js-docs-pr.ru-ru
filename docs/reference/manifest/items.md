---
title: Элемент Items в файле манифеста
description: Указывает элементы в меню.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2249bc55db662a36cf3986ebb0b90353237d4985
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467928"
---
# <a name="items-element"></a>Элемент Items

Указывает элементы в меню.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) , когда родительский **VersionOverrides** — это тип Taskpane 1.0.
- [Почтовый ящик 1.3,](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) когда родительский **VersionOverrides** — это тип Почта 1.0.
- [Почтовый ящик 1.5,](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) когда родительский **VersionOverrides** — это тип Почта 1.1.

## <a name="syntax"></a>Синтаксис

```XML
<Items>
...  
</Items>  
```

## <a name="contained-in"></a>Содержится в

[Элемент управления меню типа](control-menu.md)

## <a name="must-contain"></a>Должен содержать

[Элемент](item.md)

## <a name="examples"></a>Примеры

Примеры см. в [пункте Управление типом Меню](control-menu.md).
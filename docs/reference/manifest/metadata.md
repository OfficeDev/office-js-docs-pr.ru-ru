---
title: Элемент Metadata в файле манифеста
description: Элемент Metadata определяет параметры метаданных, которые настраиваемая функция использует в Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52938155442bb5424a170634d1324de77de2b788
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855536"
---
# <a name="metadata-element"></a>Элемент Metadata

Определяет параметры метаданных, используемые пользовательской функцией в Excel.

**Тип надстройки:** Настраиваемая функция

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|  Элемент  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса JSON-файла, используемого пользовательскими функциями. |

## <a name="example"></a>Пример

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```

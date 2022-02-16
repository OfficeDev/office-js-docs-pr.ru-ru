---
title: Элемент Script в файле манифеста
description: Элемент Script определяет параметры скрипта, которые настраиваемая функция использует в Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0f32314912dd66d8578750bf4818af8483c8ef36
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855543"
---
# <a name="script-element"></a>Элемент Script

Определяет параметры сценариев, используемых пользовательской функцией в Excel.

**Тип надстройки:** Настраиваемая функция

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|Элементы  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса файла JavaScript, используемого пользовательскими функциями.|

## <a name="example"></a>Пример

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```

---
title: Элемент AllFormFactors в файле манифеста
description: Указывает параметры всех форм-факторов для надстройки.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa15eb48ec8d3fde125973efcea36067f7cdac39
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340409"
---
# <a name="allformfactors-element"></a>Элемент AllFormFactors

Указывает параметры всех форм-факторов для надстройки. В настоящее время пользовательская функция — единственная, где применяется **AllFormFactors**. Элемент **AllFormFactors** является обязательным при использовании пользовательских функций.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

> [!NOTE]
> Этот элемент поддерживается только в Excel Windows, на Mac и в Интернете. Он не поддерживается в других Office приложениях или на iOS или Android.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Да |  Определяет, где предоставляются функции надстройки. |

## <a name="allformfactors-example"></a>Пример использования AllFormFactors

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```

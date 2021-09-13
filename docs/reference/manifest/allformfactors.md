---
title: Элемент AllFormFactors в файле манифеста
description: Указывает параметры всех форм-факторов для надстройки.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: b579612d73216fab6141501e1c969fb6be1e8495
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153733"
---
# <a name="allformfactors-element"></a>Элемент AllFormFactors

Указывает параметры всех форм-факторов для надстройки. В настоящее время пользовательская функция — единственная, где применяется **AllFormFactors**. Элемент **AllFormFactors** является обязательным при использовании пользовательских функций.

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

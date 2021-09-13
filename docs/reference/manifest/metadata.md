---
title: Элемент Metadata в файле манифеста
description: Элемент Metadata определяет параметры метаданных, которые настраиваемая функция использует в Excel.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6f58b00bb13bde1e2b1742462716119b8b6d369d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153878"
---
# <a name="metadata-element"></a>Элемент Metadata

Определяет параметры метаданных, используемые пользовательской функцией в Excel.

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

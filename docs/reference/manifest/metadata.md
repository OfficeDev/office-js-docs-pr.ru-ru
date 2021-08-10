---
title: Элемент Metadata в файле манифеста
description: Элемент Metadata определяет параметры метаданных, которые настраиваемая функция использует в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d6b7af8988265baf8fbdea51e1414646f88868ede76ed7194c47db1eb874608d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092493"
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

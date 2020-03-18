---
title: Элемент Metadata в файле манифеста
description: Элемент Metadata определяет параметры метаданных, используемые пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8ea81818aa96b407ce386ec318495ec5ba773d05
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718071"
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

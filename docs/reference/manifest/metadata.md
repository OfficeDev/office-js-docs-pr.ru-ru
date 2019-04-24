---
title: Элемент Metadata в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a3aecb1983905658f3a55fdb8bf0629a8d5ef474
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452048"
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

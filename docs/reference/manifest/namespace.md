---
title: Элемент Namespace в файле манифеста
description: Элемент Namespace определяет пространство имен, которое настраиваемая функция использует в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938996"
---
# <a name="namespace-element"></a>Элемент Namespace

Определяет пространство имен, используемых пользовательской функцией в Excel.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Нет  | Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md). Может быть не более 32 символов. |

## <a name="child-elements"></a>Дочерние элементы

Нет

## <a name="example"></a>Пример

```xml
<Namespace resid="namespace" />
```

---
title: Элемент Namespace в файле манифеста
description: Элемент Namespace определяет пространство имен, которое настраиваемая функция использует в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 3f20e744839d5791797642a9019f546922efd710367d5f23446241eebad0e48f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089682"
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

---
title: Элемент Namespace в файле манифеста
description: Элемент Namespace определяет пространство имен, которое пользовательская функция использует в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771263"
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

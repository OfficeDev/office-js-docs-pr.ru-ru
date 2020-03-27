---
title: Элемент Namespace в файле манифеста
description: Элемент namespace определяет пространство имен, используемое пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eabd73d3be98271c81723787dd3d1bdb6ee2ebcd
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978671"
---
# <a name="namespace-element"></a>Элемент Namespace

Определяет пространство имен, используемых пользовательской функцией в Excel.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Нет  | Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md). |

## <a name="child-elements"></a>Дочерние элементы

Нет

## <a name="example"></a>Пример

```xml
<Namespace resid="namespace" />
```

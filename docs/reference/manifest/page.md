---
title: Элемент Page в файле манифеста
description: Элемент Page определяет параметры страницы HTML, используемые пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0c56b955b79f9052ee2c89a391dd95b2975d69c2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720486"
---
# <a name="page-element"></a>Элемент Page

Определяет параметры HTML-страницы, используемые пользовательской функцией в Excel.

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|  Элемент  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса HTML-файла, используемого пользовательскими функциями. |

## <a name="example"></a>Пример

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```

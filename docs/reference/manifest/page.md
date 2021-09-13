---
title: Элемент Page в файле манифеста
description: Элемент Page определяет параметры HTML-страниц, которые настраиваемая функция использует в Excel.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6bde3ba86270874b1d9059b2f1c44952241bf00f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154860"
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

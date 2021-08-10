---
title: Элемент Page в файле манифеста
description: Элемент Page определяет параметры HTML-страниц, которые настраиваемая функция использует в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0d10ed04b73dfd786d50150dd8d01629f826cde17cb5ed1a7633d7319b5d6490
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57090168"
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

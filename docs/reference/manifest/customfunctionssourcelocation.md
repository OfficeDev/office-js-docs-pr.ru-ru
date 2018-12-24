---
title: Элемент SourceLocation в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 878a8184984e31fdbcf46192a2f56507edaf4b37
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432410"
---
# <a name="sourcelocation-element"></a>Элемент SourceLocation

Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.

## <a name="attributes"></a>Атрибуты

| **Атрибут** | **Обязательный** | **Описание**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | Да          | Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте. |

## <a name="child-elements"></a>Дочерние элементы

Нет

## <a name="example"></a>Пример

```xml
<SourceLocation resid="pageURL"/>
```
---
title: Элемент SourceLocation в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450690"
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

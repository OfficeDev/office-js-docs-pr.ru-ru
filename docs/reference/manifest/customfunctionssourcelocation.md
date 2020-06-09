---
title: Элемент SourceLocation в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 88ae0558577167074a870170833617c4f60730f1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612314"
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

---
title: Элемент SourceLocation для настраиваемой функции в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 6001673f1954a4af2de66ff7611069c3fb402a13
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939285"
---
# <a name="sourcelocation-element-custom-functions"></a>Элемент SourceLocation (настраиваемые функции)

Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.

## <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Да      | Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте. Может быть не более 32 символов. |

## <a name="child-elements"></a>Дочерние элементы

Нет

## <a name="example"></a>Пример

```xml
<SourceLocation resid="pageURL"/>
```

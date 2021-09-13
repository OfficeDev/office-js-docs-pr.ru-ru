---
title: Элемент SourceLocation для настраиваемой функции в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 08/07/2020
ms.localizationpriority: medium
ms.openlocfilehash: 84d5607fbb02c1925137e1a143b7715c7c87c6fa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151665"
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

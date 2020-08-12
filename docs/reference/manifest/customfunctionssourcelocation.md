---
title: Элемент SourceLocation для пользовательских функций в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 1c509987b0ce7948a63fa8ad51f7cf9c84144c5f
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641384"
---
# <a name="sourcelocation-element-custom-functions"></a>Элемент SourceLocation (пользовательские функции)

Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.

## <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Да      | Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте. |

## <a name="child-elements"></a>Дочерние элементы

Нет

## <a name="example"></a>Пример

```xml
<SourceLocation resid="pageURL"/>
```

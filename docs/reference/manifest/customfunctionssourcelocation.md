---
title: Элемент SourceLocation для настраиваемой функции в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5f2d881f31f4e46e7f5bb8ab30d78abd0e9b7200
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990686"
---
# <a name="sourcelocation-element-custom-functions"></a>Элемент SourceLocation (настраиваемые функции)

Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.

**Тип надстройки:** Настраиваемая функция

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

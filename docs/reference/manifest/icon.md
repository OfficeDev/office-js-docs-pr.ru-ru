---
title: Элемент Icon в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45f3dcda8e74430cf70aa765efc6b3aae0e2b448
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450620"
---
# <a name="icon-element"></a>Элемент Icon

Определяет элементы **Image** для элементов управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Нет  | Тип определяемого значка. Относится только к значкам в форм-факторах мобильных устройств. Для элементов **Icon**, содержащихся в элементе [MobileFormFactor](mobileformfactor.md), этому атрибуту присвоено значение `bt:MobileIconList`. |

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Image](#image)        | Да |   атрибут resid используемого изображения         |

### <a name="image"></a>Изображение

Изображение кнопки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **Image** в элементе **Images** в элементе [Resources](resources.md). Атрибут **size** указывает размер изображения в пикселях. Обязательными являются три размера изображения (16, 32 и 80 пикселей), а поддерживаются еще пять (20, 24, 40, 48 и 64 пикселя).|

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a>Дополнительные требования для форм-факторов мобильных устройств

Когда родительский элемент **Icon** является потомком элемента [MobileFormFactor](mobileformfactor.md), минимальные требуемые размеры несколько отличаются. В манифесте должны быть указаны размеры, составляющие по крайней мере 48 x 48, 32 x 32 и 25 x 25 пикселей. Каждый указанный размер должен встречаться три раза, при этом атрибуту `scale` должно быть присвоено значение `1`, `2` или `3`.

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```

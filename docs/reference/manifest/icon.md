---
title: Элемент Icon в файле манифеста
description: Определяет элементы Image для элементов управления Button или Menu.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: ddacd3fdb8b9306e9f26b7bf2431f3564b579264
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467731"
---
# <a name="icon-element"></a>Элемент Icon

Определяет **элементы изображения** для элементов [управления кнопками и](control-button.md) [элементов управления меню](control-menu.md).

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Нет  | Тип определяемого значка. Относится только к значкам в форм-факторах мобильных устройств. Для элементов **Icon**, содержащихся в элементе [MobileFormFactor](mobileformfactor.md), этому атрибуту присвоено значение `bt:MobileIconList`. |

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Image](#image)        | Да |   атрибут resid используемого изображения         |

### <a name="image"></a>Изображение

Изображение кнопки. Атрибут **resid** может быть не более 32 символов и должен быть задатки значению атрибута **id** элемента **Image** в элементе **Images** в [элементе Resources](resources.md) . Атрибут **size** указывает размер изображения в пикселях. Необходимо использовать три размера изображения (16, 32 и 80 пикселей), а поддерживается еще пять размеров (20, 24, 40, 48 и 64 пикселя).

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

> [!IMPORTANT]
> Если это изображение является символом представительства надстройки, см. в этой записи Создание эффективных списков в [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) и в Office для размера и других требований.

## <a name="additional-requirements-for-mobile-form-factors"></a>Дополнительные требования для форм-факторов мобильных устройств

Когда родительский элемент **Icon** является потомком элемента [MobileFormFactor](mobileformfactor.md), минимальные требуемые размеры несколько отличаются. В манифесте должны быть указаны размеры, составляющие по крайней мере 48 x 48, 32 x 32 и 25 x 25 пикселей. Каждый указанный размер должен встречаться три раза, при этом атрибуту `scale` должно быть присвоено значение `1`, `2` или `3`. Этот атрибут указывает свойство `UIScreen.scale` для устройств iOS. Дополнительные сведения см. в ок [. масштаб](https://developer.apple.com/documentation/uikit/uiscreen/1617836-scale).

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

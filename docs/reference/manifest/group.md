---
title: Элемент Group в файле манифеста
description: Определяет группу элементов управления пользовательского интерфейса на вкладке.
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 6fe07497e98bd77aad7ad296850a0b9f9e9bf9a4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718183"
---
# <a name="group-element"></a>Элемент Group

Определяет группу элементов пользовательского интерфейса на вкладке.  На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Да  | Уникальный идентификатор группы.|

### <a name="id-attribute"></a>Атрибут id

Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.

## <a name="child-elements"></a>Дочерние элементы
|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Label](#label)      | Да |  Метка элемента CustomTab или группы.  |
|  [Icon](icon.md)      | Да |  Изображение для группы.  |
|  [Control](#control)    | Да |  Коллекция одного или нескольких объектов Control.  |

### <a name="label"></a>Label 

Обязательный. Метка группы. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .

### <a name="icon"></a>Icon

Обязательный элемент. Если вкладка содержит большое количество групп и изменяется размер окна программы, вместо этого может отображаться указанное изображение.

### <a name="control"></a>Control
В группе должен быть по крайней мере один элемент управления. Дополнительные сведения о поддерживаемых типах элементов управления приведены в элементе [Control](control.md) .

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

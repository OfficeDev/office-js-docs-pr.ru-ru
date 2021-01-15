---
title: Элемент Group в файле манифеста
description: Определяет группу элементов управления пользовательского интерфейса на вкладке.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 3872ece926cc399ed2b30d4dabaacfb741e060ab
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771405"
---
# <a name="group-element"></a>Элемент Group

Определяет группу элементов управления пользовательского интерфейса на вкладке. На настраиваемой вкладке надстройка может создать несколько групп. Надстройка может создать не более одной специальной вкладки.

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
|  [Control](#control)    | Нет |  Представляет объект Control. Может иметь значение ноль или больше.  |
|  [OfficeControl](#officecontrol)  | Нет | Представляет один из встроенных элементов управления Office. Может иметь значение ноль или больше. |

### <a name="label"></a>Label

Обязательный элемент. Метка группы. Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)

### <a name="icon"></a>Icon

Обязательный элемент. Если вкладка содержит большое количество групп и размер окна программы не задан, вместо него может отображаться указанное изображение.

### <a name="control"></a>Средство контроля

Необязательный, но если его нет, должен быть хотя бы один **OfficeControl.** Подробные сведения о поддерживаемых типах элементов управления см. в [элементе Control.](control.md) Порядок элементов **управления** и **OfficeControl** в манифесте является взаимозаменяемым, и их можно перемещать, если существует несколько элементов, но все они должны быть под элементом **Icon.**

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

### <a name="officecontrol"></a>OfficeControl

Необязательный, но если его нет, должен быть хотя бы один **control.** Включаем один или несколько встроенных элементов управления Office в группу с `<OfficeControl>` элементами. Атрибут `id` указывает ИД встроенного в Office управления. Чтобы найти ИД элементов управления, см. поиск ИД элементов управления [и групп элементов управления.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Порядок элементов **управления** и **OfficeControl** в манифесте является взаимозаменяемым, и их можно перемещать, если существует несколько элементов, но все они должны быть под элементом **Icon.**

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
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

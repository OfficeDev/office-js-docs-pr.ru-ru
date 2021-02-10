---
title: Элемент Group в файле манифеста
description: Определяет группу элементов управления пользовательского интерфейса на вкладке.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173964"
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
|  [Control](#control)    | Нет |  Представляет объект Control. Может быть ноль или больше.  |
|  [OfficeControl](#officecontrol)  | Нет | Представляет один из встроенных элементов управления Office. Может быть ноль или больше. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должна ли группа отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки.  |

### <a name="label"></a>Label

Обязательный элемент. Метка группы. Атрибут **resid** может быть не более 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)

### <a name="icon"></a>Icon

Обязательный элемент. Если вкладка содержит большое количество групп и размер окна программы будет меняться, вместо него может отображаться указанное изображение.

### <a name="control"></a>Элемент управления

Необязательный, но если его нет, должен быть хотя бы один **OfficeControl.** Подробные сведения о поддерживаемых типах элементов управления см. в [элементе Control.](control.md) Порядок элементов **управления** и **OfficeControl** в манифесте является взаимозаменяемым, и их можно перемещать, если существует несколько элементов, но все они должны быть под элементом **Icon.**

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
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
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
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

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Необязательный (boolean). Указывает, будет  ли группа скрыта в сочетаниях приложений и платформ, которые поддерживают API, устанавливая настраиваемую контекстную вкладку на ленту во время работы. Значение по умолчанию (если его нет) `false` — . Если используется, **OverriddenByRibbonApi** должен быть первым *в* **группе.** Дополнительные сведения [см. в подразделе OverriddenByRibbonApi.](overriddenbyribbonapi.md)

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

---
title: Групповой элемент в файле манифеста
description: Определяет группу элементов управления пользовательским интерфейсом на вкладке.
ms.date: 06/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 09260ab52910235ab63149769cc989ffbda03ffb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154787"
---
# <a name="group-element"></a>Элемент Group

Определяет группу элементов управления пользовательским интерфейсом на вкладке. На настраиваемой вкладке надстройка может создавать несколько групп. Надстройка может создать не более одной специальной вкладки.

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
|  [Icon](icon.md)      | Да |  Изображение для группы. Не поддерживается Outlook надстройки. |
|  [Control](#control)    | Нет |  Представляет объект Control. Может быть ноль или больше.  |
|  [OfficeControl](#officecontrol)  | Нет | Представляет один из встроенных элементов Office элементов управления. Может быть ноль или больше. Не поддерживается Outlook надстройки.|
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должна ли группа отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки. Не поддерживается Outlook надстройки. |

### <a name="label"></a>Label

Обязательный элемент. Метка группы. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)

### <a name="icon"></a>Icon

Обязательный элемент. Если вкладка содержит большое количество групп и окно программы повторно, указанное изображение может отображаться вместо этого.

> [!NOTE]
> Этот элемент не поддерживается Outlook надстройки.

### <a name="control"></a>Средство контроля

Необязательный, но если его нет, то должен быть хотя бы один **OfficeControl.** Сведения о типах поддерживаемых элементов управления см. в [элементе Control.](control.md) Порядок  управления и **OfficeControl** в манифесте взаимозаменяем, и они могут быть взаимозаменяемыми, если существует несколько элементов, но все они должны быть ниже элемента **Icon.**

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

Необязательный, но если нет, должен быть хотя бы один **контроль.** Включай один или несколько встроенных элементов Office в группу с `<OfficeControl>` элементами. Атрибут указывает ID встроенного Office `id` управления. Чтобы найти ID элементов управления, см. в рублях [Find the IDs of controls and control groups.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Порядок  управления и **OfficeControl** в манифесте взаимозаменяем, и они могут быть взаимозаменяемыми, если существует несколько элементов, но все они должны быть ниже элемента **Icon.**

> [!NOTE]
> Этот элемент не поддерживается Outlook надстройки.

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

Необязательный (boolean). Указывает, будет ли **группа** скрыта в сочетаниях приложений и платформ, поддерживаюх API, который устанавливает настраиваемую контекстную вкладку на ленту во время запуска. Значение по умолчанию, если не присутствует, `false` является . Если используется, **OverriddenByRibbonApi** должен быть *первым* ребенком **группы**. Дополнительные сведения см. в [веб-сведениях OverriddenByRibbonApi](overriddenbyribbonapi.md).

> [!NOTE]
> Этот элемент не поддерживается Outlook надстройки.

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

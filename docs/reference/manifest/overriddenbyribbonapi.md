---
title: Элемент OverriddenByRibbonApi в файле манифеста
description: Узнайте, как указать, что настраиваемая вкладка, группа, элемент управления или пункт меню не должны отображаться, когда она также является частью настраиваемой контекстной вкладки.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 62aa484057221f9cd7f41af9c8b9210cdb5b3376
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50174002"
---
# <a name="overriddenbyribbonapi-element"></a>Элемент OverriddenByRibbonApi

Указывает, будет ли элемент [customTab,](customtab.md) [](control.md#menu-dropdown-button-controls) [group,](group.md) [button,](control.md#button-control) menu или элемент меню скрытыми в сочетаниях приложений и платформ, которые поддерживают API[(Office.ribbon.requestCreateControls),](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)устанавливая настраиваемые контекстные вкладки на ленте.

Если он опущен, значение по `false` умолчанию: . Если он используется, он должен быть первым *родительским* элементом.

> [!NOTE]
> Чтобы полностью понять этот элемент, прочитайте статью ["Реализация альтернативного интерфейса](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)пользователя, если настраиваемые контекстные вкладки не поддерживаются".

Цель этого элемента — создать в надстройке возможность отката, которая реализует настраиваемые контекстные вкладки, когда надстройка работает в приложении или платформе, которые не поддерживают настраиваемые контекстные вкладки. Основная стратегия заключается в том, что вы дублируете некоторые или все группы и элементы управления  с настраиваемой контекстной вкладки на одну или несколько настраиваемой основной вкладки (то есть неконтекстуальные настраиваемые вкладки). Затем, чтобы эти группы и элементы управления появлялись, когда настраиваемые контекстные вкладки  не поддерживаются, но не отображаются при поддержке настраиваемой контекстной вкладки, необходимо добавить в качестве первого потомка элементов  `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **CustomTab,** **Group,** **Control** или menu **Item.** Это может быть следующим образом:

- Если надстройка работает на приложениях и платформах, поддерживаюх настраиваемые контекстные вкладки, дублированные вкладки, группы и элементы управления не будут отображаться на ленте. Вместо этого настраиваемая контекстная вкладка будет установлена, когда надстройка вызывает `requestCreateControls` метод.
- Если надстройка работает в приложении или платформе, не поддерживаю которой настраиваемые контекстные вкладки, дублированные вкладки, группы и элементы управления будут отображаться на ленте. 

## <a name="examples"></a>Примеры

### <a name="overriding-an-entire-tab"></a>Переопределение всей вкладки

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-group"></a>Переопределение группы

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-control"></a>Переопределение управления

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
        <!-- Other child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-menu-item"></a>Переопределение элемента меню


```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Menu" id="MyMenu">
        <!-- Other child elements omitted. -->
        <Items>
          <Item id="showGallery">
            <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
            <!-- Other child elements omitted. -->
          </Item>
        </Items>
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

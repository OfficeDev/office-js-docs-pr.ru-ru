---
title: Элемент OverriddenByRibbonApi в файле манифеста
description: Узнайте, как указать, что настраиваемая вкладка, группа, элемент управления или меню не должны отображаться, когда он также является частью настраиваемой контекстной вкладки.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 0f314761f686ca559caea4e04ec5d5a66fab9618ea21a221a6cf2affde897578
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092214"
---
# <a name="overriddenbyribbonapi-element"></a>Элемент OverriddenByRibbonApi

Указывает, будет ли элемент [](group.md) [CustomTab, Group,](customtab.md) [Button](control.md#button-control) Control, [Menu](control.md#menu-dropdown-button-controls) или элемент меню скрыт в сочетаниях приложений и платформ, поддерживаюх[API (Office.ribbon.requestCreateControls),](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)который устанавливает настраиваемые контекстные вкладки на ленте.

Если он опущен, по умолчанию `false` . Если он используется, он должен быть *первым* детским элементом родительского элемента.

> [!NOTE]
> Полное представление об этом элементе читайте в публикации [Implement an alternate UI experience when custom contextual tabs are not supported.](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)

Цель этого элемента заключается в создании опыта отката в надстройке, которая реализует настраиваемые контекстные вкладки, когда надстройка выполняется на приложении или платформе, которая не поддерживает настраиваемые контекстные вкладки. Основная стратегия заключается в том, что вы дублируете некоторые или все группы и элементы управления  из настраиваемой контекстной вкладки на одну или несколько пользовательских вкладок ядра (то есть неконтекстуальных пользовательских вкладок). Затем, чтобы убедиться, что эти группы и  элементы управления отображаются, когда настраиваемые  контекстные вкладки не поддерживаются, но не отображаются при поддержке настраиваемой контекстной вкладки, вы добавляете в качестве первого детского элемента `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` элементы **CustomTab**, **Group,** **Control** или menu **Item.** Эффект от этого ниже:

- Если надстройка работает на приложении и платформе, поддерживаюх настраиваемые контекстные вкладки, то дублированные вкладки, группы и элементы управления не будут отображаться на ленте. Вместо этого настраиваемая контекстная вкладка будет установлена, когда надстройка вызывает `requestCreateControls` метод.
- Если надстройка работает на приложении или платформе, не поддерживаюх настраиваемые контекстные вкладки, на ленте будут отображаться дублированные вкладки, группы и элементы управления. 

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

---
title: Элемент OverriddenByRibbonApi в файле манифеста
description: Узнайте, как указать, что настраиваемая вкладка, группа, элемент управления или меню не должны отображаться, когда он также является частью настраиваемой контекстной вкладки.
ms.date: 09/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 35893bba5c00d8b6d63f02cc12ac6902197ab0d8
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154979"
---
# <a name="overriddenbyribbonapi-element"></a>Элемент OverriddenByRibbonApi

Указывает, будет ли элемент [управления](group.md) [группой,](control.md#button-control) [кнопкой,](control.md#menu-dropdown-button-controls) элементом меню или элементом меню скрываться в сочетаниях приложений и платформ, поддерживаюх[API (Office.ribbon.requestCreateControls),](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)который устанавливает настраиваемые контекстные вкладки на ленте.

Если он опущен, по умолчанию `false` . Если он используется, он должен быть *первым* детским элементом родительского элемента.

> [!NOTE]
> Полное представление об этом элементе читайте в публикации [Implement an alternate UI experience when custom contextual tabs are not supported.](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)

Цель этого элемента заключается в создании опыта отката в надстройке, которая реализует настраиваемые контекстные вкладки, когда надстройка выполняется на приложении или платформе, которая не поддерживает настраиваемые контекстные вкладки. Основная стратегия заключается в том, что вы дублируете некоторые или все группы и элементы управления  из настраиваемой контекстной вкладки на одну или несколько пользовательских вкладок ядра (то есть неконтекстуальных пользовательских вкладок). Затем, чтобы убедиться, что эти группы и  элементы управления отображаются, когда настраиваемые  контекстные вкладки не поддерживаются, но не отображаются при поддержке настраиваемой контекстной вкладки, вы добавляете в качестве первого детского элемента `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **элементов Group,** **Control** или menu **Item.** Эффект от этого ниже:

- Если надстройка работает на приложении и платформе, поддерживаюх настраиваемые контекстные вкладки, то дублированные группы и элементы управления не будут отображаться на ленте. Вместо этого настраиваемая контекстная вкладка будет установлена, когда надстройка вызывает `requestCreateControls` метод.
- Если надстройка работает на приложении или платформе, не поддерживаюх настраиваемые контекстные вкладки, на ленте будут отображаться дублированные группы и элементы управления. 

## <a name="examples"></a>Примеры

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

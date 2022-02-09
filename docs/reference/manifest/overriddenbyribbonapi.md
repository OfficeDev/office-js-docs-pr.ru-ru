---
title: Элемент OverriddenByRibbonApi в файле манифеста
description: Узнайте, как указать, что настраиваемая вкладка, группа, элемент управления или меню не должны отображаться, когда он также является частью настраиваемой контекстной вкладки.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48977691ee4bf2ccd71bc146647dae452ce9e2fc
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467689"
---
# <a name="overriddenbyribbonapi-element"></a>Элемент OverriddenByRibbonApi

Указывает, будет ли элемент [управления группой](group.md)[,](control-button.md) кнопкой[,](control-menu.md) элементом меню или элементом меню скрываться в сочетаниях приложений и платформ, которые поддерживают API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1))), который устанавливает настраиваемые контекстные вкладки на ленте.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [Лента 1.2](../requirement-sets/add-in-commands-requirement-sets.md) (требуется для Excel, PowerPoint и Word.)

Если этот элемент опущен, по умолчанию .`false` Если он используется, он должен быть первым детским  элементом родительского элемента.

> [!NOTE]
> Полное представление об этом элементе читайте в публикации [Implement an alternate UI experience when custom contextual tabs are not supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

Цель этого элемента заключается в создании опыта отката в надстройке, которая реализует настраиваемые контекстные вкладки, когда надстройка выполняется на приложении или платформе, которая не поддерживает настраиваемые контекстные вкладки. Основная стратегия заключается в том, что вы дублируете некоторые или все группы и элементы управления из настраиваемой контекстной вкладки на одну или несколько пользовательских вкладок ядра (то есть *неконтекстуальных* пользовательских вкладок). Затем, чтобы убедиться, что эти группы и элементы управления отображаются, когда настраиваемые контекстные вкладки не поддерживаются, но не отображаются  при поддержке настраиваемой контекстной вкладки, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` вы добавляете в качестве первого детского элемента элементов **Group**, **Control** или menu **Item**. Эффект от этого ниже:

- Если надстройка работает на приложении и платформе, поддерживаюх настраиваемые контекстные вкладки, то дублированные группы и элементы управления не будут отображаться на ленте. Вместо этого настраиваемая контекстная вкладка будет установлена, когда надстройка вызывает `requestCreateControls` метод.
- Если надстройка работает на приложении или платформе, не поддерживаюх настраиваемые контекстные вкладки, на ленте будут отображаться дублированные группы и элементы управления.

## <a name="examples"></a>Примеры

### <a name="overriding-a-group"></a>Переопределение группы

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.CustomTab1.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="Contoso.MyButton1">
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
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.CustomTab2.group2">
      <Control  xsi:type="Button" id="Contoso.MyButton2">
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
  <CustomTab id="Contoso.TabCustom3">
    <Group id="Contoso.CustomTab3.group3">
      <Control  xsi:type="Menu" id="Contoso.MyMenu">
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

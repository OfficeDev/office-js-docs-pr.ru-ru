---
title: Элемент управления типом Меню в файле манифеста
description: Определяет меню, элементы которого могут выполнять действия или запускать области задач.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7287b8e2cdf2378140ef50a41306820a0fd4002f
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467922"
---
# <a name="control-element-of-type-menu"></a>Элемент управления меню типа

Меню определяет список параметров. Каждый элемент меню либо выполняет функцию, либо отображает область задач.

> [!NOTE]
> В этой статье предполагается знакомство с базовой справочной [статьей Control,](control.md) которая содержит важные сведения о атрибутах элемента.

Управление меню определяет:

- Управление меню на корневом уровне.
- Список элементов меню.

Когда используется точка **расширения PrimaryCommandSurface**[, элемент](extensionpoint.md) корневого меню отображается в качестве кнопки на ленте. Когда выбрана кнопка, меню отображается в качестве списка выпаданий. Подменю не поддерживаются.

Когда используется с **точкой расширения ContextMenu**[, элемент](extensionpoint.md) корневого меню отображается в контексте меню. При выборе корневого элемента элементы меню отображаются в качестве подмену. Ни один из элементов не может быть подмену, так как поддерживается только один уровень подменуса.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Label](#label)     | Да |  Текст меню. |
|  **ToolTip**    |Нет|Всплывающая подсказка для меню. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** . **String** — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).|
|  [Supertip](supertip.md)  | Да |  Супертип этого меню.    |
|  [Icon](icon.md)      | Да |  Изображение для меню.         |
|  **Items**     | Да |  Коллекция элементов, отображаемая в меню. Содержит элемент **Item** для каждого элемента. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должно ли меню отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки. Если используется, он должен быть первым *элементом* ребенка. |

### <a name="label"></a>Метка

Указывает текст для имени меню с помощью его только атрибута **resid**, который может быть не более 32 символов и должен быть задан значению атрибута **id** элемента **String** в ребенке **ShortStrings** элемента [Resources](resources.md) .

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) , когда родительский **VersionOverrides** — это тип Taskpane 1.0.
- [Почтовый ящик 1.3,](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) когда родительский **VersionOverrides** — это тип Почта 1.0.
- [Почтовый ящик 1.5,](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) когда родительский **VersionOverrides** — это тип Почта 1.1.

## <a name="examples"></a>Примеры

В следующем примере меню имеет два пункта. Первый отображает области задач. Второй выполняет функцию. Меню настроено на то, чтобы  не быть видимым при запуске надстройки на платформе, которая поддерживает контекстные вкладки. Дополнительные сведения читайте в материале [Реализация альтернативного интерфейса,](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported) когда пользовательские контекстные вкладки не поддерживаются.

```xml
<Control xsi:type="Menu" id="Contoso.TestMenu2">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="ShowMainTaskPane">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="GetData">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getData</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

В следующем примере второй элемент меню настроен на то, чтобы не быть видимым при запуске надстройки на платформе, которая поддерживает контекстные вкладки. Дополнительные сведения читайте в материале [Реализация альтернативного интерфейса,](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported) когда пользовательские контекстные вкладки не поддерживаются.

```xml
<Control xsi:type="Menu" id="Contoso.msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="ShowMainTaskPane">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="msgReadMenuItem1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

---
title: Элемент управления типом Button в файле манифеста
description: Определяет кнопку, которая выполняет действие или запускает области задач.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: adc58424fe9898bffcbd9e16bed8f3b13b9df4a2
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467911"
---
# <a name="control-element-of-type-button"></a>Элемент управления типом Button

Определяет кнопку, которая выполняет действие или запускает области задач.

> [!NOTE]
> В этой статье предполагается знакомство с базовой справочной [статьей Control,](control.md) которая содержит важные сведения о атрибутах элемента.

Когда пользователь нажимает кнопку, она выполняет одно действие. Она может выполнять функцию или отображать область задач. Каждый элемент управления кнопками `id` должен иметь уникальное значение атрибута среди всех элементов **управления** в манифесте.

> [!IMPORTANT]
> Элементы управления типа "Кнопка" игнорируются на мобильных платформах. Чтобы поддерживать мобильные платформы, необходимо также иметь управление типом "MobileButton" для каждого управления типом "Button".

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Label](#label)     | Да |  Текст для кнопки. |
|  **ToolTip**    |Нет|Подсказка для кнопки. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** . **String** — это дочерний элемент **LongStrings**, являющийся дочерним для элемента [Resources](resources.md).|
|  [Supertip](supertip.md)  | Да |  Суперподсказка для кнопки.    |
|  [Icon](icon.md)      | Да |  Изображение для кнопки.         |
|  [Action](action.md)    | Да |  Указание действия, которое предстоит выполнить. **Элементу Control** может быть только один ребенок **действия**. |
|  [Enabled](enabled.md)    | Нет |  Указывает, включен ли контроль при запуске надстройки.  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должна ли кнопка отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки. Если используется, он должен быть первым *элементом* ребенка. |

### <a name="label"></a>Метка

Указывает текст для кнопки с помощью его только атрибута **resid**, который может быть не более 32 символов и должен быть задан к значению атрибута **id** элемента **String** в ребенке **ShortStrings** элемента [Resources](resources.md) .

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

В следующем примере кнопка выполняет функцию. Он также настроен на отключение при запуске надстройки. Его можно включить программным путем. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).

```xml
<Control xsi:type="Button" id="Contoso.msgReadFunctionButton">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
  <Enabled>false</Enabled>
</Control>
```

В следующем примере кнопка отображает области задач.

```xml
<Control xsi:type="Button" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

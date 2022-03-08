---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6a9540fd7e98464681a90021a36f7a7529186f7f
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340115"
---
# <a name="customtab-element"></a>Элемент CustomTab

Определяет настраиваемую вкладку для Office ленты. Добавьте элементы управления лентой и группы для надстройки в одну из вкладок Office или на собственную настраиваемую вкладку. Используйте элемент **CustomTab**, чтобы добавить настраиваемую вкладку в ленту. На настраиваемой вкладке надстройка может иметь настраиваемые или встроенные группы. Надстройка может создать не более одной специальной вкладки.

> [!IMPORTANT]
> В Outlook Mac элемент **CustomTab** не доступен, но вместо этого можно поместить настраиваемые группы элементов управления  на один из встроенных [OfficeTabs](officetab.md). Встроенные *группы нельзя* ставить на встроенные  вкладки в Outlook на любой платформе.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

> [!NOTE]
> Некоторые детские элементы не допустимы в схемах Почты. См [. элементы Child](#child-elements).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)
- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md). Необходимые для некоторых элементов ребенка. См [. элементы Child](#child-elements).

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Да  | Уникальный ID для настраиваемой вкладки.|

### <a name="id-attribute"></a>Атрибут id

Обязательный элемент. Уникальный идентификатор для настраиваемой вкладки. Это строка с максимальным значением 125 символов. Это должно быть уникальным в манифесте.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Нет |  Определяет группу команд.  |
|  [OfficeGroup](#officegroup)      | Нет |  Представляет встроенную группу Office управления. **Важно**: в Outlook. |
|  [Label](#label-tab)      | Да |  Метка для CustomTab.  |
|  [InsertAfter](#insertafter)      | Нет |  Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки **Office. Важно**: доступна только в PowerPoint. |
|  [InsertBefore](#insertbefore)      | Нет |  Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке **Office. Важно**: доступна только в PowerPoint. |

### <a name="group"></a>Group

Необязательный, но если его нет, должен быть по крайней мере один **элемент OfficeGroup** . См [. элемент Group](group.md). Порядок **Групповой и** **OfficeGroup** в манифесте должен быть тем, который вы хотите, чтобы они появились на настраиваемой вкладке. Они могут быть перемеяны, если существует несколько элементов, но все они должны быть выше элемента **Label** .

### <a name="officegroup"></a>OfficeGroup

Необязательный, но если его нет, то должен быть по крайней мере один **элемент Group** . Представляет встроенную группу Office управления. Атрибут **id** указывает ID встроенной Office группы. Чтобы найти ID встроенной группы, см. в рублях [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). Порядок **Групповой и** **OfficeGroup** в манифесте должен быть тем, который вы хотите, чтобы они появились на настраиваемой вкладке. Они могут быть перемеяны, если существует несколько элементов, но все они должны быть выше элемента **Label** .

> [!IMPORTANT]
> Элемент **OfficeGroup** не доступен в Outlook. В PowerPoint он находится в предварительном режиме для Mac и Windows; но доступен для надстройок в PowerPoint в Интернете.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="label-tab"></a>Label (Tab)

Обязательный элемент. Метка настраиваемой вкладки. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources](resources.md) .

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertafter"></a>InsertAfter

Необязательное свойство. Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной Office вкладки. Значение элемента — это ID встроенной вкладки, например или `TabHome` `TabReview`.  Список встроенных вкладок см. в [officeTab](officetab.md). Если присутствует, должно быть после элемента **Label** . Нельзя иметь **и InsertAfter,** и **InsertBefore**.

> [!IMPORTANT]
> Элемент **InsertAfter** доступен только в PowerPoint.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertbefore"></a>InsertBefore

Необязательное свойство. Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной Office вкладке. Значение элемента — это ID встроенной вкладки, например или `TabHome` `TabReview`. Значение элемента — это ID встроенной вкладки, например или `TabHome` `TabReview`.  Список встроенных вкладок см. в [officeTab](officetab.md). Если присутствует, должно быть после элемента **Label** . Нельзя иметь **и InsertAfter,** и **InsertBefore**.

> [!IMPORTANT]
> Элемент **InsertBefore** доступен только в PowerPoint.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)


## <a name="examples"></a>Примеры

В следующем примере разметки группа управления Office абзаца добавляется в настраиваемую вкладку и позиционет ее, чтобы она появляться сразу после настраиваемой группы.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.TabCustom1.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

В следующем примере разметки добавляется Office superscript в настраиваемую группу и он должен отображаться сразу после настраиваемой кнопки.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 08/13/2021
localization_priority: Normal
ms.openlocfilehash: 3656f68a722e5e0c224f18f80a0e0214fce47cfb
ms.sourcegitcommit: bc6203dd8f21d1c375039c5ee8f1388ede9be93b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/18/2021
ms.locfileid: "58382965"
---
# <a name="customtab-element"></a>Элемент CustomTab

На ленте укажите вкладку и группу для команд надстройки. Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.

На настраиваемой вкладке надстройка может иметь настраиваемые или встроенные группы. Надстройка может создать не более одной специальной вкладки.

Атрибут **id** должен быть уникальным в манифесте.

> [!IMPORTANT]
> В Outlook Mac элемент не доступен, поэтому вместо него необходимо использовать `CustomTab` [OfficeTab.](officetab.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Нет |  Определяет группу команд.  |
|  [OfficeGroup](#officegroup)      | Нет |  Представляет встроенную группу Office управления. **Важно:** недоступна в Outlook. |
|  [Label](#label-tab)      | Да |  Метка элемента CustomTab или Group.  |
|  [InsertAfter](#insertafter)      | Нет |  Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки Office. **Важно:** доступно только в PowerPoint. |
|  [InsertBefore](#insertbefore)      | Нет |  Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке Office. **Важно:** доступна только в PowerPoint. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должна ли настраиваемая вкладка отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки. **Важно:** недоступна в Outlook. |

### <a name="group"></a>Group

Необязательный, но если его нет, должен быть по крайней мере один **элемент OfficeGroup.** См. [элемент Group.](group.md) Порядок **Групповой и** **OfficeGroup** в манифесте должен быть тем, который вы хотите, чтобы они появились на настраиваемой вкладке. Они могут быть перемеяны, если существует несколько элементов, но все они должны быть выше элемента **Label.**

### <a name="officegroup"></a>OfficeGroup

Необязательный, но если его нет, то должен быть по крайней мере один **элемент Group.** Представляет встроенную группу Office управления. Атрибут **id** указывает ID встроенной Office группы. Чтобы найти ID встроенной группы, см. в рублях [Find the IDs of controls and control groups.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Порядок **Групповой и** **OfficeGroup** в манифесте должен быть тем, который вы хотите, чтобы они появились на настраиваемой вкладке. Они могут быть перемеяны, если существует несколько элементов, но все они должны быть выше элемента **Label.**

> [!IMPORTANT]
> Элемент `OfficeGroup` не доступен в Outlook.

### <a name="label-tab"></a>Label (Tab)

Обязательно. Метка настраиваемой вкладки. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)

### <a name="insertafter"></a>InsertAfter

Необязательно. Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной Office вкладки. Значение элемента — это ID встроенной вкладки, например TabHome или TabReview. [(См. поиск ID элементов управления и групп управления.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Если присутствует, должно быть после элемента **Label.** Нельзя иметь и **InsertAfter,** и **InsertBefore.**

> [!IMPORTANT]
> Элемент `InsertAfter` доступен только в PowerPoint.

### <a name="insertbefore"></a>InsertBefore

Необязательно. Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной Office вкладке. Значение элемента — это ID встроенной вкладки, например TabHome или TabReview. [(См. поиск ID элементов управления и групп управления.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  Если присутствует, должно быть после элемента **Label.** Нельзя иметь и **InsertAfter,** и **InsertBefore.**

> [!IMPORTANT]
> Элемент `InsertBefore` доступен только в PowerPoint.

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Необязательный (boolean). Указывает, будет ли **CustomTab** скрыт в сочетаниях приложений и платформ, поддерживаюх API, устанавливаемую настраиваемую контекстную вкладку на ленту во время работы. Значение по умолчанию, если не присутствует, `false` является . Если используется, **OverriddenByRibbonApi** должен быть *первым* ребенком **CustomTab**. Дополнительные сведения см. в [веб-сведениях OverriddenByRibbonApi](overriddenbyribbonapi.md).

> [!IMPORTANT]
> Элемент `OverriddenByRibbonApi` не доступен в Outlook.

## <a name="customtab-example"></a>Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

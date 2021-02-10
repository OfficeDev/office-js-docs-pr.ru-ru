---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: d74859d1326d29517b5a8226a86f901322957933
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173929"
---
# <a name="customtab-element"></a>Элемент CustomTab

На ленте укажите вкладку и группу для команд надстройки. Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.

На настраиваемой вкладке надстройка может иметь настраиваемые или встроенные группы. Надстройка может создать не более одной специальной вкладки.

Атрибут **id** должен быть уникальным в манифесте.

> [!IMPORTANT]
> В Outlook для Mac элемент не доступен, поэтому необходимо использовать `CustomTab` [OfficeTab.](officetab.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Нет |  Определяет группу команд.  |
|  [OfficeGroup](#officegroup)      | Нет |  Представляет встроенную группу управления Office. **Важно!** Отсутствует в Outlook. |
|  [Label](#label-tab)      | Да |  Метка элемента CustomTab или Group.  |
|  [InsertAfter](#insertafter)      | Нет |  Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки **Office.** Важно: отсутствует в Outlook. |
|  [InsertBefore](#insertbefore)      | Нет |  Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке **Office.** Важно! Отсутствует в Outlook. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должна ли настраиваемая вкладка отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки. **Важно!** Отсутствует в Outlook. |

### <a name="group"></a>Группа

Необязательный, но если его нет, должен быть хотя бы один **элемент OfficeGroup.** См. [элемент Group.](group.md) Порядок **групп и** **OfficeGroup** в манифесте должен быть в том порядке, в который они должны отображаться на настраиваемой вкладке. Их можно перемесить, если существует несколько элементов, но все они должны быть над **элементом Label.**

### <a name="officegroup"></a>OfficeGroup

Необязательный, но если его нет, должен быть хотя бы один **элемент Group.** Представляет встроенную группу управления Office. Атрибут **id** указывает ИД встроенной группы Office. Чтобы найти ИД встроенной группы, см. поиск ИД элементов управления [и групп элементов управления.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Порядок **групп и** **OfficeGroup** в манифесте должен быть в том порядке, в который они должны отображаться на настраиваемой вкладке. Их можно перемесить, если существует несколько элементов, но все они должны быть над **элементом Label.**

> [!IMPORTANT]
> Элемент `OfficeGroup` не доступен в Outlook.

### <a name="label-tab"></a>Label (Tab)

Обязательно. Метка пользовательской вкладки. Атрибут **resid** может быть не более 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)

### <a name="insertafter"></a>InsertAfter

Необязательно. Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки Office. Значением элемента является ИД встроенной вкладки, например TabHome или TabReview. [(См. поиск ИД элементов управления и групп элементов управления.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Если этот элемент заметим, он должен быть после **элемента Label.** Невозможно одновременное **добавление InsertAfter** и **InsertBefore.**

> [!IMPORTANT]
> Элемент `InsertAfter` не доступен в Outlook.

### <a name="insertbefore"></a>InsertBefore

Необязательно. Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке Office. Значением элемента является ИД встроенной вкладки, например TabHome или TabReview. [(См. поиск ИД элементов управления и групп элементов управления.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  Если этот элемент заметим, он должен быть после **элемента Label.** Невозможно одновременное **добавление InsertAfter** и **InsertBefore.**

> [!IMPORTANT]
> Элемент `InsertBefore` не доступен в Outlook.

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Необязательный (boolean). Указывает, будет ли **customTab** скрыт в сочетаниях приложений и платформ, которые поддерживают API, устанавливая настраиваемую контекстную вкладку на ленту во время работы. Значение по умолчанию (если его нет) `false` — . Если используется, **OverriddenByRibbonApi**  должен быть первым child of **CustomTab.** Дополнительные сведения [см. в подразделе OverriddenByRibbonApi.](overriddenbyribbonapi.md)

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

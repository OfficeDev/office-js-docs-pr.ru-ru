---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 642222af02431814e4e64141504911c67ca829fa
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771328"
---
# <a name="customtab-element"></a>Элемент CustomTab

На ленте укажите вкладку и группу для команд надстройки. Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.

На настраиваемой вкладке надстройка может иметь настраиваемые или встроенные группы. Надстройка может создать не более одной специальной вкладки.

Атрибут **id** должен быть уникальным в манифесте.

> [!IMPORTANT]
> В Outlook для Mac элемент не доступен, поэтому придется `CustomTab` использовать [OfficeTab.](officetab.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Нет |  Определяет группу команд.  |
|  [OfficeGroup](#officegroup)      | Нет |  Представляет встроенную группу управления Office.  |
|  [Label](#label-tab)      | Да |  Метка элемента CustomTab или Group.  |
|  [InsertAfter](#insertafter)      | Нет |  Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки Office.  |
|  [InsertBefore](#insertbefore)      | Нет |  Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке Office.  |

### <a name="group"></a>Группа

Необязательный, но если его нет, должен быть хотя бы один **элемент OfficeGroup.** См. [элемент Group.](group.md) Порядок **групп и** **OfficeGroup** в манифесте должен быть в том порядке, в который они должны отображаться на настраиваемой вкладке. Они могут быть перемелены, если существует несколько элементов, но все они должны быть над **элементом Label.**

### <a name="officegroup"></a>OfficeGroup

Необязательный, но если его нет, должен быть хотя бы один **элемент Group.** Представляет встроенную группу управления Office. Атрибут **id** указывает ИД встроенной группы Office. Чтобы найти ИД встроенной группы, см. поиск ИД элементов управления [и групп элементов управления.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Порядок **групп и** **OfficeGroup** в манифесте должен быть в том порядке, в который они должны отображаться на настраиваемой вкладке. Они могут быть перемелены, если существует несколько элементов, но все они должны быть над **элементом Label.**

### <a name="label-tab"></a>Label (Tab)

Обязательный. Метка пользовательской вкладки. Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)

### <a name="insertafter"></a>InsertAfter

Необязательное свойство. Указывает, что настраиваемая вкладка должна быть сразу после указанной встроенной вкладки Office. Значением элемента является ИД встроенной вкладки, например TabHome или TabReview. [(См. "Поиск ИД элементов управления и групп элементов управления".](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Если этот элемент заметок, он должен быть после **элемента Label.** Невозможно одновременно **insertAfter** и **InsertBefore.**

### <a name="insertbefore"></a>InsertBefore

Необязательное свойство. Указывает, что настраиваемая вкладка должна быть непосредственно перед указанной встроенной вкладке Office. Значением элемента является ИД встроенной вкладки, например TabHome или TabReview. [(См. "Поиск ИД элементов управления и групп элементов управления".](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  Если этот элемент заметок, он должен быть после **элемента Label.** Невозможно одновременно **insertAfter** и **InsertBefore.**

## <a name="customtab-example"></a>Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

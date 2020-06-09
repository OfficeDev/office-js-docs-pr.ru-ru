---
title: Элемент CustomTab в файле манифеста
description: На ленте можно указать вкладку и группу для команд надстройки.
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: a81b64a17eeeb463d55024e189b09048b2eb96ac
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612306"
---
# <a name="customtab-element"></a>Элемент CustomTab

На ленте можно указать вкладку и группу для команд надстройки. Они могут находиться либо на вкладке по умолчанию (**Главная**, **Сообщение** или **Собрание**), либо на вкладке, определенной надстройкой.

На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.

Атрибут **ID** должен быть уникальным в пределах манифеста.

> [!IMPORTANT]
> В Outlook на Mac `CustomTab` элемент недоступен, поэтому необходимо использовать [OfficeTab](officetab.md) .

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Да |  Определяет группу команд.  |
|  [Label](#label-tab)      | Да |  Метка элемента CustomTab или Group.  |

### <a name="group"></a>Group

Обязательный. См. статью об [элементе Group](group.md).

### <a name="label-tab"></a>Label (Tab)

Обязательный элемент. Метка настраиваемой вкладки. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .


## <a name="customtab-example"></a>Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

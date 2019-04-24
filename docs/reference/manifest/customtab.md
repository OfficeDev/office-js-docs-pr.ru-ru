---
title: Элемент CustomTab в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c1c3c6883a1feb94299feb35c078431e6e2e322c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450634"
---
# <a name="customtab-element"></a>Элемент CustomTab

На ленте можно указать вкладку и группу для команд надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка.

На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.

Атрибут **id** должен быть уникальным для манифеста.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Да |  Определяет группу команд.  |
|  [Label](#label-tab)      | Да |  Метка элемента CustomTab или Group.  |
|  [Control](control.md)    | Да |  Коллекция из одного или нескольких объектов Control.  |

### <a name="group"></a>Group

Обязательный. См. статью об [элементе Group](group.md).

### <a name="label-tab"></a>Label (Tab)

Обязательный элемент. Метка настраиваемой вкладки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).


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

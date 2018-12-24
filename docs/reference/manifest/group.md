---
title: Элемент Group в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 13cd9bbe6f602fd1779caea487e34177c3e9d483
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433703"
---
# <a name="group-element"></a>Элемент Group

Определяет группу элементов пользовательского интерфейса на вкладке.  На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Да  | Уникальный идентификатор группы.|

### <a name="id-attribute"></a>Атрибут id

Обязательный. Уникальный идентификатор группы. Это строка длиной до 125 символов. Она должна быть уникальной в пределах манифеста. В противном случае отобразить группу не удастся.

## <a name="child-elements"></a>Дочерние элементы
|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Label](#label)      | Да |  Метка элемента CustomTab или группы.  |
|  [Control](#control)    | Да |  Коллекция одного или нескольких объектов Control.  |

### <a name="label"></a>Label 

Обязательный элемент. Метка группы. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).

### <a name="control"></a>Control
Для группы требуется по крайней мере один элемент управления.

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
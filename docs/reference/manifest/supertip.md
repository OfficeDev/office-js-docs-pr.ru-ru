---
title: Элемент Supertip в файле манифеста
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 269a3723db6f98cdb25c61e5a88608c5fb5f3191
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659658"
---
# <a name="supertip"></a>Supertip

Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Title](#title) | Да | Текст подсказки. |
| [Description](#description) | Да | Описание подсказки.<br>**Note**: (Outlook) поддерживаются только клиенты Windows и Mac. |

### <a name="title"></a>Название

Обязательный элемент. Текст суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).

### <a name="description"></a>Описание

Обязательный элемент. Описание суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **LongStrings**, вложенном в элемент [Resources](resources.md).

> [!NOTE]
> В Outlook только клиенты Windows и Mac поддерживают элемент **Description** .

## <a name="example"></a>Пример

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

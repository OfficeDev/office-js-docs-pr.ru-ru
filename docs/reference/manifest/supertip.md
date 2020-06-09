---
title: Элемент Supertip в файле манифеста
description: Элемент SuperTip определяет расширенную подсказку (название и описание).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 8061c9dcd7903db0f1265084498d6c86654e1dfa
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608721"
---
# <a name="supertip"></a>Supertip

Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Title](#title) | Да | Текст подсказки. |
| [Description](#description) | Да | Описание подсказки.<br>**Note**: (Outlook) поддерживаются только клиенты Windows и Mac. |

### <a name="title"></a>Название

Обязательный элемент. Текст суперподсказки. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .

### <a name="description"></a>Описание

Обязательный. Описание суперподсказки. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **LongStrings** элемента [Resources](resources.md) .

> [!NOTE]
> В Outlook только клиенты Windows и Mac поддерживают элемент **Description** .

## <a name="example"></a>Пример

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

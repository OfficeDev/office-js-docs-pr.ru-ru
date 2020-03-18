---
title: Элемент Supertip в файле манифеста
description: Элемент SuperTip определяет расширенную подсказку (название и описание).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: cf88473b72979c839e5d55f44938fda19be24084
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720353"
---
# <a name="supertip"></a>Supertip

Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Title](#title) | Да | Текст подсказки. |
| [Description](#description) | Да | Описание подсказки.<br>**Note**: (Outlook) поддерживаются только клиенты Windows и Mac. |

### <a name="title"></a>Название

Обязательный. Текст суперподсказки. Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .

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

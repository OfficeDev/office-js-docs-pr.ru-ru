---
title: Элемент Supertip в файле манифеста
description: Элемент Supertip определяет богатый инструментарий (как название, так и описание).
ms.date: 05/07/2019
ms.localizationpriority: medium
ms.openlocfilehash: 6c1e73b0aba5923992fba03b78744ae5d34fb5da
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154434"
---
# <a name="supertip"></a>Supertip

Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Title](#title) | Да | Текст подсказки. |
| [Description](#description) | Да | Описание подсказки.<br>**Примечание.**(Outlook) поддерживаются только Windows и mac-клиенты. |

### <a name="title"></a>Title

Обязательный. Текст суперподсказки. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)

### <a name="description"></a>Описание

Обязательный. Описание суперподсказки. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе LongStrings** в [элементе Resources.](resources.md)

> [!NOTE]
> Для Outlook только клиенты Windows и Mac поддерживают элемент **Description.**

## <a name="example"></a>Пример

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

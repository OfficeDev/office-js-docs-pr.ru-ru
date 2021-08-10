---
title: Элемент Supertip в файле манифеста
description: Элемент Supertip определяет богатый инструментарий (как название, так и описание).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 79120cc72aa4804eaaa2330d9298f6521a13552d325d9134814581402ace8210
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093264"
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

---
title: Элемент Event в файле манифеста
description: Определяет обработчик событий в надстройке.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 80f21d1819e3d7e335389070ccac0db583026045
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275709"
---
# <a name="event-element"></a>Элемент Event

Определяет обработчик событий в надстройке.

> [!NOTE]
> Сведения о поддержке и использовании можно найти [в статье функция отправки почты для надстроек Outlook](../../outlook/outlook-on-send-addins.md).

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Тип](#type-attribute)  |  Да  | Задает обрабатываемое событие. |
|  [функтионексекутион](#functionexecution-attribute)  |  Да  | Задает способ выполнения обработчика событий (асинхронное или синхронное). В настоящее время поддерживаются только синхронные обработчики событий. |
|  [FunctionName](#functionname-attribute)  |  Да  | Задает имя функции для обработчика событий. |

### <a name="type-attribute"></a>Атрибут Type

Обязательный. Указывает событие, при возникновении которого вызывается обработчик. В приведенной ниже таблице представлены допустимые значения этого атрибута.

|  Тип события  |  Описание  |
|:-----|:-----|
|  `ItemSend`  |  Обработчик события будет вызван, когда пользователь отправляет сообщение или приглашение на собрание.  |

### <a name="functionexecution-attribute"></a>Атрибут FunctionExecution

Обязательный. ДОЛЖНО быть задано значение `synchronous`.

### <a name="functionname-attribute"></a>Атрибут FunctionName

Обязательный. Задает имя функции для обработчика событий. Это значение должно совпадать с именем функции в [файле функции](functionfile.md) надстройки.

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```

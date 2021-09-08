---
title: Элемент события в файле манифеста
description: Определяет обработчик событий в надстройке.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 3d8e94c10bed214dd976b3048e11328f10f99325
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937642"
---
# <a name="event-element"></a>Элемент Event

Определяет обработчик событий в надстройке.

> [!NOTE]
> Сведения о поддержке и использовании см. в сайте [On-send feature for Outlook надстройки.](../../outlook/outlook-on-send-addins.md)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Тип](#type-attribute)  |  Да  | Задает обрабатываемое событие. |
|  [FunctionExecution](#functionexecution-attribute)  |  Да  | Задает способ выполнения обработчика событий (асинхронное или синхронное). В настоящее время поддерживаются только синхронные обработчики событий. |
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

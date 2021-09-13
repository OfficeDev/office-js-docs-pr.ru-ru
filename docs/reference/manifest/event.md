---
title: Элемент события в файле манифеста
description: Определяет обработчик событий в надстройке.
ms.date: 05/15/2020
ms.localizationpriority: medium
ms.openlocfilehash: d5ccddc64ffecd9ebc06b28eb37c0aee46dcc2f4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153997"
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

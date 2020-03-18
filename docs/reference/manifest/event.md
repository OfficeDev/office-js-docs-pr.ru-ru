---
title: Элемент Event в файле манифеста
description: Определяет обработчик событий в надстройке.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 02037a54ad4b7e91a3697b53b04fa30e8a4909a9
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718232"
---
# <a name="event-element"></a>Элемент Event

Определяет обработчик событий в надстройке.

> [!NOTE] 
> В `Event` настоящее время элемент поддерживается только в Outlook в Интернете в Office 365.

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

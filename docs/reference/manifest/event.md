---
title: Элемент события в файле манифеста
description: Определяет обработчик событий в надстройке.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 095023a8f2d8cd5a01835e09cd50ae7289c98c01
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990665"
---
# <a name="event-element"></a>Элемент Event

Определяет обработчик событий в надстройке.

> [!NOTE]
> Сведения о поддержке и использовании см. в сайте [On-send feature for Outlook надстройки.](../../outlook/outlook-on-send-addins.md)

**Тип надстройки:** почтовая

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Type](#type-attribute)  |  Да  | Задает обрабатываемое событие. |
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

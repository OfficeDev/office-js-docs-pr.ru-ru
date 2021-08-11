---
title: Элемент Script в файле манифеста
description: Элемент Script определяет параметры скрипта, которые настраиваемая функция использует в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51902864081e135faed778de1bc6fdee15d67490de8eabc9febf493cb0c09889
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095047"
---
# <a name="script-element"></a>Элемент Script

Определяет параметры сценариев, используемых пользовательской функцией в Excel.

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|Элементы  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса файла JavaScript, используемого пользовательскими функциями.|

## <a name="example"></a>Пример

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```

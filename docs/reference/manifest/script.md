---
title: Элемент Script в файле манифеста
description: Элемент Script определяет параметры скрипта, которые настраиваемая функция использует в Excel.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 259976f752cf3fca72c5012bedd92b9bf021f6aa
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990672"
---
# <a name="script-element"></a>Элемент Script

Определяет параметры сценариев, используемых пользовательской функцией в Excel.

**Тип надстройки:** Настраиваемая функция

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

---
title: Элемент TargetDialect в файле манифеста
description: Элемент TargetDialect определяет региональный язык, поддерживаемый этим словарем, представленный в качестве строки имени культуры.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d0f60989ee5375f356343a8b3495f9c84120d467
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938998"
---
# <a name="targetdialect-element"></a>Элемент TargetDialect

Задает поддерживаемый этим словарем региональный язык в виде строки с названием языка и региональных параметров.

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<TargetDialect>
   string 
</TargetDialect>
```

## <a name="contained-in"></a>Содержится в

[TargetDialects](targetdialects.md)

## <a name="remarks"></a>Замечания

Укажите значение в формате языковых тегов BCP 47, например `en-US`.

## <a name="see-also"></a>См. также

- [Создание надстройки области задач словаря](../../word/dictionary-task-pane-add-ins.md)

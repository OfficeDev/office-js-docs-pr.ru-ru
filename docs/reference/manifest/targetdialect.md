---
title: Элемент TargetDialect в файле манифеста
description: Элемент TargetDialect определяет региональный язык, поддерживаемый этим словарем, представленный в виде строки имени языка и региональных параметров.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ba5c43b6471f11d7599da8542c30618ea1de78e0
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720332"
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

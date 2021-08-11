---
title: Элемент TargetDialect в файле манифеста
description: Элемент TargetDialect определяет региональный язык, поддерживаемый этим словарем, представленный в качестве строки имени культуры.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 04f78be33f215fc79abbcd52be716036f4369fc8cb6de59e2a725cc5228334c0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095563"
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

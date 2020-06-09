---
title: Элемент TargetDialect в файле манифеста
description: Элемент TargetDialect определяет региональный язык, поддерживаемый этим словарем, представленный в виде строки имени языка и региональных параметров.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d0f60989ee5375f356343a8b3495f9c84120d467
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609015"
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

---
title: Элемент Екуивалентаддин в файле манифеста
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356912"
---
# <a name="equivalentaddin-element"></a>Элемент Екуивалентаддин

Задает обратную совместимость для эквивалентной надстройки COM или XLL.

**Тип надстройки:** Область задач, настраиваемая функция

## <a name="syntax"></a>Синтаксис

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Содержится в

[Екуивалентадд](equivalentaddins.md)

## <a name="must-contain"></a>Должен содержать

[Type](type.md)

## <a name="can-contain"></a>Может содержать

[](progid.md)
[Имя файла](filename.md) ProgID

## <a name="remarks"></a>Примечания

Чтобы указать надстройку COM в качестве эквивалентной надстройки, укажите оба `ProgID` `Type` элемента:. Чтобы указать XLL в качестве эквивалентной надстройки, укажите оба `FileName` `Type` элемента:

## <a name="see-also"></a>См. также

- [Обеспечение совместимости пользовательских функций с пользовательскими функциями XLL](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Обеспечение совместимости надстройки Office с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
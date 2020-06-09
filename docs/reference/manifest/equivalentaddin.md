---
title: Элемент Екуивалентаддин в файле манифеста
description: Задает обратную совместимость для эквивалентной надстройки COM или XLL.
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: e14fe91bf7a5fe321019acf205ddb1753fedd569
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611563"
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

[Тип](type.md)

## <a name="can-contain"></a>Может содержать

[ProgID](progid.md) 
 [Имя файла](filename.md)

## <a name="remarks"></a>Примечания

Чтобы указать надстройку COM в качестве эквивалентной надстройки, укажите оба `ProgId` `Type` элемента:. Чтобы указать XLL в качестве эквивалентной надстройки, укажите оба `FileName` `Type` элемента:

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Обеспечение совместимости надстройки Excel с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
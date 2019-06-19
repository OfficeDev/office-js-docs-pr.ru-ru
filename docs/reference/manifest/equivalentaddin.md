---
title: Элемент Екуивалентаддин в файле манифеста
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059925"
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

[](progid.md)
[Имя файла](filename.md) ProgID

## <a name="remarks"></a>Примечания

Чтобы указать надстройку COM в качестве эквивалентной надстройки, укажите оба `ProgId` `Type` элемента:. Чтобы указать XLL в качестве эквивалентной надстройки, укажите оба `FileName` `Type` элемента:

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Обеспечение совместимости надстройки Excel с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
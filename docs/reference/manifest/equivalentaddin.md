---
title: Элемент EquivalentAddin в файле манифеста
description: Указывает обратную совместимость для эквивалентной надстройки COM или XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 412a3ce7bd12d886b7b88b5b84938e28295aba5d
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836839"
---
# <a name="equivalentaddin-element"></a>Элемент EquivalentAddin

Указывает обратную совместимость для эквивалентной надстройки COM или XLL.

**Тип надстройки:** Области задач, настраиваемая функция

## <a name="syntax"></a>Синтаксис

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Содержится в

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>Должен содержать

[Тип](type.md)

## <a name="can-contain"></a>Может содержать

[ProgId](progid.md) 
 [FileName](filename.md)

## <a name="remarks"></a>Примечания

Чтобы указать надстройки COM в качестве эквивалентной надстройки, укажите как элементы, так `ProgId` `Type` и элементы. Чтобы указать XLL в качестве эквивалентной надстройки, укажите как элементы, так `FileName` `Type` и элементы.

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
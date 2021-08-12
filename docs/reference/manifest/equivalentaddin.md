---
title: Элемент EquivalentAddin в файле манифеста
description: Указывает обратную совместимость для эквивалентной надстройки COM или XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 6404ad561a14a451e4685cc23be930b7ba612e85d1b37e78aa45f9366becf3bc
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085765"
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
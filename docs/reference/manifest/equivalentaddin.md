---
title: Элемент EquivalentAddin в файле манифеста
description: Указывает обратную совместимость для эквивалентной надстройки COM или XLL.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: f77a70681c8a12674d9e22022276e511552861ad
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990693"
---
# <a name="equivalentaddin-element"></a>Элемент EquivalentAddin

Указывает обратную совместимость для эквивалентной надстройки COM или XLL.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Тип надстройки:** Области задач, почты, настраиваемой функции

## <a name="syntax"></a>Синтаксис

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Содержится в

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>Должен содержать

[Type](type.md)

## <a name="can-contain"></a>Может содержать

[ProgId](progid.md) 
 [FileName](filename.md)

## <a name="remarks"></a>Комментарии

Чтобы указать надстройки COM в качестве эквивалентной надстройки, укажите как элементы, так `ProgId` `Type` и элементы. Чтобы указать XLL в качестве эквивалентной надстройки, укажите как элементы, так `FileName` `Type` и элементы.

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
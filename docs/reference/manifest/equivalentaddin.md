---
title: Элемент EquivalentAddin в файле манифеста
description: Указывает обратную совместимость для эквивалентной надстройки COM или XLL.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: e318a9028ebefdeca9aaf5baac465a1ec1af0a73
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042135"
---
# <a name="equivalentaddin-element"></a>Элемент EquivalentAddin

Указывает обратную совместимость для эквивалентной надстройки COM или XLL.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Тип надстройки:** Области задач, почты, настраиваемой функции

**Допустимо только в этих схемах VersionOverrides:**

- Область задач 1.0
- Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

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

## <a name="remarks"></a>Замечания

Чтобы указать надстройки COM в качестве эквивалентной надстройки, укажите как элементы, так `ProgId` `Type` и элементы. Чтобы указать XLL в качестве эквивалентной надстройки, укажите как элементы, так `FileName` `Type` и элементы.

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
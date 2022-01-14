---
title: Элемент EquivalentAddins в файле манифеста
description: Указывает обратную совместимость с эквивалентной надстройки COM, XLL или обоих.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48f3ef86f71ad3d4f0c759df4583af4cd95e5c5a
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042156"
---
# <a name="equivalentaddins-element"></a>Элемент EquivalentAddins

Указывает обратную совместимость с эквивалентной надстройки COM, XLL или обоих.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Тип надстройки:** Области задач, почты, настраиваемой функции

**Допустимо только в этих схемах VersionOverrides:**

- Область задач 1.0
- Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

## <a name="syntax"></a>Синтаксис

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## <a name="contained-in"></a>Содержится в

[VersionOverrides](versionoverrides.md)

## <a name="must-contain"></a>Должен содержать

[EquivalentAddin](equivalentaddin.md)

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
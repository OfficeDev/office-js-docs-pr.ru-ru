---
title: Элемент EquivalentAddins в файле манифеста
description: Указывает обратную совместимость с эквивалентной надстройки COM, XLL или обоих.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: d32f67f49d334a75433aec2d079b45a44a04121a
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990812"
---
# <a name="equivalentaddins-element"></a>Элемент EquivalentAddins

Указывает обратную совместимость с эквивалентной надстройки COM, XLL или обоих.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Тип надстройки:** Области задач, почты, настраиваемой функции

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
---
title: Элемент Type в файле манифеста
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356921"
---
# <a name="type-element"></a>Элемент Type

Указывает, является ли эквивалентная надстройка надстройкой COM или XLL.

**Тип надстройки:** Область задач, настраиваемая функция

## <a name="syntax"></a>Синтаксис

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Содержится в

[Екуивалентадд](equivalentaddin.md)

## <a name="add-in-type-values"></a>Значения типов надстроек

Необходимо указать одно из следующих значений для `Type` элемента.

- COM: эквивалентная надстройка — это надстройка COM.
- XLL: определяет эквивалентную надстройку в формате XLL.

## <a name="see-also"></a>См. также

- [Обеспечение совместимости пользовательских функций с пользовательскими функциями XLL](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Обеспечение совместимости надстройки Office с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
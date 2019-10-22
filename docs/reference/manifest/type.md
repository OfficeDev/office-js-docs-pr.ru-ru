---
title: Элемент Type в файле манифеста
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628230"
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

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Обеспечение совместимости надстройки Excel с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
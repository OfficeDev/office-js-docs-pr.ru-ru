---
title: Тип элемента в файле манифеста
description: Элемент Type указывает, является ли эквивалентная надстройка com надстройка или XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 5af3359c232e91b097311bfc06fc9b1c932b0703
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836811"
---
# <a name="type-element"></a>Элемент Type

Указывает, является ли эквивалентная надстройка com надстройка или XLL.

**Тип надстройки:** Области задач, настраиваемая функция

## <a name="syntax"></a>Синтаксис

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Содержится в

[EquivalentAddin](equivalentaddin.md)

## <a name="add-in-type-values"></a>Значения типа надстройки

Необходимо указать одно из следующих значений `Type` элемента.

- COM. Указывает, что эквивалентная надстройка — это надстройка COM.
- XLL. Указывает эквивалентную надстройка Excel XLL.

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
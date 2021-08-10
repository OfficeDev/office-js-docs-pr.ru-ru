---
title: Тип элемента в файле манифеста
description: Элемент Type указывает, является ли эквивалентная надстройка com надстройка или XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: ca6fa7183727870593dd3e726abc72fdc0d6f0b518fdb8451ec80c6b590f8c83
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092481"
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
- XLL. Указывает, что эквивалентная надстройка является Excel XLL.

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
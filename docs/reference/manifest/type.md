---
title: Тип элемента в файле манифеста
description: Элемент Type указывает, является ли эквивалентная надстройка com надстройка или XLL.
ms.date: 03/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 65860ff7aa3e241c227f96c8a8e7c71d7799a04c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154427"
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
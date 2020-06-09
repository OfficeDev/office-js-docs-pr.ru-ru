---
title: Элемент Type в файле манифеста
description: Элемент Type указывает, является ли эквивалентная надстройка надстройкой COM или XLL.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: b59f903af39facd7543e7384189817d5365cf8c9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604561"
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
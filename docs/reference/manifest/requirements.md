---
title: Элемент Requirements в файле манифеста
description: Элемент указывает минимальный набор обязательных требований и методы, необходимые надстройке Office для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c6a9a7b5923401fc2551f239b2c6cbc0d1e90755
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641321"
---
# <a name="requirements-element"></a>Элемент Requirements

Указывает минимальный набор требований к API JavaScript для Office ([набор требований](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) и/или методов), которые должна активировать надстройка Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Методы](methods.md)|x||x|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

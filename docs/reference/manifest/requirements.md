---
title: Элемент Requirements в файле манифеста
description: Элемент Requirements указывает минимальный набор требований и методы, необходимые Office надстройки для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3020037b48e3f759acf6a7e2758bb8c1fd2dd36429e0b21613e22fca33cacc1a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098109"
---
# <a name="requirements-element"></a>Элемент Requirements

Указывает минимальный набор Office API JavaScript[(наборы](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) требований и/или методы), который необходимо активировать Office надстройки.

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

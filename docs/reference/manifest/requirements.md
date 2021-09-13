---
title: Элемент Requirements в файле манифеста
description: Элемент Requirements указывает минимальный набор требований и методы, необходимые Office надстройки для активации.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 3a5a393485094b5cc830b5120c3abd8c211eff1e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154074"
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

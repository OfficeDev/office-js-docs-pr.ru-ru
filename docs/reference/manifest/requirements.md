---
title: Элемент Requirements в файле манифеста
description: Элемент указывает минимальный набор обязательных требований и методы, необходимые надстройке Office для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 586f05ec68257462cb64a96abf2a34eb31861a5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611717"
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

|**Элемент**|**Content**|**Почтовая надстройка**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Методы](methods.md)|x||x|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

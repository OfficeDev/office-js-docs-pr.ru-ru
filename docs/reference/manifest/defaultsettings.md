---
title: Элемент DefaultSettings в файле манифеста
description: Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 11e398d86a702f4e45a5cea7b63e0380ce65d1749d0660789e96477744d73079
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095925"
---
# <a name="defaultsettings-element"></a>Элемент DefaultSettings

Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.

**Тип надстройки:** контентные надстройки и надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Замечания

Расположение источника и другие параметры в **элементе DefaultSettings** применяются только к надстройке контента и области задач. Для надстройок почты в элементе [FormSettings](formsettings.md) указывается расположение исходных файлов по умолчанию и другие параметры по умолчанию.

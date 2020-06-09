---
title: Элемент DefaultSettings в файле манифеста
description: Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ace4f971d342f98d0aca5c21a7a48ceaf2563a2f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611584"
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

|**Элемент**|**Content**|**Почтовая надстройка**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Замечания

Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к контентным надстройкам и надстройкам области задач. Для почтовых надстроек указываются расположения по умолчанию для исходных файлов и другие параметры по умолчанию в элементе [FormSettings](formsettings.md) .


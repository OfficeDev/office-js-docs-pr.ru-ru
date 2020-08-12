---
title: Элемент DefaultSettings в файле манифеста
description: Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a9711fb44390bcbda8979b8018eed1318c5579bc
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641468"
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

Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к контентным надстройкам и надстройкам области задач. Для почтовых надстроек указываются расположения по умолчанию для исходных файлов и другие параметры по умолчанию в элементе [FormSettings](formsettings.md) .

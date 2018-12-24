---
title: Элемент DefaultSettings в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0c109d5d893cf9d3502f1cbf1724007f01e623e6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433757"
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

|**Element**|**Контентная надстройка**|**Почтовая надстройка**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Замечания

Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к надстройкам области задач и контентным надстройкам. В случае почтовых надстроек следует задавать расположения по умолчанию для исходных файлов и другие стандартные параметры с помощью элемента [FormSettings](formsettings.md).


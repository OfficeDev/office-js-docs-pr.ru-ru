---
title: Элемент Methods в файле манифеста
description: Элемент Methods указывает список методов API Office JavaScript, которые требуются Office надстройки для активации Office или переопределения параметров базового манифеста.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4c39c6363cd33e103cf40c0f7f047fa694db1411
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222278"
---
# <a name="methods-element"></a>Элемент Methods

Значение этого элемента зависит от того, где он используется в манифесте.

## <a name="in-the-base-manifest"></a>В базовом манифесте

Когда используется в базовом манифесте  (то есть элемент родительских требований является прямым ребенком [OfficeApp),](officeapp.md)элемент **Methods** указывает список методов API javaScript Office, которые необходимы вашей Office надстройки для активации Office.

**Тип надстройки:** контентные надстройки и надстройки области задач.

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>Как внук элемента VersionOverrides

Указывает минимальный набор методов API Office JavaScript, которые должны поддерживаться версией и платформой Office (например, Windows, Mac, web и iOS или iPad) для того, чтобы [версияOverrides](versionoverrides.md) вступила в силу.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides:**

- То же самое, что и элемент [родительских требований.](requirements.md)

**Связанные с этими наборами требований:**

- То же самое, что и элемент [родительских требований.](requirements.md)

## <a name="syntax"></a>Синтаксис

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a>Содержится в

[Requirements](requirements.md)

## <a name="can-contain"></a>Может содержать

[Метод](method.md)

## <a name="remarks"></a>Замечания

Элементы **Methods** и **Method** не поддерживаются в надстройки почты при их использования в базовом манифесте. Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

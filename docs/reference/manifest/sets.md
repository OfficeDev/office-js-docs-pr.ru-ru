---
title: Элемент Sets в файле манифеста
description: Элемент Sets указывает минимальный набор API Office JavaScript, который требуется Office надстройки для активации Office или переопределения параметров базового манифеста.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: df0cf686fe213a51321595a000438ca2a411f2c7
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222145"
---
# <a name="sets-element"></a>Элемент Sets

Значение этого элемента зависит от того, где он используется в манифесте.

## <a name="in-the-base-manifest"></a>В базовом манифесте

Когда используется в базовом манифесте  (то есть элемент родительских требований является прямым ребенком [OfficeApp),](officeapp.md)элемент **Sets** указывает минимальный подмножество требований API javaScript [Office](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)(наборы требований), которые необходимы вашей Office надстройки для активации Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>Как внук элемента VersionOverrides

Указывает минимальный набор Office API JavaScript[(наборы](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)требований), которые должны поддерживаться версией и платформой Office (например, Windows, Mac, web и iOS или iPad) для того, чтобы [ВерсияOverrides](versionoverrides.md) вступила в силу.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides:**

- То же самое, что и элемент [родительских требований.](requirements.md)

**Связанные с этими наборами требований:**

- То же самое, что и элемент [родительских требований.](requirements.md)

## <a name="syntax"></a>Синтаксис

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Содержится в

[Requirements](requirements.md)

## <a name="can-contain"></a>Может содержать

[Set](set.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|необязательный|Указывает значение атрибута **MinVersion** по умолчанию для всех [элементов](set.md) набора детей. Значение по умолчанию: "1.1".|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Дополнительные сведения о атрибуте **MinVersion** элемента **Set** и **атрибуте DefaultMinVersion** элемента **Sets** см. в Office версиях и платформах надстройки. [](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)


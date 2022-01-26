---
title: Элемент Set в файле манифеста
description: Элемент Set указывает Office API JavaScript, задаваемого Office надстройки для активации Office или переопределения параметров базового манифеста.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 55e1b25765bfbe53108bc9201c0c851c6ef9161d
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222236"
---
# <a name="set-element"></a>Элемент Set

Значение этого элемента зависит от того, где он используется в манифесте.

## <a name="in-the-base-manifest"></a>В базовом манифесте

Когда используется в базовом манифесте  (то есть элемент требования к дедушке является прямым [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) ребенком [OfficeApp),](officeapp.md)элемент Set указывает набор требований из API JavaScript Office, необходимый вашему Office надстройке для активации Office. 

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>Как правнук элемента VersionOverrides

Указывает набор [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) требований из API Office JavaScript, который должен поддерживаться версией и платформой Office (например, Windows, Mac, web и iOS или iPad) для того, чтобы [версияOverrides](versionoverrides.md) вступила в силу.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides:**

- То же, что и элемент ["Требования к бабушке и дедушке".](requirements.md)

**Связанные с этими наборами требований:**

- То же, что и элемент ["Требования к бабушке и дедушке".](requirements.md)

## <a name="syntax"></a>Синтаксис

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Содержится в

[Sets](sets.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Имя [набора требований](../../develop/office-versions-and-requirement-sets.md).|
|MinVersion|string|необязательный|Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion,** если оно указано в элементе родительских [наборов.](sets.md)|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Дополнительные сведения о атрибуте **MinVersion** элемента **Set** и **атрибуте DefaultMinVersion** элемента **Sets** см. в Office версиях и платформах надстройки. [](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)


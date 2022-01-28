---
title: Элемент Requirements в файле манифеста
description: Элемент Requirements указывает минимальный набор требований и методы, Office надстройка должна быть активирована Office или переопределять базовые параметры манифеста.
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e7953ca1e47c492849fe9d0c79384376ffdec347
ms.sourcegitcommit: e837f966d7360ed11b3ff9363ff20380f7d0c45e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/28/2022
ms.locfileid: "62263039"
---
# <a name="requirements-element"></a>Элемент Requirements

Значение этого элемента зависит от того, используется ли он в базовом манифесте или как ребенок элемента [**VersionOverrides**](#as-a-child-of-a-versionoverrides-element).[](#in-the-base-manifest)

> [!TIP]
> Перед использованием этого элемента ознакомьтесь с требованиями [Office и API](../../develop/specify-office-hosts-and-api-requirements.md)

## <a name="in-the-base-manifest"></a>В базовом манифесте

Элемент **Requirements**, используемый в базовом манифесте (то есть как прямой ребенок [OfficeApp](officeapp.md)), указывает минимальный набор требований Office API JavaScript [(](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)наборы требований и/или методы), который необходимо активировать Office надстройки Office. Надстройка не будет активирована на любом сочетании Office версии и платформы (например, Windows, Mac, web и iOS или iPad), которое не поддерживает указанные методы и наборы требований.

**Тип надстройки:** Области задач, Почта

## <a name="as-a-child-of-a-versionoverrides-element"></a>Как ребенок элемента VersionOverrides

Если используется в качестве ребенка [VersionOverrides](versionoverrides.md), указывает минимальный набор требований Office API JavaScript [(наборы](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) требований и/или методы), которые должны поддерживаться версией и платформой Office (например, Windows, Mac, web и iOS или iPad) для параметров элемента **VersionOverrides**, переопределяющие базовые параметры манифеста.  вступает в силу.

Рассмотрим надстройка, которая указывает требование A в базовом манифесте и указывает требование B внутри **VersionOverrides**. 

- Если платформа и Office не поддерживают A, надстройка не активируется и Office не размыкает раздел **VersionOverrides** манифеста. 
- Если поддерживается как A, так и B, надстройка активируется и вся разметка **в VersionOverrides** вступает в силу. 
- Если A поддерживается, а B — нет, то надстройка активируется и часть  разметки **в VersionOverrides** вступает в силу. В частности, вступает в силу детский элемент **VersionOverrides** , не переопределяющий базовые элементы манифеста. Например, вступает **в силу элемент WebApplicationInfo** или **equivalentAddins** . Однако все детские элементы **VersionOverrides** , переопределяющие базовый элемент манифеста, например **Хосты**, не вступает в силу. Вместо этого Office использует значения базовой разметки манифеста, которые в противном случае были бы переопределены. 

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) , когда родительский **VersionOverrides** — это тип Taskpane 1.0.
- [Почтовый ящик 1.3,](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) когда родительский **VersionOverrides** — это тип Почта 1.0.
- [Почтовый ящик 1.5,](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) когда родительский **VersionOverrides** — это тип Почта 1.1.

### <a name="remarks"></a>Комментарии

Элемент **Requirements** не служит никакой цели в **VersionOverrides**, если он не указывает дополнительных требований, не указанных в требованиях в  базовом манифесте. Если Office версия и платформа не поддерживают требования базового манифеста, надстройка не активируется, а элемент **VersionOverrides** не размыкается. По этой причине элемент **Requirements** следует использовать в **VersionOverrides** только при условии:

- У вашей надстройки есть дополнительные функции, реализованные с конфигурацией **в VersionOverrides** (например, команды надстройки), которые требуют набора методов или требований, которые не указаны в элементе **Requirements** в базовом манифесте.
- Надстройка полезна и должна активироваться (но без дополнительных функций), даже в сочетании платформы и Office версии, которая не поддерживает требования, необходимые для дополнительных функций.

> [!TIP]
> Не **повторяйте элементы Требования** из базового манифеста внутри **VersionOverrides**. Это не влияет на назначение элемента **Requirements** в **VersionOverrides**.

> [!WARNING]
> Прежде чем использовать элемент **Requirements** в **VersionOverrides***, используйте* большую внимательность, так как на платформах и версиях, которые не поддерживают это требование, ни одна из команд надстройки не будет установлена *, даже* те, которые вызывают функциональность, которая не нуждается в этом требовании. Рассмотрим, например, надстройка, которая имеет две настраиваемые кнопки ленты. Один из них вызывает Office API JavaScript, доступные в наборе требований **ExcelApi 1.4** (и более поздней части). Другие вызовы API, доступные только в **ExcelApi 1.9** (и более поздней). Если вы поместите требование **для ExcelApi 1.9** в **VersionOverrides**, то если кнопка 1.9 не поддерживается, на ленте  не появится ни одна кнопка. Лучшей стратегией в этом сценарии было бы использование метода, описанного в проверках времени запуска для поддержки набора [методов и требований](../../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). Код, на который ссылается вторая кнопка, `isSetSupported` сначала используется для проверки поддержки **ExcelApi 1.9**. Если он не поддерживается, код дает пользователю сообщение о том, что эта функция надстройки недоступна в версии Office. 

> [!NOTE]
> В надстройки Mail можно вложить **в VersionOverrides** 1.1 внутри **VersionOverrides** 1.0. Office всегда будет использовать самые высокие версии **VersionOverrides**, поддерживаемые платформой и Office версией.

## <a name="syntax"></a>Синтаксис

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)
 [VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>Может содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Методы](methods.md)|x||x|

## <a name="see-also"></a>См. также

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

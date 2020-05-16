---
title: Элемент RequestedHeight в файле манифеста
description: Элемент RequestedHeight указывает начальную высоту (в пикселях) надстройки для работы с контентом или почтовой надстройкой.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: fa40043e6192e1304e67f1f96f770898b230036c
ms.sourcegitcommit: b634bfe9a946fbd95754e87f070a904ed57586ff
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/15/2020
ms.locfileid: "44253616"
---
# <a name="requestedheight-element"></a>Элемент RequestedHeight

Указывает исходную высоту окна (в пикселях) контентной или почтовой надстройки

**Тип надстройки**: контентная, почтовая

## <a name="syntax"></a>Синтаксис

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Содержится в

- [DefaultSettings](defaultsettings.md) (контентные надстройки) со значением в диапазоне от 32 до 1000
- [DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (почтовые надстройки) со значением в диапазоне от 32 до 450
- [ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки) со значением, которое может находиться в диапазоне от 140 до 450 точки расширения **DetectedEntity** и между 32 и 450 для [точки расширения **кустомпане** (не рекомендуется)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)

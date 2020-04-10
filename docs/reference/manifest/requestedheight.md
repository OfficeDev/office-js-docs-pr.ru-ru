---
title: Элемент RequestedHeight в файле манифеста
description: Элемент RequestedHeight указывает начальную высоту (в пикселях) надстройки для работы с контентом или почтовой надстройкой.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5f4c3ca1ff39cc3150249fbc824b0db76f6b8a85
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215042"
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
- [ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки) со значением в диапазоне от 140 до 450 для точки расширения **DetectedEntity** и в диапазоне от 32 до 450 для точки расширения **CustomPane**

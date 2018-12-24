---
title: Элемент RequestedHeight в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ea8c0403146f526b28eb20b8364fd210ac357baf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433476"
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
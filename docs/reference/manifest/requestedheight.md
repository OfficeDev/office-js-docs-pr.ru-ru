---
title: Элемент RequestedHeight в файле манифеста
description: Элемент RequestedHeight указывает начальную высоту (в пикселях) контента или надстройки почты.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 0e5f9de909d32622ac244ff4118c8a3192abf2ff0fe89ed81a6188ddcb265549
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092989"
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
- [ExtensionPoint](extensionpoint.md) (надстройки контекстной почты) со значением от 140 до 450 для точки расширения **DetectedEntity** и от 32 до 450 для точки расширения [ **CustomPane** (неподготовленной)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)

---
title: Элемент RequestedHeight в файле манифеста
description: Элемент RequestedHeight указывает начальную высоту (в пикселях) контента или надстройки почты.
ms.date: 05/14/2020
ms.localizationpriority: medium
ms.openlocfilehash: e0589e81e8905c4fc8c7a8e50ec7c14038035677
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151329"
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

---
title: Элемент RequestedHeight в файле манифеста
description: Элемент RequestedHeight указывает начальную высоту (в пикселях) надстройки для работы с контентом или почтовой надстройкой.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 853d12baf290167f3e6a635201e8b5d1d0e35a51
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720458"
---
# <a name="requestedheight-element"></a><span data-ttu-id="accc7-103">Элемент RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="accc7-103">RequestedHeight element</span></span>

<span data-ttu-id="accc7-104">Указывает исходную высоту окна (в пикселях) контентной или почтовой надстройки</span><span class="sxs-lookup"><span data-stu-id="accc7-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="accc7-105">**Тип надстройки**: контентная, почтовая</span><span class="sxs-lookup"><span data-stu-id="accc7-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="accc7-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="accc7-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="accc7-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="accc7-107">Contained in</span></span>

- <span data-ttu-id="accc7-108">[DefaultSettings](defaultsettings.md) (контентные надстройки) со значением в диапазоне от 32 до 1000</span><span class="sxs-lookup"><span data-stu-id="accc7-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="accc7-109">[DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (почтовые надстройки) со значением в диапазоне от 32 до 450</span><span class="sxs-lookup"><span data-stu-id="accc7-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="accc7-110">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки) со значением в диапазоне от 140 до 450 для точки расширения **DetectedEntity** и в диапазоне от 32 до 450 для точки расширения **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="accc7-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>

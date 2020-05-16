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
# <a name="requestedheight-element"></a><span data-ttu-id="4eecf-103">Элемент RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="4eecf-103">RequestedHeight element</span></span>

<span data-ttu-id="4eecf-104">Указывает исходную высоту окна (в пикселях) контентной или почтовой надстройки</span><span class="sxs-lookup"><span data-stu-id="4eecf-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="4eecf-105">**Тип надстройки**: контентная, почтовая</span><span class="sxs-lookup"><span data-stu-id="4eecf-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4eecf-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="4eecf-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="4eecf-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="4eecf-107">Contained in</span></span>

- <span data-ttu-id="4eecf-108">[DefaultSettings](defaultsettings.md) (контентные надстройки) со значением в диапазоне от 32 до 1000</span><span class="sxs-lookup"><span data-stu-id="4eecf-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="4eecf-109">[DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (почтовые надстройки) со значением в диапазоне от 32 до 450</span><span class="sxs-lookup"><span data-stu-id="4eecf-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="4eecf-110">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки) со значением, которое может находиться в диапазоне от 140 до 450 точки расширения **DetectedEntity** и между 32 и 450 для [точки расширения **кустомпане** (не рекомендуется)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span><span class="sxs-lookup"><span data-stu-id="4eecf-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the [**CustomPane** extension point (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span></span>

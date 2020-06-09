---
title: Элемент RequestedHeight в файле манифеста
description: Элемент RequestedHeight указывает начальную высоту (в пикселях) надстройки для работы с контентом или почтовой надстройкой.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 44675918a4208683f442fe8a6e8f4f906f484571
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611731"
---
# <a name="requestedheight-element"></a><span data-ttu-id="52d7e-103">Элемент RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="52d7e-103">RequestedHeight element</span></span>

<span data-ttu-id="52d7e-104">Указывает исходную высоту окна (в пикселях) контентной или почтовой надстройки</span><span class="sxs-lookup"><span data-stu-id="52d7e-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="52d7e-105">**Тип надстройки**: контентная, почтовая</span><span class="sxs-lookup"><span data-stu-id="52d7e-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="52d7e-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="52d7e-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="52d7e-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="52d7e-107">Contained in</span></span>

- <span data-ttu-id="52d7e-108">[DefaultSettings](defaultsettings.md) (контентные надстройки) со значением в диапазоне от 32 до 1000</span><span class="sxs-lookup"><span data-stu-id="52d7e-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="52d7e-109">[DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (почтовые надстройки) со значением в диапазоне от 32 до 450</span><span class="sxs-lookup"><span data-stu-id="52d7e-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="52d7e-110">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки) со значением, которое может находиться в диапазоне от 140 до 450 точки расширения **DetectedEntity** и между 32 и 450 для [точки расширения **кустомпане** (не рекомендуется)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span><span class="sxs-lookup"><span data-stu-id="52d7e-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the [**CustomPane** extension point (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span></span>

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
# <a name="requestedheight-element"></a><span data-ttu-id="09240-103">Элемент RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="09240-103">RequestedHeight element</span></span>

<span data-ttu-id="09240-104">Указывает исходную высоту окна (в пикселях) контентной или почтовой надстройки</span><span class="sxs-lookup"><span data-stu-id="09240-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="09240-105">**Тип надстройки**: контентная, почтовая</span><span class="sxs-lookup"><span data-stu-id="09240-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="09240-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="09240-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="09240-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="09240-107">Contained in</span></span>

- <span data-ttu-id="09240-108">[DefaultSettings](defaultsettings.md) (контентные надстройки) со значением в диапазоне от 32 до 1000</span><span class="sxs-lookup"><span data-stu-id="09240-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="09240-109">[DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (почтовые надстройки) со значением в диапазоне от 32 до 450</span><span class="sxs-lookup"><span data-stu-id="09240-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="09240-110">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки) со значением в диапазоне от 140 до 450 для точки расширения **DetectedEntity** и в диапазоне от 32 до 450 для точки расширения **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="09240-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>

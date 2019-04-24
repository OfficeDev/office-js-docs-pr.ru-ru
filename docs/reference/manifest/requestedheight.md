---
title: Элемент RequestedHeight в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e175d9012bb2f2a42fd466c35e5e28ade967d6f2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450529"
---
# <a name="requestedheight-element"></a><span data-ttu-id="3f8ee-102">Элемент RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="3f8ee-102">RequestedHeight element</span></span>

<span data-ttu-id="3f8ee-103">Указывает исходную высоту окна (в пикселях) контентной или почтовой надстройки</span><span class="sxs-lookup"><span data-stu-id="3f8ee-103">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="3f8ee-104">**Тип надстройки**: контентная, почтовая</span><span class="sxs-lookup"><span data-stu-id="3f8ee-104">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3f8ee-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="3f8ee-105">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="3f8ee-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="3f8ee-106">Contained in</span></span>

- <span data-ttu-id="3f8ee-107">[DefaultSettings](defaultsettings.md) (контентные надстройки) со значением в диапазоне от 32 до 1000</span><span class="sxs-lookup"><span data-stu-id="3f8ee-107">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="3f8ee-108">[DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (почтовые надстройки) со значением в диапазоне от 32 до 450</span><span class="sxs-lookup"><span data-stu-id="3f8ee-108">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="3f8ee-109">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки) со значением в диапазоне от 140 до 450 для точки расширения **DetectedEntity** и в диапазоне от 32 до 450 для точки расширения **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="3f8ee-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>

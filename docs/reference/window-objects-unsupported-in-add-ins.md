---
title: Объекты Window, которые не поддерживаются в надстройках Office
description: В этой статье указаны некоторые объекты среды выполнения, которые не работают в надстройках Office.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160507"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a><span data-ttu-id="2d3ba-103">Объекты Window, которые не поддерживаются в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="2d3ba-103">Window objects that are unsupported in Office Add-ins</span></span>

<span data-ttu-id="2d3ba-104">В некоторых версиях Windows и Office надстройки запускаются в среде выполнения Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="2d3ba-104">For some versions of Windows and Office, add-ins run in an Internet Explorer 11 runtime.</span></span> <span data-ttu-id="2d3ba-105">(Дополнительные сведения см. в разделе [браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).) Некоторые свойства или вложенные свойства глобального `window` объекта не поддерживаются в Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="2d3ba-105">(For details, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Some properties or subproperties of the global `window` object are not supported in Internet Explorer 11.</span></span> <span data-ttu-id="2d3ba-106">Эти свойства отключены в надстройках, чтобы надстройка гарантированно соответствовала всем пользователям, независимо от того, какой браузер использует надстройка.</span><span class="sxs-lookup"><span data-stu-id="2d3ba-106">These properties are disabled in add-ins to ensure that your add-in provides a consistent experience to all users, regardless of which browser the add-in is using.</span></span> <span data-ttu-id="2d3ba-107">Это также способствует правильной загрузке AngularJS.</span><span class="sxs-lookup"><span data-stu-id="2d3ba-107">This also helps AngularJS load properly.</span></span>

<span data-ttu-id="2d3ba-108">Ниже приведен список отключенных свойств.</span><span class="sxs-lookup"><span data-stu-id="2d3ba-108">The following is a list of the disabled properties.</span></span> <span data-ttu-id="2d3ba-109">Список является выполняемой работой.</span><span class="sxs-lookup"><span data-stu-id="2d3ba-109">The list is a work in progress.</span></span> <span data-ttu-id="2d3ba-110">Если вы обнаружите дополнительные `window` свойства, которые не работают в надстройках, воспользуйтесь средством обратной связи, приведенным ниже, чтобы сообщить нам об этом.</span><span class="sxs-lookup"><span data-stu-id="2d3ba-110">If you discover additional `window` properties that do not work in add-ins, please use the feedback tool below to tell us.</span></span>

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a><span data-ttu-id="2d3ba-111">См. также</span><span class="sxs-lookup"><span data-stu-id="2d3ba-111">See also</span></span>

- [<span data-ttu-id="2d3ba-112">Браузеры, используемые надстройками Office</span><span class="sxs-lookup"><span data-stu-id="2d3ba-112">Browsers used by Office Add-ins</span></span>](../concepts/browsers-used-by-office-web-add-ins.md)
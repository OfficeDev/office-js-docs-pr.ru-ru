---
title: Runtimes in the manifest file
description: Элемент Runtimes указывает времени работы надстройки.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: afbcc6a909c51d2ed56292ef1541193f7f698d28
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789165"
---
# <a name="runtimes-element"></a><span data-ttu-id="ea27d-103">Элемент Runtimes</span><span class="sxs-lookup"><span data-stu-id="ea27d-103">Runtimes element</span></span>

<span data-ttu-id="ea27d-104">Указывает времени работы надстройки.</span><span class="sxs-lookup"><span data-stu-id="ea27d-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="ea27d-105">Child of the [`<Host>`](host.md) element.</span><span class="sxs-lookup"><span data-stu-id="ea27d-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="ea27d-106">При запуске в Office для Windows надстройка использует браузер Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="ea27d-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="ea27d-107">В Excel этот элемент позволяет ленте, области задач и пользовательским функциям использовать ту же времени работы.</span><span class="sxs-lookup"><span data-stu-id="ea27d-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="ea27d-108">Дополнительные сведения см. в настройках надстройки Excel для использования общей времени [работы JavaScript.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="ea27d-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="ea27d-109">В Outlook этот элемент включает активацию надстройки на основе событий.</span><span class="sxs-lookup"><span data-stu-id="ea27d-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="ea27d-110">Дополнительные сведения см. в настройке [надстройки Outlook для активации на основе событий.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="ea27d-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="ea27d-111">**Тип надстройки:** Области задач, почта</span><span class="sxs-lookup"><span data-stu-id="ea27d-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ea27d-112">**Outlook**: функция активации на [](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) основе событий в настоящее время находится в предварительной версии и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="ea27d-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="ea27d-113">Дополнительные сведения см. в [предварительном просмотре функции активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="ea27d-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="ea27d-114">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="ea27d-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="ea27d-115">Содержится в</span><span class="sxs-lookup"><span data-stu-id="ea27d-115">Contained in</span></span>

[<span data-ttu-id="ea27d-116">Host</span><span class="sxs-lookup"><span data-stu-id="ea27d-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="ea27d-117">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ea27d-117">Child elements</span></span>

|  <span data-ttu-id="ea27d-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="ea27d-118">Element</span></span> |  <span data-ttu-id="ea27d-119">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ea27d-119">Required</span></span>  |  <span data-ttu-id="ea27d-120">Описание</span><span class="sxs-lookup"><span data-stu-id="ea27d-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="ea27d-121">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="ea27d-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="ea27d-122">Да</span><span class="sxs-lookup"><span data-stu-id="ea27d-122">Yes</span></span> |  <span data-ttu-id="ea27d-123">Время работы надстройки.</span><span class="sxs-lookup"><span data-stu-id="ea27d-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ea27d-124">См. также</span><span class="sxs-lookup"><span data-stu-id="ea27d-124">See also</span></span>

- [<span data-ttu-id="ea27d-125">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="ea27d-125">Runtime</span></span>](runtime.md)

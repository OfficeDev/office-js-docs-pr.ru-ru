---
title: LaunchEvents в файле манифеста (предварительная версия)
description: Элемент LaunchEvents настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 9df059879018d79a61f1c900888c8d197e0b9880
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104814"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="660cb-103">Элемент LaunchEvents (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="660cb-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="660cb-104">Настраивает надстройки для активации на основе поддерживаемых событий.</span><span class="sxs-lookup"><span data-stu-id="660cb-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="660cb-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span><span class="sxs-lookup"><span data-stu-id="660cb-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="660cb-106">Дополнительные сведения см. в настройке [надстройки Outlook для активации на основе событий.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="660cb-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="660cb-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="660cb-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="660cb-108">Активация на основе событий в настоящее время находится [в предварительной](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) версии и доступна только в Outlook в Интернете и Windows.</span><span class="sxs-lookup"><span data-stu-id="660cb-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and Windows.</span></span> <span data-ttu-id="660cb-109">Дополнительные сведения см. в [предварительном просмотре функции активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="660cb-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="660cb-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="660cb-110">Syntax</span></span>

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a><span data-ttu-id="660cb-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="660cb-111">Contained in</span></span>

<span data-ttu-id="660cb-112">[ExtensionPoint](extensionpoint.md) ( Почтовая надстройка **LaunchEvent)**</span><span class="sxs-lookup"><span data-stu-id="660cb-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="660cb-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="660cb-113">Child elements</span></span>

|  <span data-ttu-id="660cb-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="660cb-114">Element</span></span> |  <span data-ttu-id="660cb-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="660cb-115">Required</span></span>  |  <span data-ttu-id="660cb-116">Описание</span><span class="sxs-lookup"><span data-stu-id="660cb-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="660cb-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="660cb-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="660cb-118">Да</span><span class="sxs-lookup"><span data-stu-id="660cb-118">Yes</span></span> |  <span data-ttu-id="660cb-119">Соейте поддерживаемые события с его функцией в файле JavaScript для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="660cb-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="660cb-120">См. также</span><span class="sxs-lookup"><span data-stu-id="660cb-120">See also</span></span>

- [<span data-ttu-id="660cb-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="660cb-121">LaunchEvent</span></span>](launchevent.md)

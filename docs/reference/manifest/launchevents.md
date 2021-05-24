---
title: LaunchEvents в файле манифеста
description: Элемент LaunchEvents настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 16d721ca6d9402d2bd5d19787707e146358044f0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590919"
---
# <a name="launchevents-element"></a><span data-ttu-id="8a517-103">Элемент LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="8a517-103">LaunchEvents element</span></span>

<span data-ttu-id="8a517-104">Настраивает надстройка для активации на основе поддерживаемых событий.</span><span class="sxs-lookup"><span data-stu-id="8a517-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="8a517-105">Ребенок [`<ExtensionPoint>`](extensionpoint.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="8a517-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="8a517-106">Дополнительные сведения см. в Outlook [надстройки](../../outlook/autolaunch.md)для активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="8a517-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="8a517-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="8a517-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8a517-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="8a517-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="8a517-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="8a517-109">Contained in</span></span>

<span data-ttu-id="8a517-110">[ExtensionPoint](extensionpoint.md) **(Надстройка для почты LaunchEvent)**</span><span class="sxs-lookup"><span data-stu-id="8a517-110">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="8a517-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="8a517-111">Child elements</span></span>

|  <span data-ttu-id="8a517-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="8a517-112">Element</span></span> |  <span data-ttu-id="8a517-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8a517-113">Required</span></span>  |  <span data-ttu-id="8a517-114">Описание</span><span class="sxs-lookup"><span data-stu-id="8a517-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="8a517-115">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="8a517-115">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="8a517-116">Да</span><span class="sxs-lookup"><span data-stu-id="8a517-116">Yes</span></span> |  <span data-ttu-id="8a517-117">Карта поддерживаемого события для его функции в файле JavaScript для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="8a517-117">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="8a517-118">См. также</span><span class="sxs-lookup"><span data-stu-id="8a517-118">See also</span></span>

- [<span data-ttu-id="8a517-119">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="8a517-119">LaunchEvent</span></span>](launchevent.md)

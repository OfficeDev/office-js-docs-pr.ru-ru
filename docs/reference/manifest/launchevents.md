---
title: Лаунчевентс в файле манифеста (Предварительная версия)
description: Элемент Лаунчевентс настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 92416f8c646326410a8cd9ee7831e17a5c5f1ffc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611773"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="5f9c9-103">Элемент Лаунчевентс (Preview)</span><span class="sxs-lookup"><span data-stu-id="5f9c9-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="5f9c9-104">Настраивает надстройку для активации на основе поддерживаемых событий.</span><span class="sxs-lookup"><span data-stu-id="5f9c9-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="5f9c9-105">Дочерний [`<ExtensionPoint>`](extensionpoint.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="5f9c9-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="5f9c9-106">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="5f9c9-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="5f9c9-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="5f9c9-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5f9c9-108">Активация на основе событий в настоящее время находится [в режиме предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="5f9c9-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="5f9c9-109">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="5f9c9-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="5f9c9-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="5f9c9-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="5f9c9-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="5f9c9-111">Contained in</span></span>

<span data-ttu-id="5f9c9-112">[ExtensionPoint](extensionpoint.md) (почтовые надстройки**лаунчевент** )</span><span class="sxs-lookup"><span data-stu-id="5f9c9-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="5f9c9-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="5f9c9-113">Child elements</span></span>

|  <span data-ttu-id="5f9c9-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="5f9c9-114">Element</span></span> |  <span data-ttu-id="5f9c9-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5f9c9-115">Required</span></span>  |  <span data-ttu-id="5f9c9-116">Описание</span><span class="sxs-lookup"><span data-stu-id="5f9c9-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="5f9c9-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="5f9c9-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="5f9c9-118">Да</span><span class="sxs-lookup"><span data-stu-id="5f9c9-118">Yes</span></span> |  <span data-ttu-id="5f9c9-119">Сопоставление поддерживаемого события с функцией в файле JavaScript для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="5f9c9-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="5f9c9-120">См. также</span><span class="sxs-lookup"><span data-stu-id="5f9c9-120">See also</span></span>

- [<span data-ttu-id="5f9c9-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="5f9c9-121">LaunchEvent</span></span>](launchevent.md)

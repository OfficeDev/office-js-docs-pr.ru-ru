---
title: Лаунчевентс в файле манифеста (Предварительная версия)
description: Элемент Лаунчевентс настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 2e1ad56d405fca0f85fad500a113fba7d0448caf
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278557"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="e5c62-103">Элемент Лаунчевентс (Preview)</span><span class="sxs-lookup"><span data-stu-id="e5c62-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="e5c62-104">Настраивает надстройку для активации на основе поддерживаемых событий.</span><span class="sxs-lookup"><span data-stu-id="e5c62-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="e5c62-105">Дочерний [`<ExtensionPoint>`](extensionpoint.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="e5c62-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="e5c62-106">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="e5c62-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="e5c62-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="e5c62-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e5c62-108">Активация на основе событий в настоящее время находится [в режиме предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="e5c62-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="e5c62-109">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="e5c62-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="e5c62-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e5c62-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="e5c62-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e5c62-111">Contained in</span></span>

<span data-ttu-id="e5c62-112">[ExtensionPoint](extensionpoint.md) (почтовые надстройки**лаунчевент** )</span><span class="sxs-lookup"><span data-stu-id="e5c62-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="e5c62-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e5c62-113">Child elements</span></span>

|  <span data-ttu-id="e5c62-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c62-114">Element</span></span> |  <span data-ttu-id="e5c62-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e5c62-115">Required</span></span>  |  <span data-ttu-id="e5c62-116">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c62-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="e5c62-117">лаунчевент</span><span class="sxs-lookup"><span data-stu-id="e5c62-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="e5c62-118">Да</span><span class="sxs-lookup"><span data-stu-id="e5c62-118">Yes</span></span> |  <span data-ttu-id="e5c62-119">Сопоставление поддерживаемого события с функцией в файле JavaScript для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="e5c62-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="e5c62-120">См. также</span><span class="sxs-lookup"><span data-stu-id="e5c62-120">See also</span></span>

- [<span data-ttu-id="e5c62-121">лаунчевент</span><span class="sxs-lookup"><span data-stu-id="e5c62-121">LaunchEvent</span></span>](launchevent.md)

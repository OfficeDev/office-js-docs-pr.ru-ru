---
title: Лаунчевент в файле манифеста (Предварительная версия)
description: Элемент Лаунчевент настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: a4f5208ec7f735d926c3a878cae34973c3992cf9
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278558"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="bb517-103">Элемент Лаунчевент (Preview)</span><span class="sxs-lookup"><span data-stu-id="bb517-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="bb517-104">Настраивает надстройку для активации на основе поддерживаемых событий.</span><span class="sxs-lookup"><span data-stu-id="bb517-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="bb517-105">Дочерний [`<LaunchEvents>`](launchevents.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="bb517-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="bb517-106">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="bb517-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="bb517-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="bb517-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bb517-108">Активация на основе событий в настоящее время находится [в режиме предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="bb517-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="bb517-109">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="bb517-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="bb517-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="bb517-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="bb517-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="bb517-111">Contained in</span></span>

- [<span data-ttu-id="bb517-112">лаунчевентс</span><span class="sxs-lookup"><span data-stu-id="bb517-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="bb517-113">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bb517-113">Attributes</span></span>

|  <span data-ttu-id="bb517-114">Атрибут</span><span class="sxs-lookup"><span data-stu-id="bb517-114">Attribute</span></span>  |  <span data-ttu-id="bb517-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bb517-115">Required</span></span>  |  <span data-ttu-id="bb517-116">Описание</span><span class="sxs-lookup"><span data-stu-id="bb517-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bb517-117">**Тип**</span><span class="sxs-lookup"><span data-stu-id="bb517-117">**Type**</span></span>  |  <span data-ttu-id="bb517-118">Да</span><span class="sxs-lookup"><span data-stu-id="bb517-118">Yes</span></span>  | <span data-ttu-id="bb517-119">Указывает поддерживаемый тип события.</span><span class="sxs-lookup"><span data-stu-id="bb517-119">Specifies a supported event type.</span></span> <span data-ttu-id="bb517-120">Доступны типы `OnNewMessageCompose` и `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="bb517-120">Available types are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> |
|  <span data-ttu-id="bb517-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="bb517-121">**FunctionName**</span></span>  |  <span data-ttu-id="bb517-122">Да</span><span class="sxs-lookup"><span data-stu-id="bb517-122">Yes</span></span>  | <span data-ttu-id="bb517-123">Задает имя функции JavaScript для обработки события, указанного в `Type` атрибуте.</span><span class="sxs-lookup"><span data-stu-id="bb517-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="bb517-124">См. также</span><span class="sxs-lookup"><span data-stu-id="bb517-124">See also</span></span>

- [<span data-ttu-id="bb517-125">лаунчевентс</span><span class="sxs-lookup"><span data-stu-id="bb517-125">LaunchEvents</span></span>](launchevents.md)

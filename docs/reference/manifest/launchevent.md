---
title: LaunchEvent в файле манифеста (предварительный просмотр)
description: Элемент LaunchEvent настраивает надстройки для активации на основе поддерживаемых событий.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 7283e9aba9ca57793019ffe027a7f4d6e3243aa8
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555313"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="28548-103">Элемент LaunchEvent (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="28548-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="28548-104">Настраивает надстройки для активации на основе поддерживаемых событий.</span><span class="sxs-lookup"><span data-stu-id="28548-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="28548-105">Дитя [`<LaunchEvents>`](launchevents.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="28548-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="28548-106">Для получения дополнительной информации [см Outlook.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="28548-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="28548-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="28548-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="28548-108">Активация на основе событий в [настоящее время находится](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) в предварительном просмотре и доступна только Outlook веб-сайтах и Windows.</span><span class="sxs-lookup"><span data-stu-id="28548-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and on Windows.</span></span> <span data-ttu-id="28548-109">Для получения дополнительной информации [узнайте, как просмотреть функцию активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="28548-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="28548-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="28548-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="28548-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="28548-111">Contained in</span></span>

- [<span data-ttu-id="28548-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="28548-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="28548-113">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="28548-113">Attributes</span></span>

|  <span data-ttu-id="28548-114">Атрибут</span><span class="sxs-lookup"><span data-stu-id="28548-114">Attribute</span></span>  |  <span data-ttu-id="28548-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="28548-115">Required</span></span>  |  <span data-ttu-id="28548-116">Описание</span><span class="sxs-lookup"><span data-stu-id="28548-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="28548-117">**Тип**</span><span class="sxs-lookup"><span data-stu-id="28548-117">**Type**</span></span>  |  <span data-ttu-id="28548-118">Да</span><span class="sxs-lookup"><span data-stu-id="28548-118">Yes</span></span>  | <span data-ttu-id="28548-119">Определяет поддерживаемый тип события.</span><span class="sxs-lookup"><span data-stu-id="28548-119">Specifies a supported event type.</span></span> <span data-ttu-id="28548-120">Для набора поддерживаемых типов см. Как [просмотреть функцию активации на основе событий.](../../outlook/autolaunch.md#supported-events)</span><span class="sxs-lookup"><span data-stu-id="28548-120">For the set of supported types, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="28548-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="28548-121">**FunctionName**</span></span>  |  <span data-ttu-id="28548-122">Да</span><span class="sxs-lookup"><span data-stu-id="28548-122">Yes</span></span>  | <span data-ttu-id="28548-123">Указывается название функции JavaScript для обработки события, указанного в `Type` атрибуте.</span><span class="sxs-lookup"><span data-stu-id="28548-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="28548-124">См. также</span><span class="sxs-lookup"><span data-stu-id="28548-124">See also</span></span>

- [<span data-ttu-id="28548-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="28548-125">LaunchEvents</span></span>](launchevents.md)

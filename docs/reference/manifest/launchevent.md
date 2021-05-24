---
title: LaunchEvent в файле манифеста
description: Элемент LaunchEvent настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: c866a085ed6b7a33c8d7bf02d25e6ec748629e07
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591081"
---
# <a name="launchevent-element"></a><span data-ttu-id="ac945-103">Элемент LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="ac945-103">LaunchEvent element</span></span>

<span data-ttu-id="ac945-104">Настраивает надстройка для активации на основе поддерживаемых событий.</span><span class="sxs-lookup"><span data-stu-id="ac945-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="ac945-105">Ребенок [`<LaunchEvents>`](launchevents.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="ac945-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="ac945-106">Дополнительные сведения см. в Outlook [надстройки](../../outlook/autolaunch.md)для активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="ac945-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="ac945-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="ac945-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ac945-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="ac945-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="ac945-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="ac945-109">Contained in</span></span>

- [<span data-ttu-id="ac945-110">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="ac945-110">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="ac945-111">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ac945-111">Attributes</span></span>

|  <span data-ttu-id="ac945-112">Атрибут</span><span class="sxs-lookup"><span data-stu-id="ac945-112">Attribute</span></span>  |  <span data-ttu-id="ac945-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ac945-113">Required</span></span>  |  <span data-ttu-id="ac945-114">Описание</span><span class="sxs-lookup"><span data-stu-id="ac945-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ac945-115">**Тип**</span><span class="sxs-lookup"><span data-stu-id="ac945-115">**Type**</span></span>  |  <span data-ttu-id="ac945-116">Да</span><span class="sxs-lookup"><span data-stu-id="ac945-116">Yes</span></span>  | <span data-ttu-id="ac945-117">Указывает поддерживаемый тип события.</span><span class="sxs-lookup"><span data-stu-id="ac945-117">Specifies a supported event type.</span></span> <span data-ttu-id="ac945-118">Для набора поддерживаемых типов см. в Outlook надстройку для [активации на](../../outlook/autolaunch.md#supported-events)основе событий.</span><span class="sxs-lookup"><span data-stu-id="ac945-118">For the set of supported types, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="ac945-119">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="ac945-119">**FunctionName**</span></span>  |  <span data-ttu-id="ac945-120">Да</span><span class="sxs-lookup"><span data-stu-id="ac945-120">Yes</span></span>  | <span data-ttu-id="ac945-121">Указывает имя функции JavaScript для обработки события, указанного в `Type` атрибуте.</span><span class="sxs-lookup"><span data-stu-id="ac945-121">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ac945-122">См. также</span><span class="sxs-lookup"><span data-stu-id="ac945-122">See also</span></span>

- [<span data-ttu-id="ac945-123">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="ac945-123">LaunchEvents</span></span>](launchevents.md)

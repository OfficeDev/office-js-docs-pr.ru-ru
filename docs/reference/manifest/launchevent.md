---
title: Лаунчевент в файле манифеста (Предварительная версия)
description: Элемент Лаунчевент настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 4874b9f4c14e3a999f41ec3fa20a15393b031ea6
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611780"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="05b66-103">Элемент Лаунчевент (Preview)</span><span class="sxs-lookup"><span data-stu-id="05b66-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="05b66-104">Настраивает надстройку для активации на основе поддерживаемых событий.</span><span class="sxs-lookup"><span data-stu-id="05b66-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="05b66-105">Дочерний [`<LaunchEvents>`](launchevents.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="05b66-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="05b66-106">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="05b66-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="05b66-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="05b66-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="05b66-108">Активация на основе событий в настоящее время находится [в режиме предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="05b66-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="05b66-109">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="05b66-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="05b66-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="05b66-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="05b66-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="05b66-111">Contained in</span></span>

- [<span data-ttu-id="05b66-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="05b66-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="05b66-113">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="05b66-113">Attributes</span></span>

|  <span data-ttu-id="05b66-114">Атрибут</span><span class="sxs-lookup"><span data-stu-id="05b66-114">Attribute</span></span>  |  <span data-ttu-id="05b66-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="05b66-115">Required</span></span>  |  <span data-ttu-id="05b66-116">Описание</span><span class="sxs-lookup"><span data-stu-id="05b66-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="05b66-117">**Тип**</span><span class="sxs-lookup"><span data-stu-id="05b66-117">**Type**</span></span>  |  <span data-ttu-id="05b66-118">Да</span><span class="sxs-lookup"><span data-stu-id="05b66-118">Yes</span></span>  | <span data-ttu-id="05b66-119">Указывает поддерживаемый тип события.</span><span class="sxs-lookup"><span data-stu-id="05b66-119">Specifies a supported event type.</span></span> <span data-ttu-id="05b66-120">Доступны типы `OnNewMessageCompose` и `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="05b66-120">Available types are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> |
|  <span data-ttu-id="05b66-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="05b66-121">**FunctionName**</span></span>  |  <span data-ttu-id="05b66-122">Да</span><span class="sxs-lookup"><span data-stu-id="05b66-122">Yes</span></span>  | <span data-ttu-id="05b66-123">Задает имя функции JavaScript для обработки события, указанного в `Type` атрибуте.</span><span class="sxs-lookup"><span data-stu-id="05b66-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="05b66-124">См. также</span><span class="sxs-lookup"><span data-stu-id="05b66-124">See also</span></span>

- [<span data-ttu-id="05b66-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="05b66-125">LaunchEvents</span></span>](launchevents.md)

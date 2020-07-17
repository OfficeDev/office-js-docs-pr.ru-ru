---
title: Среды выполнения в файле манифеста
description: Элемент Runtimes указывает среду выполнения надстройки.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 082491befc6b9dbdc474b0e40f9defd90a4ef75f
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159362"
---
# <a name="runtimes-element"></a><span data-ttu-id="bd797-103">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="bd797-103">Runtimes element</span></span>

<span data-ttu-id="bd797-104">Задает среду выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="bd797-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="bd797-105">Дочерний [`<Host>`](host.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="bd797-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="bd797-106">При работе в Office в Windows надстройка использует браузер Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="bd797-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="bd797-107">В Excel этот элемент позволяет использовать одну и ту же среду выполнения для ленты, области задач и пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="bd797-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="bd797-108">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="bd797-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="bd797-109">В Outlook этот элемент включает активацию надстройки на основе событий.</span><span class="sxs-lookup"><span data-stu-id="bd797-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="bd797-110">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="bd797-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="bd797-111">**Тип надстройки:** Область задач, почта</span><span class="sxs-lookup"><span data-stu-id="bd797-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd797-112">**Outlook**: функция активации на основе событий в настоящее время находится [в предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="bd797-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="bd797-113">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="bd797-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="bd797-114">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="bd797-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="bd797-115">Содержится в</span><span class="sxs-lookup"><span data-stu-id="bd797-115">Contained in</span></span>

[<span data-ttu-id="bd797-116">Host</span><span class="sxs-lookup"><span data-stu-id="bd797-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="bd797-117">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="bd797-117">Child elements</span></span>

|  <span data-ttu-id="bd797-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="bd797-118">Element</span></span> |  <span data-ttu-id="bd797-119">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bd797-119">Required</span></span>  |  <span data-ttu-id="bd797-120">Описание</span><span class="sxs-lookup"><span data-stu-id="bd797-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="bd797-121">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="bd797-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="bd797-122">Да</span><span class="sxs-lookup"><span data-stu-id="bd797-122">Yes</span></span> |  <span data-ttu-id="bd797-123">Среда выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="bd797-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="bd797-124">См. также</span><span class="sxs-lookup"><span data-stu-id="bd797-124">See also</span></span>

- [<span data-ttu-id="bd797-125">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="bd797-125">Runtime</span></span>](runtime.md)

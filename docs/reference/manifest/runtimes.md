---
title: Среды выполнения в файле манифеста
description: Элемент Runtimes указывает среду выполнения надстройки.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: ef00bea317ae479d912b3a02f269ef97045b015d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608099"
---
# <a name="runtimes-element"></a><span data-ttu-id="1c294-103">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="1c294-103">Runtimes element</span></span>

<span data-ttu-id="1c294-104">Задает среду выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="1c294-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="1c294-105">Дочерний [`<Host>`](host.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="1c294-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="1c294-106">При работе в Office в Windows надстройка использует браузер Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="1c294-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="1c294-107">В Excel этот элемент позволяет использовать одну и ту же среду выполнения для ленты, области задач и пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="1c294-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="1c294-108">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="1c294-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="1c294-109">В Outlook этот элемент включает активацию надстройки на основе событий.</span><span class="sxs-lookup"><span data-stu-id="1c294-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="1c294-110">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="1c294-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="1c294-111">**Тип надстройки:** Область задач, почта</span><span class="sxs-lookup"><span data-stu-id="1c294-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1c294-112">**Excel**: общая среда выполнения в настоящее время доступна только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="1c294-112">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="1c294-113">**Outlook**: функция активации на основе событий в настоящее время находится [в предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="1c294-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="1c294-114">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="1c294-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="1c294-115">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="1c294-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="1c294-116">Содержится в</span><span class="sxs-lookup"><span data-stu-id="1c294-116">Contained in</span></span>

<span data-ttu-id="1c294-117">[Host](host.md) (Узел)</span><span class="sxs-lookup"><span data-stu-id="1c294-117">[Host](host.md)</span></span>

## <a name="child-elements"></a><span data-ttu-id="1c294-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="1c294-118">Child elements</span></span>

|  <span data-ttu-id="1c294-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="1c294-119">Element</span></span> |  <span data-ttu-id="1c294-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1c294-120">Required</span></span>  |  <span data-ttu-id="1c294-121">Описание</span><span class="sxs-lookup"><span data-stu-id="1c294-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="1c294-122">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="1c294-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="1c294-123">Да</span><span class="sxs-lookup"><span data-stu-id="1c294-123">Yes</span></span> |  <span data-ttu-id="1c294-124">Среда выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="1c294-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="1c294-125">См. также</span><span class="sxs-lookup"><span data-stu-id="1c294-125">See also</span></span>

- [<span data-ttu-id="1c294-126">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="1c294-126">Runtime</span></span>](runtime.md)

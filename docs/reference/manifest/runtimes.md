---
title: Среды выполнения в файле манифеста
description: Элемент Runtimes указывает среду выполнения надстройки.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 22156a171ca2f423024efb1b3d2a6fdae07dfef6
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278366"
---
# <a name="runtimes-element"></a><span data-ttu-id="c4ac5-103">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="c4ac5-103">Runtimes element</span></span>

<span data-ttu-id="c4ac5-104">Задает среду выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="c4ac5-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="c4ac5-105">Дочерний [`<Host>`](host.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="c4ac5-105">Child of the [`<Host>`](host.md) element.</span></span>

<span data-ttu-id="c4ac5-106">В Excel этот элемент позволяет использовать одну и ту же среду выполнения для ленты, области задач и пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="c4ac5-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="c4ac5-107">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="c4ac5-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="c4ac5-108">В Outlook этот элемент включает активацию надстройки на основе событий.</span><span class="sxs-lookup"><span data-stu-id="c4ac5-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="c4ac5-109">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="c4ac5-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="c4ac5-110">**Тип надстройки:** Область задач, почта</span><span class="sxs-lookup"><span data-stu-id="c4ac5-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c4ac5-111">**Excel**: общая среда выполнения в настоящее время находится в режиме предварительной версии и доступна только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="c4ac5-111">**Excel**: Shared runtime is currently in preview and only available in Excel on Windows.</span></span> <span data-ttu-id="c4ac5-112">Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="c4ac5-112">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>
>
> <span data-ttu-id="c4ac5-113">**Outlook**: функция активации на основе событий в настоящее время находится [в предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c4ac5-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="c4ac5-114">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="c4ac5-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="c4ac5-115">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="c4ac5-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="c4ac5-116">Содержится в</span><span class="sxs-lookup"><span data-stu-id="c4ac5-116">Contained in</span></span>

<span data-ttu-id="c4ac5-117">[Host](host.md) (Узел)</span><span class="sxs-lookup"><span data-stu-id="c4ac5-117">[Host](host.md)</span></span>

## <a name="child-elements"></a><span data-ttu-id="c4ac5-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c4ac5-118">Child elements</span></span>

|  <span data-ttu-id="c4ac5-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="c4ac5-119">Element</span></span> |  <span data-ttu-id="c4ac5-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c4ac5-120">Required</span></span>  |  <span data-ttu-id="c4ac5-121">Описание</span><span class="sxs-lookup"><span data-stu-id="c4ac5-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="c4ac5-122">Среда выполнения</span><span class="sxs-lookup"><span data-stu-id="c4ac5-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="c4ac5-123">Да</span><span class="sxs-lookup"><span data-stu-id="c4ac5-123">Yes</span></span> |  <span data-ttu-id="c4ac5-124">Среда выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="c4ac5-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="c4ac5-125">См. также</span><span class="sxs-lookup"><span data-stu-id="c4ac5-125">See also</span></span>

- [<span data-ttu-id="c4ac5-126">Среда выполнения</span><span class="sxs-lookup"><span data-stu-id="c4ac5-126">Runtime</span></span>](runtime.md)

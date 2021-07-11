---
title: Элемент GetStarted в файле манифеста
description: Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a637f3f9031d9f8e09d14f17f2095ca0647c4d50
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348687"
---
# <a name="getstarted-element"></a><span data-ttu-id="d396e-103">Элемент GetStarted</span><span class="sxs-lookup"><span data-stu-id="d396e-103">GetStarted element</span></span>

<span data-ttu-id="d396e-104">Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="d396e-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="d396e-105">Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="d396e-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="d396e-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="d396e-106">Child elements</span></span>

| <span data-ttu-id="d396e-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="d396e-107">Element</span></span>                       | <span data-ttu-id="d396e-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d396e-108">Required</span></span> | <span data-ttu-id="d396e-109">Описание</span><span class="sxs-lookup"><span data-stu-id="d396e-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="d396e-110">Title</span><span class="sxs-lookup"><span data-stu-id="d396e-110">Title</span></span>](#title)               | <span data-ttu-id="d396e-111">Да</span><span class="sxs-lookup"><span data-stu-id="d396e-111">Yes</span></span>      | <span data-ttu-id="d396e-112">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="d396e-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="d396e-113">Описание</span><span class="sxs-lookup"><span data-stu-id="d396e-113">Description</span></span>](#description)   | <span data-ttu-id="d396e-114">Да</span><span class="sxs-lookup"><span data-stu-id="d396e-114">Yes</span></span>      | <span data-ttu-id="d396e-115">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d396e-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="d396e-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="d396e-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="d396e-117">Да</span><span class="sxs-lookup"><span data-stu-id="d396e-117">Yes</span></span>       | <span data-ttu-id="d396e-118">URL-адрес страницы с подробным описанием надстройки.</span><span class="sxs-lookup"><span data-stu-id="d396e-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="d396e-119">Title</span><span class="sxs-lookup"><span data-stu-id="d396e-119">Title</span></span> 

<span data-ttu-id="d396e-120">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="d396e-120">Required.</span></span> <span data-ttu-id="d396e-121">Заголовок в верхней части выноски.</span><span class="sxs-lookup"><span data-stu-id="d396e-121">The title used for the top of the callout.</span></span> <span data-ttu-id="d396e-122">Атрибут **resid** ссылается на действительный ID в **элементе ShortStrings** в разделе [Ресурсы](resources.md) и может быть не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="d396e-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="d396e-123">Описание</span><span class="sxs-lookup"><span data-stu-id="d396e-123">Description</span></span>

<span data-ttu-id="d396e-124">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="d396e-124">Required.</span></span> <span data-ttu-id="d396e-125">Описание и основной текст выноски.</span><span class="sxs-lookup"><span data-stu-id="d396e-125">The description / body content for the callout.</span></span> <span data-ttu-id="d396e-126">Атрибут **resid** ссылается на допустимый ID в **элементе LongStrings** в разделе [Ресурсы](resources.md) и может быть не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="d396e-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="d396e-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="d396e-127">LearnMoreUrl</span></span>

<span data-ttu-id="d396e-128">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="d396e-128">Required.</span></span> <span data-ttu-id="d396e-129">URL-адрес страницы, где пользователь может узнать больше о надстройке.</span><span class="sxs-lookup"><span data-stu-id="d396e-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="d396e-130">Атрибут **resid** ссылается на допустимый ID в **элементе Urls** в разделе [Ресурсы](resources.md) и может быть не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="d396e-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="d396e-131">В настоящее время элемент **LearnMoreUrl** не отображается в клиентах Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="d396e-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="d396e-132">Рекомендуем добавить URL-адрес всех клиентов, чтобы этот адрес отображался, когда он станет доступен.</span><span class="sxs-lookup"><span data-stu-id="d396e-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="d396e-133">См. также</span><span class="sxs-lookup"><span data-stu-id="d396e-133">See also</span></span>

<span data-ttu-id="d396e-134">В следующих примерах кода используется **элемент GetStarted.**</span><span class="sxs-lookup"><span data-stu-id="d396e-134">The following code samples use the **GetStarted** element.</span></span>

* [<span data-ttu-id="d396e-135">Веб-надстройка Excel для работы с форматированием таблиц и диаграмм</span><span class="sxs-lookup"><span data-stu-id="d396e-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="d396e-136">JavaScript SpecKit для надстроек Word</span><span class="sxs-lookup"><span data-stu-id="d396e-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="d396e-137">Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d396e-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)

---
title: Элемент GetStarted в файле манифеста
description: Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0ad6196dc45e4ea06c2b43ac5da66a560ab0b899
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771419"
---
# <a name="getstarted-element"></a><span data-ttu-id="e075a-103">Элемент GetStarted</span><span class="sxs-lookup"><span data-stu-id="e075a-103">GetStarted element</span></span>

<span data-ttu-id="e075a-104">Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote.</span><span class="sxs-lookup"><span data-stu-id="e075a-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="e075a-105">Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e075a-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="e075a-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e075a-106">Child elements</span></span>

| <span data-ttu-id="e075a-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="e075a-107">Element</span></span>                       | <span data-ttu-id="e075a-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e075a-108">Required</span></span> | <span data-ttu-id="e075a-109">Описание</span><span class="sxs-lookup"><span data-stu-id="e075a-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="e075a-110">Title</span><span class="sxs-lookup"><span data-stu-id="e075a-110">Title</span></span>](#title)               | <span data-ttu-id="e075a-111">Да</span><span class="sxs-lookup"><span data-stu-id="e075a-111">Yes</span></span>      | <span data-ttu-id="e075a-112">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="e075a-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="e075a-113">Описание</span><span class="sxs-lookup"><span data-stu-id="e075a-113">Description</span></span>](#description)   | <span data-ttu-id="e075a-114">Да</span><span class="sxs-lookup"><span data-stu-id="e075a-114">Yes</span></span>      | <span data-ttu-id="e075a-115">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e075a-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="e075a-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="e075a-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="e075a-117">Да</span><span class="sxs-lookup"><span data-stu-id="e075a-117">Yes</span></span>       | <span data-ttu-id="e075a-118">URL-адрес страницы с подробным описанием надстройки.</span><span class="sxs-lookup"><span data-stu-id="e075a-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="e075a-119">Title</span><span class="sxs-lookup"><span data-stu-id="e075a-119">Title</span></span> 

<span data-ttu-id="e075a-120">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="e075a-120">Required.</span></span> <span data-ttu-id="e075a-121">Заголовок в верхней части выноски.</span><span class="sxs-lookup"><span data-stu-id="e075a-121">The title used for the top of the callout.</span></span> <span data-ttu-id="e075a-122">Атрибут **resid** ссылается на допустимый ИД в **элементе ShortStrings** в разделе ["Ресурсы"](resources.md) и может иметь не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="e075a-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="e075a-123">Описание</span><span class="sxs-lookup"><span data-stu-id="e075a-123">Description</span></span>

<span data-ttu-id="e075a-124">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="e075a-124">Required.</span></span> <span data-ttu-id="e075a-125">Описание и основной текст выноски.</span><span class="sxs-lookup"><span data-stu-id="e075a-125">The description / body content for the callout.</span></span> <span data-ttu-id="e075a-126">Атрибут **resid** ссылается на допустимый ИД в **элементе LongStrings** в разделе ["Ресурсы"](resources.md) и может иметь длину не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="e075a-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="e075a-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="e075a-127">LearnMoreUrl</span></span>

<span data-ttu-id="e075a-128">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="e075a-128">Required.</span></span> <span data-ttu-id="e075a-129">URL-адрес страницы, где пользователь может узнать больше о надстройке.</span><span class="sxs-lookup"><span data-stu-id="e075a-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="e075a-130">Атрибут **resid** ссылается на допустимый ИД в **элементе Urls** в разделе ["Ресурсы"](resources.md) и может иметь не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="e075a-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="e075a-131">В настоящее время элемент **LearnMoreUrl** не отображается в клиентах Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="e075a-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="e075a-132">Рекомендуем добавить URL-адрес всех клиентов, чтобы этот адрес отображался, когда он станет доступен.</span><span class="sxs-lookup"><span data-stu-id="e075a-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="e075a-133">См. также</span><span class="sxs-lookup"><span data-stu-id="e075a-133">See also</span></span>

<span data-ttu-id="e075a-134">В следующих примерах кода используется элемент **GetStarted**:</span><span class="sxs-lookup"><span data-stu-id="e075a-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="e075a-135">Веб-надстройка Excel для работы с форматированием таблиц и диаграмм</span><span class="sxs-lookup"><span data-stu-id="e075a-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="e075a-136">JavaScript SpecKit для надстроек Word</span><span class="sxs-lookup"><span data-stu-id="e075a-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="e075a-137">Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e075a-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)

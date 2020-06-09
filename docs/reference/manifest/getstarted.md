---
title: Элемент GetStarted в файле манифеста
description: Предоставляет сведения для выноски, которая отображается при установке надстройки в ведущих приложениях Word, Excel, PowerPoint и OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c1fbdd5d4f4365f9f8190805519fc7a70c8c87ca
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611836"
---
# <a name="getstarted-element"></a><span data-ttu-id="76d36-103">Элемент GetStarted</span><span class="sxs-lookup"><span data-stu-id="76d36-103">GetStarted element</span></span>

<span data-ttu-id="76d36-p101">Предоставляет сведения для выноски, которая отображается при установке надстройки в ведущих приложениях Word, Excel, PowerPoint и OneNote. Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="76d36-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="76d36-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="76d36-106">Child elements</span></span>

| <span data-ttu-id="76d36-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="76d36-107">Element</span></span>                       | <span data-ttu-id="76d36-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="76d36-108">Required</span></span> | <span data-ttu-id="76d36-109">Описание</span><span class="sxs-lookup"><span data-stu-id="76d36-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="76d36-110">Title</span><span class="sxs-lookup"><span data-stu-id="76d36-110">Title</span></span>](#title)               | <span data-ttu-id="76d36-111">Да</span><span class="sxs-lookup"><span data-stu-id="76d36-111">Yes</span></span>      | <span data-ttu-id="76d36-112">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="76d36-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="76d36-113">Описание</span><span class="sxs-lookup"><span data-stu-id="76d36-113">Description</span></span>](#description)   | <span data-ttu-id="76d36-114">Да</span><span class="sxs-lookup"><span data-stu-id="76d36-114">Yes</span></span>      | <span data-ttu-id="76d36-115">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="76d36-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="76d36-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="76d36-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="76d36-117">Да</span><span class="sxs-lookup"><span data-stu-id="76d36-117">Yes</span></span>       | <span data-ttu-id="76d36-118">URL-адрес страницы с подробным описанием надстройки.</span><span class="sxs-lookup"><span data-stu-id="76d36-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="76d36-119">Название</span><span class="sxs-lookup"><span data-stu-id="76d36-119">Title</span></span> 

<span data-ttu-id="76d36-p102">Обязательный. Заголовок в верхней части выноски. Атрибут **resid** ссылается на допустимый идентификатор элемента **ShortStrings** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="76d36-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="76d36-123">Описание</span><span class="sxs-lookup"><span data-stu-id="76d36-123">Description</span></span>

<span data-ttu-id="76d36-p103">Обязательный. Описание и основной текст выноски. Атрибут **resid** ссылается на допустимый идентификатор элемента **LongStrings** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="76d36-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="76d36-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="76d36-127">LearnMoreUrl</span></span>

<span data-ttu-id="76d36-p104">Обязательный. URL-адрес страницы, где пользователь может узнать больше о надстройке. Атрибут **resid** ссылается на допустимый идентификатор элемента **Urls** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="76d36-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="76d36-131">В настоящее время элемент **LearnMoreUrl** не отображается в клиентах Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="76d36-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="76d36-132">Рекомендуем добавить URL-адрес всех клиентов, чтобы этот адрес отображался, когда он станет доступен.</span><span class="sxs-lookup"><span data-stu-id="76d36-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="76d36-133">См. также</span><span class="sxs-lookup"><span data-stu-id="76d36-133">See also</span></span>

<span data-ttu-id="76d36-134">В следующих примерах кода используется элемент **GetStarted**:</span><span class="sxs-lookup"><span data-stu-id="76d36-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="76d36-135">Веб-надстройка Excel для работы с форматированием таблиц и диаграмм</span><span class="sxs-lookup"><span data-stu-id="76d36-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="76d36-136">JavaScript SpecKit для надстроек Word</span><span class="sxs-lookup"><span data-stu-id="76d36-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="76d36-137">Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint</span><span class="sxs-lookup"><span data-stu-id="76d36-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)

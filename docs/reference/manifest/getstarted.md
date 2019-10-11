---
title: Элемент GetStarted в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 82fa1b9b62674adfb05c07536a7fdf2bbabf8f45
ms.sourcegitcommit: e5a5ec4ba32bacd0ccd13291b4e7f4bfc42901a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/09/2019
ms.locfileid: "37429742"
---
# <a name="getstarted-element"></a><span data-ttu-id="90eba-102">Элемент GetStarted</span><span class="sxs-lookup"><span data-stu-id="90eba-102">GetStarted element</span></span>

<span data-ttu-id="90eba-p101">Предоставляет сведения для выноски, которая отображается при установке надстройки в ведущих приложениях Word, Excel, PowerPoint и OneNote. Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="90eba-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="90eba-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="90eba-105">Child elements</span></span>

| <span data-ttu-id="90eba-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="90eba-106">Element</span></span>                       | <span data-ttu-id="90eba-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="90eba-107">Required</span></span> | <span data-ttu-id="90eba-108">Описание</span><span class="sxs-lookup"><span data-stu-id="90eba-108">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="90eba-109">Title</span><span class="sxs-lookup"><span data-stu-id="90eba-109">Title</span></span>](#title)               | <span data-ttu-id="90eba-110">Да</span><span class="sxs-lookup"><span data-stu-id="90eba-110">Yes</span></span>      | <span data-ttu-id="90eba-111">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="90eba-111">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="90eba-112">Описание</span><span class="sxs-lookup"><span data-stu-id="90eba-112">Description</span></span>](#description)   | <span data-ttu-id="90eba-113">Да</span><span class="sxs-lookup"><span data-stu-id="90eba-113">Yes</span></span>      | <span data-ttu-id="90eba-114">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="90eba-114">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="90eba-115">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="90eba-115">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="90eba-116">Да</span><span class="sxs-lookup"><span data-stu-id="90eba-116">Yes</span></span>       | <span data-ttu-id="90eba-117">URL-адрес страницы с подробным описанием надстройки.</span><span class="sxs-lookup"><span data-stu-id="90eba-117">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="90eba-118">Название</span><span class="sxs-lookup"><span data-stu-id="90eba-118">Title</span></span> 

<span data-ttu-id="90eba-p102">Обязательный. Заголовок в верхней части выноски. Атрибут **resid** ссылается на допустимый идентификатор элемента **ShortStrings** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="90eba-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="90eba-122">Описание</span><span class="sxs-lookup"><span data-stu-id="90eba-122">Description</span></span>

<span data-ttu-id="90eba-p103">Обязательный. Описание и основной текст выноски. Атрибут **resid** ссылается на допустимый идентификатор элемента **LongStrings** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="90eba-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="90eba-126">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="90eba-126">LearnMoreUrl</span></span>

<span data-ttu-id="90eba-p104">Обязательный. URL-адрес страницы, где пользователь может узнать больше о надстройке. Атрибут **resid** ссылается на допустимый идентификатор элемента **Urls** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="90eba-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="90eba-130">В настоящее время элемент **LearnMoreUrl** не отображается в клиентах Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="90eba-130">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="90eba-131">Рекомендуем добавить URL-адрес всех клиентов, чтобы этот адрес отображался, когда он станет доступен.</span><span class="sxs-lookup"><span data-stu-id="90eba-131">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="90eba-132">См. также</span><span class="sxs-lookup"><span data-stu-id="90eba-132">See also</span></span>

<span data-ttu-id="90eba-133">В следующих примерах кода используется элемент **GetStarted**:</span><span class="sxs-lookup"><span data-stu-id="90eba-133">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="90eba-134">Веб-надстройка Excel для работы с форматированием таблиц и диаграмм</span><span class="sxs-lookup"><span data-stu-id="90eba-134">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="90eba-135">JavaScript SpecKit для надстроек Word</span><span class="sxs-lookup"><span data-stu-id="90eba-135">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="90eba-136">Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint</span><span class="sxs-lookup"><span data-stu-id="90eba-136">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)

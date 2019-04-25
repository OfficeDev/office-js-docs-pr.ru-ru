---
title: Изучение API JavaScript для Office с помощью сценария Lab
description: Используйте сценарий "Лаборатория" для изучения API Office JS и прототипов функций.
ms.topic: article
ms.date: 04/23/2019
localization_priority: Normal
ms.openlocfilehash: 76888716cec8bd1754b7baa22dfcfbe5af984ea5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32640292"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="703e6-103">Изучение API JavaScript для Office с помощью сценария Lab</span><span class="sxs-lookup"><span data-stu-id="703e6-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="703e6-104">Надстройка " [Лаборатория скриптов](https://store.office.com/app.aspx?assetid=WA104380862)", доступная бесплатно из магазина Office, позволяет изучать API JavaScript для Office при работе с программами Office, такими как Excel или Word.</span><span class="sxs-lookup"><span data-stu-id="703e6-104">The [Script Lab add-in](https://store.office.com/app.aspx?assetid=WA104380862), which is available free from the Office store, enables you to explore the Office JavaScript API while you are working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="703e6-105">Script Lab — удобное средство для добавления в набор средств разработки в качестве прототипа и проверки функциональных возможностей, которые должны быть в надстройке.</span><span class="sxs-lookup"><span data-stu-id="703e6-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="703e6-106">Что такое "Лаборатория скриптов"?</span><span class="sxs-lookup"><span data-stu-id="703e6-106">What is Script Lab?</span></span>

<span data-ttu-id="703e6-107">Script Lab — это средство для тех, кто хочет научиться разрабатывать надстройки Office с помощью API JavaScript для Office в Excel, Word или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="703e6-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="703e6-108">Он предоставляет IntelliSense, чтобы вы могли видеть доступные и созданные на платформе Монако платформы, ту же платформу, которая используется в Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="703e6-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="703e6-109">С помощью сценария Lab можно получить доступ к библиотеке образцов, чтобы быстро испытать функции, или выбрать пример в качестве основы для собственного кода.</span><span class="sxs-lookup"><span data-stu-id="703e6-109">Through Script Lab, you can access a library of samples to quickly try out features or you can choose a sample as the base for your own code.</span></span> <span data-ttu-id="703e6-110">Вы также можете развернуть учебную библиотеку, добавив фрагменты в репозиторий [Office – JS — фрагменты](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span><span class="sxs-lookup"><span data-stu-id="703e6-110">You are also welcome to expand the sample library by adding snippets to the [office-js-snippets repo](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span></span> <span data-ttu-id="703e6-111">Еще одна интересная функция лаборатории сценариев — бета-версии или предварительные функции, такие как [пользовательские функции](/office/dev/add-ins/excel/custom-functions-overview) , которые можно попробовать.</span><span class="sxs-lookup"><span data-stu-id="703e6-111">Another exciting feature of Script Lab is beta or preview functionality like [custom functions](/office/dev/add-ins/excel/custom-functions-overview) is available for you to try.</span></span>

> [!TIP]
> <span data-ttu-id="703e6-112">Чтобы принять участие в бета-версии или предварительной версии, вам может потребоваться зарегистрироваться в [программе предварительНой оценки Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="703e6-112">To participate in beta or preview, you may have to sign up for the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="703e6-113">Звучит хорошо?</span><span class="sxs-lookup"><span data-stu-id="703e6-113">Sounds good so far?</span></span> <span data-ttu-id="703e6-114">Просмотрите этот видеоролик в виде одной минуты, чтобы увидеть Лаборатория сценариев в действии.</span><span class="sxs-lookup"><span data-stu-id="703e6-114">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="703e6-115">[![ПредварительНый Просмотр видео, в котором показана Лаборатория скриптов, работающая в Excel, Word и PowerPoint Online.] (../images/screenshot-wide-youtube.png 'Видеоролик о предварительном просмотре в лаборатории сценариев')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="703e6-115">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint Online.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="script-lab-supported-clients"></a><span data-ttu-id="703e6-116">Клиенты, поддерживаемые лабораториями скриптов</span><span class="sxs-lookup"><span data-stu-id="703e6-116">Script Lab supported clients</span></span>

<span data-ttu-id="703e6-117">Лаборатория скриптов поддерживается для Excel, Word и PowerPoint на следующих клиентах.</span><span class="sxs-lookup"><span data-stu-id="703e6-117">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="703e6-118">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="703e6-118">Office 365 for Windows</span></span>
- <span data-ttu-id="703e6-119">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="703e6-119">Office 365 for Mac</span></span>
- <span data-ttu-id="703e6-120">Office Online</span><span class="sxs-lookup"><span data-stu-id="703e6-120">Office Online</span></span>
- <span data-ttu-id="703e6-121">Office 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="703e6-121">Office 2013 or later for Windows</span></span>
- <span data-ttu-id="703e6-122">Office 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="703e6-122">Office 2016 or later for Mac</span></span>

## <a name="next-steps"></a><span data-ttu-id="703e6-123">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="703e6-123">Next steps</span></span>

<span data-ttu-id="703e6-124">Когда вы будете готовы создать надстройку для Office, ознакомьтесь с [5 минутным кратким руководством](/office/dev/add-ins/#5-minute-quick-starts) для предпочитаемого приложения Office.</span><span class="sxs-lookup"><span data-stu-id="703e6-124">When you're ready to create your Office Add-in, see the [5-minute quick start](/office/dev/add-ins/#5-minute-quick-starts) for your preferred Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="703e6-125">См. также</span><span class="sxs-lookup"><span data-stu-id="703e6-125">See also</span></span>

- [<span data-ttu-id="703e6-126">Получение лаборатории сценариев</span><span class="sxs-lookup"><span data-stu-id="703e6-126">Get Script Lab</span></span>](https://store.office.com/app.aspx?assetid=WA104380862)
- [<span data-ttu-id="703e6-127">Дополнительные сведения о лаборатории сценариев</span><span class="sxs-lookup"><span data-stu-id="703e6-127">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="703e6-128">Регистрация в программе для разработки</span><span class="sxs-lookup"><span data-stu-id="703e6-128">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)

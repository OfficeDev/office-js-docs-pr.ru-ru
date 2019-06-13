---
title: Изучение API JavaScript для Office с помощью сценария Lab
description: Используйте сценарий "Лаборатория" для изучения API Office JS и прототипов функций.
ms.topic: article
ms.date: 06/07/2019
localization_priority: Normal
ms.openlocfilehash: 0bab566b08ba25dd3c01cff72f331b2dc9ce304d
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910192"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="743d3-103">Изучение API JavaScript для Office с помощью сценария Lab</span><span class="sxs-lookup"><span data-stu-id="743d3-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="743d3-104">Надстройка " [Лаборатория скриптов](https://store.office.com/app.aspx?assetid=WA104380862)", доступная бесплатно из магазина Office, позволяет изучать API JavaScript для Office при работе с программами Office, такими как Excel или Word.</span><span class="sxs-lookup"><span data-stu-id="743d3-104">The [Script Lab add-in](https://store.office.com/app.aspx?assetid=WA104380862), which is available free from the Office store, enables you to explore the Office JavaScript API while you are working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="743d3-105">Script Lab — удобное средство для добавления в набор средств разработки в качестве прототипа и проверки функциональных возможностей, которые должны быть в надстройке.</span><span class="sxs-lookup"><span data-stu-id="743d3-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="743d3-106">Что такое "Лаборатория скриптов"?</span><span class="sxs-lookup"><span data-stu-id="743d3-106">What is Script Lab?</span></span>

<span data-ttu-id="743d3-107">Script Lab — это средство для тех, кто хочет научиться разрабатывать надстройки Office с помощью API JavaScript для Office в Excel, Word или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="743d3-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="743d3-108">Он предоставляет IntelliSense, чтобы вы могли видеть доступные и созданные на платформе Монако платформы, ту же платформу, которая используется в Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="743d3-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="743d3-109">С помощью сценария Lab можно получить доступ к библиотеке образцов, чтобы быстро испытать функции, или выбрать пример в качестве основы для собственного кода.</span><span class="sxs-lookup"><span data-stu-id="743d3-109">Through Script Lab, you can access a library of samples to quickly try out features or you can choose a sample as the base for your own code.</span></span> <span data-ttu-id="743d3-110">Вы также можете развернуть учебную библиотеку, добавив фрагменты в репозиторий [Office – JS — фрагменты](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span><span class="sxs-lookup"><span data-stu-id="743d3-110">You are also welcome to expand the sample library by adding snippets to the [office-js-snippets repo](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span></span> <span data-ttu-id="743d3-111">Еще одна интересная функция лаборатории сценариев — бета-версия или предварительная версия функции, которую вы можете попробовать.</span><span class="sxs-lookup"><span data-stu-id="743d3-111">Another exciting feature of Script Lab is beta or preview functionality is available for you to try.</span></span>

> [!TIP]
> <span data-ttu-id="743d3-112">Чтобы принять участие в бета-версии или предварительной версии, вам может потребоваться зарегистрироваться в [программе предварительной оценки Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="743d3-112">To participate in beta or preview, you may have to sign up for the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="743d3-113">Звучит хорошо?</span><span class="sxs-lookup"><span data-stu-id="743d3-113">Sounds good so far?</span></span> <span data-ttu-id="743d3-114">Просмотрите этот видеоролик в виде одной минуты, чтобы увидеть Лаборатория сценариев в действии.</span><span class="sxs-lookup"><span data-stu-id="743d3-114">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="743d3-115">[![Предварительный просмотр видео, в котором показана Лаборатория скриптов, работающая в Excel, Word и PowerPoint Online.] (../images/screenshot-wide-youtube.png 'Видеоролик о предварительном просмотре в лаборатории сценариев')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="743d3-115">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint Online.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="script-lab-supported-clients"></a><span data-ttu-id="743d3-116">Клиенты, поддерживаемые лабораториями скриптов</span><span class="sxs-lookup"><span data-stu-id="743d3-116">Script Lab supported clients</span></span>

<span data-ttu-id="743d3-117">Лаборатория скриптов поддерживается для Excel, Word и PowerPoint на следующих клиентах.</span><span class="sxs-lookup"><span data-stu-id="743d3-117">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="743d3-118">Office в Windows (подключено к Office 365)</span><span class="sxs-lookup"><span data-stu-id="743d3-118">Office on Windows (connected to Office 365)</span></span>
- <span data-ttu-id="743d3-119">Office для Mac (подключено к Office 365)</span><span class="sxs-lookup"><span data-stu-id="743d3-119">Office for Mac (connected to Office 365)</span></span>
- <span data-ttu-id="743d3-120">Office Online</span><span class="sxs-lookup"><span data-stu-id="743d3-120">Office Online</span></span>
- <span data-ttu-id="743d3-121">Office 2013 или более поздней версии в Windows</span><span class="sxs-lookup"><span data-stu-id="743d3-121">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="743d3-122">Office 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="743d3-122">Office 2016 or later for Mac</span></span>

## <a name="next-steps"></a><span data-ttu-id="743d3-123">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="743d3-123">Next steps</span></span>

<span data-ttu-id="743d3-124">Когда вы будете готовы создать надстройку для Office, ознакомьтесь с [5 минутным кратким руководством](/office/dev/add-ins/#5-minute-quick-starts) для предпочитаемого приложения Office.</span><span class="sxs-lookup"><span data-stu-id="743d3-124">When you're ready to create your Office Add-in, see the [5-minute quick start](/office/dev/add-ins/#5-minute-quick-starts) for your preferred Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="743d3-125">См. также</span><span class="sxs-lookup"><span data-stu-id="743d3-125">See also</span></span>

- [<span data-ttu-id="743d3-126">Получение лаборатории сценариев</span><span class="sxs-lookup"><span data-stu-id="743d3-126">Get Script Lab</span></span>](https://store.office.com/app.aspx?assetid=WA104380862)
- [<span data-ttu-id="743d3-127">Дополнительные сведения о лаборатории сценариев</span><span class="sxs-lookup"><span data-stu-id="743d3-127">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="743d3-128">Регистрация в программе для разработки</span><span class="sxs-lookup"><span data-stu-id="743d3-128">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)

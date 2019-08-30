---
title: Изучение API JavaScript для Office с помощью сценария Lab
description: Используйте сценарий "Лаборатория" для изучения API Office JS и прототипов функций.
ms.date: 07/05/2019
ms.topic: overview
scenarios: getting-started
localization_priority: Normal
ms.openlocfilehash: 908d27cdb5c8a7d4bc080c266cdb4d604114c42f
ms.sourcegitcommit: 49af31060aa56c1e1ec1e08682914d3cbefc3f1c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/29/2019
ms.locfileid: "36672840"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="5dbb4-103">Изучение API JavaScript для Office с помощью сценария Lab</span><span class="sxs-lookup"><span data-stu-id="5dbb4-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="5dbb4-104">Надстройка " [Лаборатория скриптов](https://appsource.microsoft.com/product/office/WA104380862)", доступная бесплатно из AppSource, позволяет изучать API JavaScript для Office при работе с программами Office, такими как Excel или Word.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-104">The [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862), which is available free from AppSource, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="5dbb4-105">Script Lab — удобное средство для добавления в набор средств разработки в качестве прототипа и проверки функциональных возможностей, которые должны быть в надстройке.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="5dbb4-106">Что такое "Лаборатория скриптов"?</span><span class="sxs-lookup"><span data-stu-id="5dbb4-106">What is Script Lab?</span></span>

<span data-ttu-id="5dbb4-107">Script Lab — это средство для тех, кто хочет научиться разрабатывать надстройки Office с помощью API JavaScript для Office в Excel, Word или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="5dbb4-108">Он предоставляет IntelliSense, чтобы вы могли видеть доступные и созданные на платформе Монако платформы, ту же платформу, которая используется в Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="5dbb4-109">С помощью сценария Lab вы можете получить доступ к библиотеке образцов, чтобы быстро испытать функции, или вы можете использовать пример в качестве отправной точки для собственного кода.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="5dbb4-110">Вы также можете воспользоваться лабораториями скриптов для предварительной версии API.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="5dbb4-111">Звучит хорошо?</span><span class="sxs-lookup"><span data-stu-id="5dbb4-111">Sounds good so far?</span></span> <span data-ttu-id="5dbb4-112">Просмотрите этот видеоролик в виде одной минуты, чтобы увидеть Лаборатория сценариев в действии.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="5dbb4-113">[![Предварительный просмотр видео, в котором показана Лаборатория скриптов, работающая в Excel, Word и PowerPoint.] (../images/screenshot-wide-youtube.png 'Видеоролик о предварительном просмотре в лаборатории сценариев')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="5dbb4-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="5dbb4-114">Основные возможности</span><span class="sxs-lookup"><span data-stu-id="5dbb4-114">Key features</span></span>

<span data-ttu-id="5dbb4-115">В разделе script Lab предусмотрен ряд функций, которые помогут вам изучить функциональные возможности API JavaScript для Office и прототипа надстройки.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="5dbb4-116">Обзор примеров</span><span class="sxs-lookup"><span data-stu-id="5dbb4-116">Explore samples</span></span>

<span data-ttu-id="5dbb4-117">Быстро приступите к работе со статьей встроенных примеров фрагментов, демонстрирующих выполнение задач с помощью API.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="5dbb4-118">Вы можете запустить примеры, чтобы сразу увидеть результат в области задач или документе, изучить примеры, чтобы узнать, как работает API, и даже использовать примеры для создания прототипа собственной надстройки.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![Примеры](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="5dbb4-120">Код и стиль</span><span class="sxs-lookup"><span data-stu-id="5dbb4-120">Code and style</span></span>

<span data-ttu-id="5dbb4-121">В дополнение к коду JavaScript или TypeScript, вызывающему API Office JS, каждый фрагмент также содержит HTML-разметку, определяющую содержимое области задач и CSS, определяющую внешний вид области задач.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="5dbb4-122">Вы можете настроить HTML-разметку и CSS, чтобы поэкспериментировать с размещением элементов и стилизацией при создании прототипа области задач для собственной надстройки.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="5dbb4-123">Чтобы вызывать API предварительного просмотра внутри фрагмента, вам потребуется обновить библиотеки фрагментов кода, чтобы использовать бета-версию CDN`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`() и определения `@types/office-js-preview`типов предварительного просмотра.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="5dbb4-124">Кроме того, некоторые API предварительной версии доступны только в том случае, если вы зарегистрировались в [программе предварительной оценки Office](https://products.office.com/office-insider) и у вас установлена сборка Office для участников.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://products.office.com/office-insider) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="5dbb4-125">Сохранение и совместное использование фрагментов</span><span class="sxs-lookup"><span data-stu-id="5dbb4-125">Save and share snippets</span></span>

<span data-ttu-id="5dbb4-126">По умолчанию фрагменты кода, открываемые в лаборатории сценариев, будут сохранены в кэше браузера.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="5dbb4-127">Для окончательного сохранения фрагмента его можно экспортировать в [GitHub](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="5dbb4-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="5dbb4-128">Создайте секретный объект, чтобы сохранить фрагмент исключительно для собственного использования, или создайте общедоступного пользователя, если вы планируете поделиться им с другими пользователями.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![Параметры общего доступа](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="5dbb4-130">Импорт фрагментов кода</span><span class="sxs-lookup"><span data-stu-id="5dbb4-130">Import snippets</span></span>

<span data-ttu-id="5dbb4-131">Вы можете импортировать фрагмент в тестовый сценарий, указав URL-адрес общедоступного [GitHub](https://gist.github.com) , в котором ХРАНИТСЯ фрагмент ямл, или ВСТАВИВ полный ямл для фрагмента.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="5dbb4-132">Эта функция может быть полезна в тех случаях, когда кто-то другой предоставил доступ к своему фрагменту, опубликовав его в GitHub или предоставляя свой фрагмент кода ЯМЛ.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![Параметр "импортировать фрагмент"](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="5dbb4-134">Поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="5dbb4-134">Supported clients</span></span>

<span data-ttu-id="5dbb4-135">Лаборатория скриптов поддерживается для Excel, Word и PowerPoint на следующих клиентах.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="5dbb4-136">Office 2013 или более поздней версии в Windows</span><span class="sxs-lookup"><span data-stu-id="5dbb4-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="5dbb4-137">Office 2016 или более поздней версии на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="5dbb4-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="5dbb4-138">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="5dbb4-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="5dbb4-139">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="5dbb4-139">Next steps</span></span>

<span data-ttu-id="5dbb4-140">Чтобы использовать сценарий "Лаборатория" в Excel, Word или PowerPoint, установите [надстройку "Лаборатория скриптов](https://appsource.microsoft.com/product/office/WA104380862) " из AppSource.</span><span class="sxs-lookup"><span data-stu-id="5dbb4-140">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="5dbb4-141">Вы можете развернуть учебную библиотеку в лаборатории сценариев, дополнив новые фрагменты кода в репозиторий GitHub для [Office – JS: Snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) .</span><span class="sxs-lookup"><span data-stu-id="5dbb4-141">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="5dbb4-142">Когда вы будете готовы создать свою первую надстройку Office, ознакомьтесь с кратким руководством для [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md)или [Project](../quickstarts/project-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="5dbb4-142">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="5dbb4-143">См. также</span><span class="sxs-lookup"><span data-stu-id="5dbb4-143">See also</span></span>

- [<span data-ttu-id="5dbb4-144">Получение лаборатории сценариев</span><span class="sxs-lookup"><span data-stu-id="5dbb4-144">Get Script Lab</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="5dbb4-145">Дополнительные сведения о лаборатории сценариев</span><span class="sxs-lookup"><span data-stu-id="5dbb4-145">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="5dbb4-146">Регистрация в программе для разработки</span><span class="sxs-lookup"><span data-stu-id="5dbb4-146">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)

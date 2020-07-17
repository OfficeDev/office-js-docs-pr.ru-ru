---
title: Изучение API JavaScript для Office с помощью Script Lab
description: Используйте Script Lab для изучения API JS Office и использования функциональности работы с прототипами.
ms.date: 06/10/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: ab2d086551dbfa5063615f505d8cb8aa5a210b7a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094136"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="57b84-103">Изучение API JavaScript для Office с помощью Script Lab</span><span class="sxs-lookup"><span data-stu-id="57b84-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="57b84-104">Надстройки [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) и [Script Lab для Outlook](https://appsource.microsoft.com/product/office/wa200001603), которые можно бесплатно получить в AppSource, дают возможность изучать API JavaScript для Office при работе в приложениях Office, таких как Excel или Outlook.</span><span class="sxs-lookup"><span data-stu-id="57b84-104">The [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) and [Script Lab for Outlook](https://appsource.microsoft.com/product/office/wa200001603) add-ins, available free from AppSource, enable you to explore the Office JavaScript API while you're working in an Office program such as Excel or Outlook.</span></span> <span data-ttu-id="57b84-105">Script Lab — удобный инструмент, который пополнит ваш инструментарий разработки для прототипирования и проверки нужной функциональности собственных надстроек.</span><span class="sxs-lookup"><span data-stu-id="57b84-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your own add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="57b84-106">Что такое Script Lab?</span><span class="sxs-lookup"><span data-stu-id="57b84-106">What is Script Lab?</span></span>

<span data-ttu-id="57b84-107">Script Lab — это инструмент для всех, кто хочет научиться разрабатывать надстройки Office с помощью API JavaScript для Office в Excel, Outlook, Word и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="57b84-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="57b84-108">Благодаря поддержке IntelliSense можно видеть доступные возможности. Этот инструмент построен на платформе Monaco, которая используется решением Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="57b84-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="57b84-109">С помощью Script Lab можно получить доступ к библиотеке примеров, чтобы быстро опробовать доступные функции. Также можно использовать пример в качестве отправной точки для разработки собственного кода.</span><span class="sxs-lookup"><span data-stu-id="57b84-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="57b84-110">Можно даже использовать Script Lab для предварительного ознакомления с API.</span><span class="sxs-lookup"><span data-stu-id="57b84-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="57b84-111">Звучит неплохо?</span><span class="sxs-lookup"><span data-stu-id="57b84-111">Sounds good so far?</span></span> <span data-ttu-id="57b84-112">Посмотрите этот минутный видеоролик, чтобы увидеть Script Lab в действии.</span><span class="sxs-lookup"><span data-stu-id="57b84-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="57b84-113">[![Ознакомительное видео, демонстрирующее работу Script Lab в Excel, Word и PowerPoint.](../images/screenshot-wide-youtube.png 'Ознакомительное видео о Script Lab')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="57b84-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="57b84-114">Основные возможности</span><span class="sxs-lookup"><span data-stu-id="57b84-114">Key features</span></span>

<span data-ttu-id="57b84-115">В Script Lab доступен ряд функций, которые помогут изучить API JavaScript для Office и функциональность прототипов надстроек.</span><span class="sxs-lookup"><span data-stu-id="57b84-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="57b84-116">Изучите примеры</span><span class="sxs-lookup"><span data-stu-id="57b84-116">Explore samples</span></span>

<span data-ttu-id="57b84-117">Встроенные примеры фрагментов кода, демонстрирующие выполнение задач с помощью API, помогут быстро начать работу.</span><span class="sxs-lookup"><span data-stu-id="57b84-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="57b84-118">Можно запускать примеры, чтобы сразу видеть результат в области задач или документе, изучать примеры, чтобы понять принципы действия API, и даже использовать примеры для создания прототипов собственных надстроек.</span><span class="sxs-lookup"><span data-stu-id="57b84-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![Примеры](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="57b84-120">Код и стиль</span><span class="sxs-lookup"><span data-stu-id="57b84-120">Code and style</span></span>

<span data-ttu-id="57b84-121">В дополнение к коду JavaScript или TypeScript, который вызывает API JS для Office, каждый фрагмент также содержит разметку HTML, определяющую содержимое области задач, и таблицы стилей CSS, определяющие внешний вид области задач.</span><span class="sxs-lookup"><span data-stu-id="57b84-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="57b84-122">Можно настроить разметку HTML и  CSS, чтобы поэкспериментировать с размещением и стилем элементов при создании прототипа дизайна панели задач для вашей собственной надстройки.</span><span class="sxs-lookup"><span data-stu-id="57b84-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="57b84-123">Чтобы вызвать API предварительной версии во фрагменте кода, потребуется обновить библиотеки фрагмента кода для использования сети доставки содержимого бета-версии (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) и определения типов предварительной версии `@types/office-js-preview`.</span><span class="sxs-lookup"><span data-stu-id="57b84-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="57b84-124">Кроме того, некоторые API предварительной версии доступны только при наличии регистрации в [программе предварительной оценки Office](https://insider.office.com) и используете сборку Office, предназначенную для участников этой программы.</span><span class="sxs-lookup"><span data-stu-id="57b84-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://insider.office.com) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="57b84-125">Сохранение фрагментов кода и общий доступ к ним</span><span class="sxs-lookup"><span data-stu-id="57b84-125">Save and share snippets</span></span>

<span data-ttu-id="57b84-126">Фрагменты кода, которые вы открываете в Script Lab, по умолчанию сохраняются в кэше браузера.</span><span class="sxs-lookup"><span data-stu-id="57b84-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="57b84-127">Чтобы навсегда сохранить фрагмент кода, можно экспортировать его в [gist GitHub](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="57b84-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="57b84-128">Можно создать секретный gist, чтобы сохранить фрагмент кода только для собственного использования, или создать общедоступный gist, если вы планируете поделиться этим фрагментом кода с другими пользователями.</span><span class="sxs-lookup"><span data-stu-id="57b84-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![Возможности общего доступа](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="57b84-130">Импорт фрагментов кода</span><span class="sxs-lookup"><span data-stu-id="57b84-130">Import snippets</span></span>

<span data-ttu-id="57b84-131">Можно импортировать фрагмент кода в Script Lab, указав URL-адрес общедоступного [gist GitHub](https://gist.github.com), в котором хранится YAML этого фрагмента кода, или вставить полный код YAML этого фрагмента кода.</span><span class="sxs-lookup"><span data-stu-id="57b84-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="57b84-132">Эта функция может оказаться полезной в случае, если кто-то другой поделился с вами своим фрагментом кода, опубликовав его в gist GitHub или предоставив YAML этого фрагмента кода.</span><span class="sxs-lookup"><span data-stu-id="57b84-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![Возможность импорта фрагментов кода](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="57b84-134">Поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="57b84-134">Supported clients</span></span>

<span data-ttu-id="57b84-135">Script Lab поддерживается для Excel, Word и  PowerPoint в следующих клиентах.</span><span class="sxs-lookup"><span data-stu-id="57b84-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="57b84-136">Office 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="57b84-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="57b84-137">Office 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="57b84-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="57b84-138">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="57b84-138">Office on the web</span></span>

<span data-ttu-id="57b84-139">Приложение Script Lab для Outlook доступно в следующих клиентах.</span><span class="sxs-lookup"><span data-stu-id="57b84-139">Script Lab for Outlook is available on the following clients.</span></span>

- <span data-ttu-id="57b84-140">Outlook 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="57b84-140">Outlook 2013 or later on Windows</span></span>
- <span data-ttu-id="57b84-141">Outlook 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="57b84-141">Outlook 2016 or later on Mac</span></span>
- <span data-ttu-id="57b84-142">Outlook в Интернете при использовании браузеров Chrome, Microsoft EDGE или Safari</span><span class="sxs-lookup"><span data-stu-id="57b84-142">Outlook on the web when using Chrome, Microsoft Edge, or Safari browsers</span></span>

<span data-ttu-id="57b84-143">Подробнее см. в соответствующей [записи блога](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/).</span><span class="sxs-lookup"><span data-stu-id="57b84-143">For more details on Script Lab for Outlook, see the related [blog post](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/).</span></span>

## <a name="next-steps"></a><span data-ttu-id="57b84-144">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="57b84-144">Next steps</span></span>

<span data-ttu-id="57b84-145">Чтобы использовать Script Lab в Excel, Word или  PowerPoint, установите [надстройку Script Lab](https://appsource.microsoft.com/product/office/WA104380862) из AppSource.</span><span class="sxs-lookup"><span data-stu-id="57b84-145">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="57b84-146">Чтобы использовать Script Lab для Outlook, установите [надстройку Script Lab для Outlook](https://appsource.microsoft.com/product/office/wa200001603) из AppSource.</span><span class="sxs-lookup"><span data-stu-id="57b84-146">To use Script Lab for Outlook, install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) from AppSource.</span></span>

<span data-ttu-id="57b84-147">Вы можете пополнить библиотеку примеров в Script Lab, добавив новые фрагменты кода в репозиторий GitHub [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span><span class="sxs-lookup"><span data-stu-id="57b84-147">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="57b84-148">Когда вы будете готовы приступить к созданию своей первой надстройки Office, ознакомьтесь с кратким руководством для [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md) или [Project](../quickstarts/project-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="57b84-148">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="57b84-149">См. также</span><span class="sxs-lookup"><span data-stu-id="57b84-149">See also</span></span>

- [<span data-ttu-id="57b84-150">Получение Script Lab для Excel, Word и Powerpoint</span><span class="sxs-lookup"><span data-stu-id="57b84-150">Get Script Lab for Excel, Word, or Powerpoint</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="57b84-151">Получение Script Lab для Outlook</span><span class="sxs-lookup"><span data-stu-id="57b84-151">Get Script Lab for Outlook</span></span>](https://appsource.microsoft.com/product/office/wa200001603)
- [<span data-ttu-id="57b84-152">Подробнее о Script Lab</span><span class="sxs-lookup"><span data-stu-id="57b84-152">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- <span data-ttu-id="57b84-153">[Присоединяйтесь к программе для разработчиков Microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="57b84-153">[Join the Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program)</span></span>
- [<span data-ttu-id="57b84-154">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="57b84-154">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)

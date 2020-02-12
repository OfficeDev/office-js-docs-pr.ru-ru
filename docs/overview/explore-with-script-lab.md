---
title: Изучение API JavaScript для Office с помощью Script Lab
description: Используйте Script Lab для изучения API JS Office и использования функциональности работы с прототипами.
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6b8e344460d11cbd85b44fb9a2ab52ef4785cd18
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950756"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="8ac4b-103">Изучение API JavaScript для Office с помощью Script Lab</span><span class="sxs-lookup"><span data-stu-id="8ac4b-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="8ac4b-104">[Надстройка Script Lab](https://appsource.microsoft.com/product/office/WA104380862), бесплатно доступная в AppSource, дает возможность изучать API JavaScript для Office при работе в приложениях Office, таких как Excel или Word.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-104">The [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862), which is available free from AppSource, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="8ac4b-105">Script Lab — удобный инструмент, который целесообразно добавить в набор средств разработки при работе с прототипами и при проверке нужной функциональности надстроек.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="8ac4b-106">Что такое Script Lab?</span><span class="sxs-lookup"><span data-stu-id="8ac4b-106">What is Script Lab?</span></span>

<span data-ttu-id="8ac4b-107">Script Lab — это инструмент для всех, кто хочет научиться разрабатывать надстройки Office с помощью API JavaScript для Office в Excel, Word и  PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="8ac4b-108">Благодаря поддержке IntelliSense можно видеть доступные возможности. Этот инструмент построен на платформе Monaco, которая используется решением Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="8ac4b-109">С помощью Script Lab можно получить доступ к библиотеке примеров, чтобы быстро опробовать доступные функции. Также можно использовать пример в качестве отправной точки для разработки собственного кода.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="8ac4b-110">Можно даже использовать Script Lab для предварительного ознакомления с API.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="8ac4b-111">Звучит неплохо?</span><span class="sxs-lookup"><span data-stu-id="8ac4b-111">Sounds good so far?</span></span> <span data-ttu-id="8ac4b-112">Посмотрите этот минутный видеоролик, чтобы увидеть Script Lab в действии.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="8ac4b-113">[![Ознакомительное видео, демонстрирующее работу Script Lab в Excel, Word и PowerPoint.](../images/screenshot-wide-youtube.png 'Ознакомительное видео о Script Lab')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="8ac4b-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="8ac4b-114">Основные возможности</span><span class="sxs-lookup"><span data-stu-id="8ac4b-114">Key features</span></span>

<span data-ttu-id="8ac4b-115">В Script Lab доступен ряд функций, которые помогут изучить API JavaScript для Office и функциональность прототипов надстроек.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="8ac4b-116">Изучите примеры</span><span class="sxs-lookup"><span data-stu-id="8ac4b-116">Explore samples</span></span>

<span data-ttu-id="8ac4b-117">Встроенные примеры фрагментов кода, демонстрирующие выполнение задач с помощью API, помогут быстро начать работу.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="8ac4b-118">Можно запускать примеры, чтобы сразу видеть результат в области задач или документе, изучать примеры, чтобы понять принципы действия API, и даже использовать примеры для создания прототипов собственных надстроек.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![Примеры](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="8ac4b-120">Код и стиль</span><span class="sxs-lookup"><span data-stu-id="8ac4b-120">Code and style</span></span>

<span data-ttu-id="8ac4b-121">В дополнение к коду JavaScript или TypeScript, который вызывает API JS для Office, каждый фрагмент также содержит разметку HTML, определяющую содержимое области задач, и таблицы стилей CSS, определяющие внешний вид области задач.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="8ac4b-122">Можно настроить разметку HTML и  CSS, чтобы поэкспериментировать с размещением и стилем элементов при создании прототипа дизайна панели задач для вашей собственной надстройки.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="8ac4b-123">Чтобы вызвать API предварительной версии во фрагменте кода, потребуется обновить библиотеки фрагмента кода для использования сети доставки содержимого бета-версии (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) и определения типов предварительной версии `@types/office-js-preview`.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="8ac4b-124">Кроме того, некоторые API предварительной версии доступны только при наличии регистрации в [программе предварительной оценки Office](https://products.office.com/office-insider) и используете сборку Office, предназначенную для участников этой программы.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://products.office.com/office-insider) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="8ac4b-125">Сохранение фрагментов кода и общий доступ к ним</span><span class="sxs-lookup"><span data-stu-id="8ac4b-125">Save and share snippets</span></span>

<span data-ttu-id="8ac4b-126">Фрагменты кода, которые вы открываете в Script Lab, по умолчанию сохраняются в кэше браузера.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="8ac4b-127">Чтобы навсегда сохранить фрагмент кода, можно экспортировать его в [gist GitHub](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="8ac4b-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="8ac4b-128">Можно создать секретный gist, чтобы сохранить фрагмент кода только для собственного использования, или создать общедоступный gist, если вы планируете поделиться этим фрагментом кода с другими пользователями.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![Возможности общего доступа](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="8ac4b-130">Импорт фрагментов кода</span><span class="sxs-lookup"><span data-stu-id="8ac4b-130">Import snippets</span></span>

<span data-ttu-id="8ac4b-131">Можно импортировать фрагмент кода в Script Lab, указав URL-адрес общедоступного [gist GitHub](https://gist.github.com), в котором хранится YAML этого фрагмента кода, или вставить полный код YAML этого фрагмента кода.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="8ac4b-132">Эта функция может оказаться полезной в случае, если кто-то другой поделился с вами своим фрагментом кода, опубликовав его в gist GitHub или предоставив YAML этого фрагмента кода.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![Возможность импорта фрагментов кода](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="8ac4b-134">Поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="8ac4b-134">Supported clients</span></span>

<span data-ttu-id="8ac4b-135">Script Lab поддерживается для Excel, Word и  PowerPoint в следующих клиентах.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="8ac4b-136">Office 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="8ac4b-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="8ac4b-137">Office 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="8ac4b-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="8ac4b-138">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="8ac4b-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="8ac4b-139">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="8ac4b-139">Next steps</span></span>

<span data-ttu-id="8ac4b-140">Чтобы использовать Script Lab в Excel, Word или  PowerPoint, установите [надстройку Script Lab](https://appsource.microsoft.com/product/office/WA104380862) из AppSource.</span><span class="sxs-lookup"><span data-stu-id="8ac4b-140">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="8ac4b-141">Вы можете пополнить библиотеку примеров в Script Lab, добавив новые фрагменты кода в репозиторий GitHub [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span><span class="sxs-lookup"><span data-stu-id="8ac4b-141">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="8ac4b-142">Когда вы будете готовы приступить к созданию своей первой надстройки Office, ознакомьтесь с кратким руководством для [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md) или [Project](../quickstarts/project-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="8ac4b-142">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8ac4b-143">См. также</span><span class="sxs-lookup"><span data-stu-id="8ac4b-143">See also</span></span>

- [<span data-ttu-id="8ac4b-144">Получить Script Lab</span><span class="sxs-lookup"><span data-stu-id="8ac4b-144">Get Script Lab</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="8ac4b-145">Подробнее о Script Lab</span><span class="sxs-lookup"><span data-stu-id="8ac4b-145">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="8ac4b-146">Присоединяйтесь к программе для разработчиков Office 365</span><span class="sxs-lookup"><span data-stu-id="8ac4b-146">Join the Office 365 Developer Program</span></span>](https://developer.microsoft.com/office/dev-program)
- [<span data-ttu-id="8ac4b-147">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="8ac4b-147">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)

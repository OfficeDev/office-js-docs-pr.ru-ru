---
title: Начните отсюда! Руководство для начинающих, делающих Office-надстройки
description: Рекомендуемый путь для начинающих через учебные ресурсы для надстроек Office.
ms.date: 04/16/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 026f90ea62960cbbf5ab4420d40a4a9165139cae
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547623"
---
# <a name="start-here-a-guide-for-beginners-making-office-add-ins"></a><span data-ttu-id="feec8-104">Начните отсюда!</span><span class="sxs-lookup"><span data-stu-id="feec8-104">Start Here!</span></span> <span data-ttu-id="feec8-105">Руководство для начинающих, делающих Office-надстройки</span><span class="sxs-lookup"><span data-stu-id="feec8-105">A guide for beginners making Office Add-ins</span></span>

<span data-ttu-id="feec8-106">Хотите начать создавать собственные кроссплатформенные расширения Office?</span><span class="sxs-lookup"><span data-stu-id="feec8-106">Want to get started building your own cross-platform Office extensions?</span></span> <span data-ttu-id="feec8-107">Следующие шаги покажут вам, что читать в первую очередь, какие инструменты установить и какие учебные пособия рекомендуется выполнить.</span><span class="sxs-lookup"><span data-stu-id="feec8-107">The following steps show you what to read first, what tools to install, and recommended tutorials to complete.</span></span>

## <a name="step-0-prerequisites"></a><span data-ttu-id="feec8-108">Шаг 0. Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="feec8-108">Step 0: Prerequisites</span></span>

- <span data-ttu-id="feec8-109">Надстройки Office - это веб-приложения, встроенные в Office.</span><span class="sxs-lookup"><span data-stu-id="feec8-109">Office Add-ins are essentially web applications embedded in Office.</span></span> <span data-ttu-id="feec8-110">Итак, сначала вы должны иметь общее представление о веб-приложениях и о том, как они размещаются в сети.</span><span class="sxs-lookup"><span data-stu-id="feec8-110">So, you should first have a basic understanding of web applications and how they are hosted on the web.</span></span> <span data-ttu-id="feec8-111">Об этом огромное количество информации в Интернете, книгах и онлайн-курсах.</span><span class="sxs-lookup"><span data-stu-id="feec8-111">There is an enormous amount of information about this on the Internet, in books, and in online courses.</span></span> <span data-ttu-id="feec8-112">Хороший способ начать, если у вас нет предварительных знаний о веб-приложениях, - это поиск "Что такое веб-приложение?"</span><span class="sxs-lookup"><span data-stu-id="feec8-112">A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?"</span></span> <span data-ttu-id="feec8-113">в Bing.</span><span class="sxs-lookup"><span data-stu-id="feec8-113">on Bing.</span></span>
- <span data-ttu-id="feec8-114">Основной язык программирования, который вы будете использовать при создании надстроек Office, - это JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="feec8-114">The primary programming language you will use in creating Office Add-ins is JavaScript or TypeScript.</span></span> <span data-ttu-id="feec8-115">Вы можете думать о TypeScript как о строго типизированной версии JavaScript.</span><span class="sxs-lookup"><span data-stu-id="feec8-115">You can think of TypeScript as a strongly-typed version of JavaScript.</span></span> <span data-ttu-id="feec8-116">Если вы не знакомы ни с одним из этих языков, но у вас есть опыт работы с VBA, VB.Net, C#, вам, вероятно, будет легче освоить TypeScript.</span><span class="sxs-lookup"><span data-stu-id="feec8-116">If you are not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you will probably find TypeScript easier to learn.</span></span> <span data-ttu-id="feec8-117">Опять же, есть много информации об этих языках в Интернете, книгах и онлайн-курсах.</span><span class="sxs-lookup"><span data-stu-id="feec8-117">Again, there is a wealth of information about these languages on the Internet, in books, and in online courses.</span></span>

## <a name="step-1-begin-with-fundamentals"></a><span data-ttu-id="feec8-118">Шаг 1. Начните с основ</span><span class="sxs-lookup"><span data-stu-id="feec8-118">Step 1: Begin with fundamentals</span></span>

<span data-ttu-id="feec8-119">Мы знаем, что вам не терпится начать программирование, но есть некоторые вещи о надстройках Office, которые вы должны прочитать, прежде чем открывать свою IDE или редактор кода.</span><span class="sxs-lookup"><span data-stu-id="feec8-119">We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.</span></span>

- <span data-ttu-id="feec8-120">[Обзор платформы надстроек Office](office-add-ins.md): узнайте, что такое надстройки Office Web и чем они отличаются от более старых способов расширения Office, таких как надстройки VSTO.</span><span class="sxs-lookup"><span data-stu-id="feec8-120">[Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.</span></span>
- <span data-ttu-id="feec8-121">[Создание надстроек Office](office-add-ins-fundamentals.md): Ознакомьтесь с обзором разработки и жизненного цикла надстроек Office, включая инструменты, создание пользовательского интерфейса надстройки и использование API-интерфейсов JavaScript для взаимодействия с документом Office.</span><span class="sxs-lookup"><span data-stu-id="feec8-121">[Building Office Add-ins](office-add-ins-fundamentals.md): Get an overview of Office add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.</span></span>

<span data-ttu-id="feec8-122">В этих статьях много ссылок, но если вы новичок в надстройках Office, мы рекомендуем вам вернуться сюда после прочтения и перейти к следующему разделу.</span><span class="sxs-lookup"><span data-stu-id="feec8-122">There are a lot of links in those articles, but if you're a beginner with Office Add-ins, we recommend that you come back here when you've read them and continue with the next section.</span></span>

## <a name="step-2-install-tools-and-create-your-first-add-in"></a><span data-ttu-id="feec8-123">Шаг 2. Установите инструменты и создайте свою первую надстройку.</span><span class="sxs-lookup"><span data-stu-id="feec8-123">Step 2: Install tools and create your first add-in</span></span>

<span data-ttu-id="feec8-124">Теперь у вас есть общая картина, так что погрузитесь с одним из наших быстрых стартов.</span><span class="sxs-lookup"><span data-stu-id="feec8-124">You've got the big picture now, so dive in with one of our quick starts.</span></span> <span data-ttu-id="feec8-125">В целях изучения платформы мы рекомендуем быстрый запуск Excel.</span><span class="sxs-lookup"><span data-stu-id="feec8-125">For purposes of learning the platform, we recommend the Excel quick start.</span></span> <span data-ttu-id="feec8-126">Существует версия, основанная на Visual Studio, и версия, основанная на Node.js и Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="feec8-126">There is a version that is based on Visual Studio and a version that is based in Node.js and Visual Studio Code.</span></span>

- [<span data-ttu-id="feec8-127">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="feec8-127">Visual Studio</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="feec8-128">Node.js и Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="feec8-128">Node.js and Visual Studio Code</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a><span data-ttu-id="feec8-129">Шаг 3. Код</span><span class="sxs-lookup"><span data-stu-id="feec8-129">Step 3: Code</span></span>

<span data-ttu-id="feec8-130">Вы не можете научиться водить, читая руководство пользователя, поэтому начните программировать с этого [учебника Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="feec8-130">You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md).</span></span> <span data-ttu-id="feec8-131">Вы будете использовать библиотеку Office JavaScript и немного XML в манифесте надстроек.</span><span class="sxs-lookup"><span data-stu-id="feec8-131">You'll be using the Office JavaScript library and some XML in the add-in's manifest.</span></span> <span data-ttu-id="feec8-132">Нет необходимости запоминать что-либо, потому что на следующих шагах вы получите больше информации об обоих.</span><span class="sxs-lookup"><span data-stu-id="feec8-132">There's no need to memorize anything, because you'll be getting more background about both in a later steps.</span></span>

## <a name="step-4-understand-the-javascript-library"></a><span data-ttu-id="feec8-133">Шаг 4. Понять библиотеки JavaScript</span><span class="sxs-lookup"><span data-stu-id="feec8-133">Step 4: Understand the JavaScript library</span></span>

<span data-ttu-id="feec8-134">Во-первых, вы можете получить общее представление о библиотеке JavaScript Office с этим учебным пособием от Microsoft Learn: [Понимание API-интерфейсов Office JavaScript](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).</span><span class="sxs-lookup"><span data-stu-id="feec8-134">First, get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).</span></span>

<span data-ttu-id="feec8-135">Затем изучите API-интерфейсы Office JavaScript с помощью нашего [инструмента Script Lab](explore-with-script-lab.md) - песочницы для запуска и изучения API-интерфейсов.</span><span class="sxs-lookup"><span data-stu-id="feec8-135">Then explore the Office JavaScript APIs with our [the Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.</span></span>

## <a name="step-5-understand-the-manifest"></a><span data-ttu-id="feec8-136">Шаг 5: Понять манифест</span><span class="sxs-lookup"><span data-stu-id="feec8-136">Step 5: Understand the manifest</span></span>

<span data-ttu-id="feec8-137">Получите представление о целях манифеста надстройки и ознакомьтесь с его разметкой XML в [манифесте надстроек Office XML](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="feec8-137">Get an understanding of the purposes of the add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="feec8-138">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="feec8-138">Next Steps</span></span>

<span data-ttu-id="feec8-139">Поздравляем с окончанием курса обучения начинающих для надстроек Office!</span><span class="sxs-lookup"><span data-stu-id="feec8-139">Congratulations on finishing the beginner's learning path for Office Add-ins!</span></span> <span data-ttu-id="feec8-140">Вот несколько предложений для дальнейшего изучения нашей документации:</span><span class="sxs-lookup"><span data-stu-id="feec8-140">Here are some suggestions for further exploration of our documentation:</span></span>

- <span data-ttu-id="feec8-141">Учебные материалы и краткое руководство для других приложений Office.</span><span class="sxs-lookup"><span data-stu-id="feec8-141">Tutorials or quick starts for other Office applications:</span></span>

  - [<span data-ttu-id="feec8-142">Руководство по началу работы с OneNote</span><span class="sxs-lookup"><span data-stu-id="feec8-142">OneNote quick start</span></span>](../quickstarts/onenote-quickstart.md)
  - [<span data-ttu-id="feec8-143">Учебник по Outlook</span><span class="sxs-lookup"><span data-stu-id="feec8-143">Outlook tutorial</span></span>](/outlook/add-ins/addin-tutorial)
  - [<span data-ttu-id="feec8-144">Учебник по PowerPoint</span><span class="sxs-lookup"><span data-stu-id="feec8-144">PowerPoint tutorial</span></span>](../tutorials/powerpoint-tutorial.md)
  - [<span data-ttu-id="feec8-145">Руководство по началу работы с Project</span><span class="sxs-lookup"><span data-stu-id="feec8-145">Project quick start</span></span>](../quickstarts/project-quickstart.md)
  - [<span data-ttu-id="feec8-146">Учебник по Word</span><span class="sxs-lookup"><span data-stu-id="feec8-146">Word tutorial</span></span>](../tutorials/word-tutorial.md)

- <span data-ttu-id="feec8-147">Другие важные темы:</span><span class="sxs-lookup"><span data-stu-id="feec8-147">Other important subjects:</span></span>

  - [<span data-ttu-id="feec8-148">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="feec8-148">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
  - [<span data-ttu-id="feec8-149">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="feec8-149">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
  - [<span data-ttu-id="feec8-150">Проектирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="feec8-150">Design Office Add-ins</span></span>](../design/add-in-design.md)
  - [<span data-ttu-id="feec8-151">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="feec8-151">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
  - [<span data-ttu-id="feec8-152">Развертывание и публикация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="feec8-152">Deploy and publish Office Add-ins</span></span>](../publish/publish.md)
  - [<span data-ttu-id="feec8-153">Ресурсы</span><span class="sxs-lookup"><span data-stu-id="feec8-153">Resources</span></span>](../resources/resources-links-help.md)

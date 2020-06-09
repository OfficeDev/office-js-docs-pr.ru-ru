---
title: Начните отсюда! Руководство для начинающих, делающих Office-надстройки
description: Рекомендуемый путь для начинающих через учебные ресурсы для надстроек Office.
ms.date: 04/16/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: b62c7a5d2117c52f4bd3f91c1a2e1b735554028e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604504"
---
# <a name="start-here-a-guide-for-beginners-making-office-add-ins"></a><span data-ttu-id="449bd-104">Начните отсюда!</span><span class="sxs-lookup"><span data-stu-id="449bd-104">Start Here!</span></span> <span data-ttu-id="449bd-105">Руководство для начинающих, делающих Office-надстройки</span><span class="sxs-lookup"><span data-stu-id="449bd-105">A guide for beginners making Office Add-ins</span></span>

<span data-ttu-id="449bd-106">Хотите начать создавать собственные кроссплатформенные расширения Office?</span><span class="sxs-lookup"><span data-stu-id="449bd-106">Want to get started building your own cross-platform Office extensions?</span></span> <span data-ttu-id="449bd-107">Следующие шаги покажут вам, что читать в первую очередь, какие инструменты установить и какие учебные пособия рекомендуется выполнить.</span><span class="sxs-lookup"><span data-stu-id="449bd-107">The following steps show you what to read first, what tools to install, and recommended tutorials to complete.</span></span>

> [!NOTE]
> <span data-ttu-id="449bd-108">Если у вас есть опыт создания надстроек VSTO для Office, рекомендуем сразу перейти к статье [Выполните переход! Руководство для авторов надстроек VSTO, создающих веб-надстройки Office](learning-path-transition.md), которая дополняет сведения, приведенные в этой статье.</span><span class="sxs-lookup"><span data-stu-id="449bd-108">If you're experienced in creating VSTO add-ins for Office, we recommend that you immediately turn to [Transition Here! A guide for VSTO add-in creators making Office Web Add-ins](learning-path-transition.md), which is a superset of the information in this article.</span></span>

## <a name="step-0-prerequisites"></a><span data-ttu-id="449bd-109">Шаг 0. Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="449bd-109">Step 0: Prerequisites</span></span>

- <span data-ttu-id="449bd-110">Надстройки Office - это веб-приложения, встроенные в Office.</span><span class="sxs-lookup"><span data-stu-id="449bd-110">Office Add-ins are essentially web applications embedded in Office.</span></span> <span data-ttu-id="449bd-111">Итак, сначала вы должны иметь общее представление о веб-приложениях и о том, как они размещаются в сети.</span><span class="sxs-lookup"><span data-stu-id="449bd-111">So, you should first have a basic understanding of web applications and how they are hosted on the web.</span></span> <span data-ttu-id="449bd-112">Об этом огромное количество информации в Интернете, книгах и онлайн-курсах.</span><span class="sxs-lookup"><span data-stu-id="449bd-112">There is an enormous amount of information about this on the Internet, in books, and in online courses.</span></span> <span data-ttu-id="449bd-113">Хороший способ начать, если у вас нет предварительных знаний о веб-приложениях, - это поиск "Что такое веб-приложение?"</span><span class="sxs-lookup"><span data-stu-id="449bd-113">A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?"</span></span> <span data-ttu-id="449bd-114">в Bing.</span><span class="sxs-lookup"><span data-stu-id="449bd-114">on Bing.</span></span>
- <span data-ttu-id="449bd-115">Основной язык программирования, который вы будете использовать при создании надстроек Office, - это JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="449bd-115">The primary programming language you will use in creating Office Add-ins is JavaScript or TypeScript.</span></span> <span data-ttu-id="449bd-116">Вы можете думать о TypeScript как о строго типизированной версии JavaScript.</span><span class="sxs-lookup"><span data-stu-id="449bd-116">You can think of TypeScript as a strongly-typed version of JavaScript.</span></span> <span data-ttu-id="449bd-117">Если вы не знакомы ни с одним из этих языков, но у вас есть опыт работы с VBA, VB.Net, C#, вам, вероятно, будет легче освоить TypeScript.</span><span class="sxs-lookup"><span data-stu-id="449bd-117">If you are not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you will probably find TypeScript easier to learn.</span></span> <span data-ttu-id="449bd-118">Опять же, есть много информации об этих языках в Интернете, книгах и онлайн-курсах.</span><span class="sxs-lookup"><span data-stu-id="449bd-118">Again, there is a wealth of information about these languages on the Internet, in books, and in online courses.</span></span>

## <a name="step-1-begin-with-fundamentals"></a><span data-ttu-id="449bd-119">Шаг 1. Начните с основ</span><span class="sxs-lookup"><span data-stu-id="449bd-119">Step 1: Begin with fundamentals</span></span>

<span data-ttu-id="449bd-120">Мы знаем, что вам не терпится начать программирование, но есть некоторые вещи о надстройках Office, которые вы должны прочитать, прежде чем открывать свою IDE или редактор кода.</span><span class="sxs-lookup"><span data-stu-id="449bd-120">We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.</span></span>

- <span data-ttu-id="449bd-121">[Обзор платформы надстроек Office](office-add-ins.md): узнайте, что такое надстройки Office Web и чем они отличаются от более старых способов расширения Office, таких как надстройки VSTO.</span><span class="sxs-lookup"><span data-stu-id="449bd-121">[Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.</span></span>
- <span data-ttu-id="449bd-122">[Создание надстроек Office](office-add-ins-fundamentals.md): Ознакомьтесь с обзором разработки и жизненного цикла надстроек Office, включая инструменты, создание пользовательского интерфейса надстройки и использование API-интерфейсов JavaScript для взаимодействия с документом Office.</span><span class="sxs-lookup"><span data-stu-id="449bd-122">[Building Office Add-ins](office-add-ins-fundamentals.md): Get an overview of Office add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.</span></span>

<span data-ttu-id="449bd-123">В этих статьях много ссылок, но если вы новичок в надстройках Office, мы рекомендуем вам вернуться сюда после прочтения и перейти к следующему разделу.</span><span class="sxs-lookup"><span data-stu-id="449bd-123">There are a lot of links in those articles, but if you're a beginner with Office Add-ins, we recommend that you come back here when you've read them and continue with the next section.</span></span>

## <a name="step-2-install-tools-and-create-your-first-add-in"></a><span data-ttu-id="449bd-124">Шаг 2. Установите инструменты и создайте свою первую надстройку.</span><span class="sxs-lookup"><span data-stu-id="449bd-124">Step 2: Install tools and create your first add-in</span></span>

<span data-ttu-id="449bd-125">Теперь у вас есть общая картина, так что погрузитесь с одним из наших быстрых стартов.</span><span class="sxs-lookup"><span data-stu-id="449bd-125">You've got the big picture now, so dive in with one of our quick starts.</span></span> <span data-ttu-id="449bd-126">В целях изучения платформы мы рекомендуем быстрый запуск Excel.</span><span class="sxs-lookup"><span data-stu-id="449bd-126">For purposes of learning the platform, we recommend the Excel quick start.</span></span> <span data-ttu-id="449bd-127">Существует версия, основанная на Visual Studio, и версия, основанная на Node.js и Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="449bd-127">There is a version that is based on Visual Studio and a version that is based in Node.js and Visual Studio Code.</span></span>

- [<span data-ttu-id="449bd-128">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="449bd-128">Visual Studio</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="449bd-129">Node.js и Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="449bd-129">Node.js and Visual Studio Code</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a><span data-ttu-id="449bd-130">Шаг 3. Код</span><span class="sxs-lookup"><span data-stu-id="449bd-130">Step 3: Code</span></span>

<span data-ttu-id="449bd-131">Вы не можете научиться водить, читая руководство пользователя, поэтому начните программировать с этого [учебника Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="449bd-131">You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md).</span></span> <span data-ttu-id="449bd-132">Вы будете использовать библиотеку Office JavaScript и немного XML в манифесте надстроек.</span><span class="sxs-lookup"><span data-stu-id="449bd-132">You'll be using the Office JavaScript library and some XML in the add-in's manifest.</span></span> <span data-ttu-id="449bd-133">Нет необходимости запоминать что-либо, потому что на следующих шагах вы получите больше информации об обоих.</span><span class="sxs-lookup"><span data-stu-id="449bd-133">There's no need to memorize anything, because you'll be getting more background about both in a later steps.</span></span>

## <a name="step-4-understand-the-javascript-library"></a><span data-ttu-id="449bd-134">Шаг 4. Знакомство с библиотекой JavaScript</span><span class="sxs-lookup"><span data-stu-id="449bd-134">Step 4: Understand the JavaScript library</span></span>

<span data-ttu-id="449bd-135">Во-первых, вы можете получить общее представление о библиотеке JavaScript Office с этим учебным пособием от Microsoft Learn: [Понимание API-интерфейсов Office JavaScript](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).</span><span class="sxs-lookup"><span data-stu-id="449bd-135">First, get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).</span></span>

<span data-ttu-id="449bd-136">Затем изучите API-интерфейсы Office JavaScript с помощью нашего [инструмента Script Lab](explore-with-script-lab.md) - песочницы для запуска и изучения API-интерфейсов.</span><span class="sxs-lookup"><span data-stu-id="449bd-136">Then explore the Office JavaScript APIs with our [the Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.</span></span>

## <a name="step-5-understand-the-manifest"></a><span data-ttu-id="449bd-137">Шаг 5. Знакомство с манифестом</span><span class="sxs-lookup"><span data-stu-id="449bd-137">Step 5: Understand the manifest</span></span>

<span data-ttu-id="449bd-138">Получите представление о целях манифеста надстройки и ознакомьтесь с его разметкой XML в [манифесте надстроек Office XML](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="449bd-138">Get an understanding of the purposes of the add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="449bd-139">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="449bd-139">Next Steps</span></span>

<span data-ttu-id="449bd-140">Поздравляем с окончанием курса обучения начинающих для надстроек Office!</span><span class="sxs-lookup"><span data-stu-id="449bd-140">Congratulations on finishing the beginner's learning path for Office Add-ins!</span></span> <span data-ttu-id="449bd-141">Вот несколько предложений для дальнейшего изучения нашей документации:</span><span class="sxs-lookup"><span data-stu-id="449bd-141">Here are some suggestions for further exploration of our documentation:</span></span>

- <span data-ttu-id="449bd-142">Учебные материалы и краткое руководство для других приложений Office.</span><span class="sxs-lookup"><span data-stu-id="449bd-142">Tutorials or quick starts for other Office applications:</span></span>

  - [<span data-ttu-id="449bd-143">Руководство по началу работы с OneNote</span><span class="sxs-lookup"><span data-stu-id="449bd-143">OneNote quick start</span></span>](../quickstarts/onenote-quickstart.md)
  - [<span data-ttu-id="449bd-144">Учебник по Outlook</span><span class="sxs-lookup"><span data-stu-id="449bd-144">Outlook tutorial</span></span>](/outlook/add-ins/addin-tutorial)
  - [<span data-ttu-id="449bd-145">Учебник по PowerPoint</span><span class="sxs-lookup"><span data-stu-id="449bd-145">PowerPoint tutorial</span></span>](../tutorials/powerpoint-tutorial.md)
  - [<span data-ttu-id="449bd-146">Руководство по началу работы с Project</span><span class="sxs-lookup"><span data-stu-id="449bd-146">Project quick start</span></span>](../quickstarts/project-quickstart.md)
  - [<span data-ttu-id="449bd-147">Учебник по Word</span><span class="sxs-lookup"><span data-stu-id="449bd-147">Word tutorial</span></span>](../tutorials/word-tutorial.md)

- <span data-ttu-id="449bd-148">Другие важные темы:</span><span class="sxs-lookup"><span data-stu-id="449bd-148">Other important subjects:</span></span>

  - [<span data-ttu-id="449bd-149">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="449bd-149">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
  - [<span data-ttu-id="449bd-150">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="449bd-150">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
  - [<span data-ttu-id="449bd-151">Проектирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="449bd-151">Design Office Add-ins</span></span>](../design/add-in-design.md)
  - [<span data-ttu-id="449bd-152">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="449bd-152">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
  - [<span data-ttu-id="449bd-153">Развертывание и публикация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="449bd-153">Deploy and publish Office Add-ins</span></span>](../publish/publish.md)
  - [<span data-ttu-id="449bd-154">Ресурсы</span><span class="sxs-lookup"><span data-stu-id="449bd-154">Resources</span></span>](../resources/resources-links-help.md)

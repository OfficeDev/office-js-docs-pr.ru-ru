---
title: Разработка надстроек Office с помощью Visual Studio
description: Разработка надстроек Office с помощью Visual Studio
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 9f250078a4da80dea3276c51a2183a072da44f81
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42162813"
---
# <a name="develop-office-add-ins-with-visual-studio"></a><span data-ttu-id="91cee-103">Разработка надстроек Office с помощью Visual Studio</span><span class="sxs-lookup"><span data-stu-id="91cee-103">Develop Office Add-ins with Visual Studio</span></span>

<span data-ttu-id="91cee-104">В этой статье описано, как использовать Visual Studio для разработки надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="91cee-104">This article describes how to use Visual Studio to develop an Office Add-in.</span></span> <span data-ttu-id="91cee-105">Если надстройка уже создана, можно перейти к разделу [Разработка надстройки с помощью Visual Studio](#develop-the-add-in-using-visual-studio).</span><span class="sxs-lookup"><span data-stu-id="91cee-105">If you've already created your add-in, you can skip ahead to the [Develop the add-in using Visual Studio](#develop-the-add-in-using-visual-studio) section.</span></span>

> [!NOTE]
> <span data-ttu-id="91cee-106">Вместо Visual Studio можно использовать генератор Yeoman для надстроек Office и VS Code для создания надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="91cee-106">As an alternative to using Visual Studio, you may choose to use the Yeoman generator for Office Add-ins and VS Code to create an Office Add-in.</span></span> <span data-ttu-id="91cee-107">Дополнительные сведения о выборе средств создания см. в разделе [Создание надстроек Office](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in).</span><span class="sxs-lookup"><span data-stu-id="91cee-107">For more information about this choice, see [Creating an Office Add-in](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in).</span></span>

## <a name="create-the-add-in-project-using-visual-studio"></a><span data-ttu-id="91cee-108">Создание проекта надстройки с помощью Visual Studio</span><span class="sxs-lookup"><span data-stu-id="91cee-108">Create the add-in project using Visual Studio</span></span>

<span data-ttu-id="91cee-109">С помощью Visual Studio можно создавать надстройки Office для Excel, Outlook, Word и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="91cee-109">Visual Studio can be used to create Office Add-ins for Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="91cee-110">Проект надстройки Office создается в рамках решения Visual Studio и использует HTML, CSS и JavaScript.</span><span class="sxs-lookup"><span data-stu-id="91cee-110">An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript.</span></span> <span data-ttu-id="91cee-111">Чтобы создать надстройку Office с помощью Visual Studio, следуйте указаниям из краткого руководства, соответствующего типу надстройки, которую нужно создать.</span><span class="sxs-lookup"><span data-stu-id="91cee-111">To create an Office Add-in with Visual Studio, follow instructions in the quick start that corresponds to the add-in you'd like to create:</span></span>

- [<span data-ttu-id="91cee-112">Краткое руководство по началу работы с Excel</span><span class="sxs-lookup"><span data-stu-id="91cee-112">Excel quick start</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="91cee-113">Краткое руководство по началу работы с Outlook</span><span class="sxs-lookup"><span data-stu-id="91cee-113">Outlook quick start</span></span>](../quickstarts/outlook-quickstart.md?tabs=visualstudio)
- [<span data-ttu-id="91cee-114">Краткое руководство по началу работы с Word</span><span class="sxs-lookup"><span data-stu-id="91cee-114">Word quick start</span></span>](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [<span data-ttu-id="91cee-115">Краткое руководство по началу работы с PowerPoint</span><span class="sxs-lookup"><span data-stu-id="91cee-115">PowerPoint quick start</span></span>](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

<span data-ttu-id="91cee-116">В Visual Studio не поддерживается создание надстроек Office для OneNote и Project.</span><span class="sxs-lookup"><span data-stu-id="91cee-116">Visual Studio doesn't support creating Office Add-ins for OneNote or Project.</span></span> <span data-ttu-id="91cee-117">Чтобы создать надстройки Office для любого из этих ведущих приложений потребуется использовать генератор Yeoman для надстроек Office, как описано в [кратком руководстве по началу работы с OneNote](../quickstarts/onenote-quickstart.md) и в [кратком руководстве по началу работы с Project](../quickstarts/project-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="91cee-117">To create Office Add-ins for either of these hosts, you'll need to use the Yeoman generator for Office Add-ins, as described in the [OneNote quick start](../quickstarts/onenote-quickstart.md) or the [Project quick start](../quickstarts/project-quickstart.md).</span></span>

## <a name="develop-the-add-in-using-visual-studio"></a><span data-ttu-id="91cee-118">Разработка надстройки с помощью Visual Studio</span><span class="sxs-lookup"><span data-stu-id="91cee-118">Develop the add-in using Visual Studio</span></span>

<span data-ttu-id="91cee-119">В Visual Studio создается простая надстройка с ограниченными возможностями.</span><span class="sxs-lookup"><span data-stu-id="91cee-119">Visual Studio creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="91cee-120">Можно настроить надстройку, отредактировав файлы [манифеста](add-in-manifests.md), HTML, JavaScript и CSS в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="91cee-120">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript, and CSS files in Visual Studio.</span></span> <span data-ttu-id="91cee-121">Общее описание структуры проекта и файлов в проекте надстройки, создаваемом в Visual Studio, см. в справочнике по Visual Studio в составе краткого руководства по началу работы, с помощью которого вы создали надстройку.</span><span class="sxs-lookup"><span data-stu-id="91cee-121">For a high-level description of the project structure and files in the add-in project that Visual Studio creates, see the Visual Studio guidance within the quick start that you completed to create your add-in.</span></span> 

> [!TIP]
> <span data-ttu-id="91cee-122">Надстройка Office представляет собой веб-приложение, поэтому для изменения надстройки требуются базовые навыки веб-разработки.</span><span class="sxs-lookup"><span data-stu-id="91cee-122">Because an Office Add-in is a web application, you'll need at least basic web development skills to customize your add-in.</span></span> <span data-ttu-id="91cee-123">Если вы впервые работаете с JavaScript, рекомендуем прочесть [учебник Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="91cee-123">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

<span data-ttu-id="91cee-124">Чтобы настроить надстройку, потребуется понять принципы, описанные в разделе [Основные принципы > Разработка](develop-overview.md) этой документации, а также принципы, описанные в соответствующем разделе документации ведущего приложения, для которого вы создаете надстройку (например, [Excel](../excel/index.md)).</span><span class="sxs-lookup"><span data-stu-id="91cee-124">To customize your add-in, you'll need to understand concepts described in the [Core concepts > Develop](develop-overview.md) area of this documentation, as well as concepts described in the host-specific area of documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.md)).</span></span> 

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="91cee-125">Тестирование и отладка надстройки</span><span class="sxs-lookup"><span data-stu-id="91cee-125">Test and debug the add-in</span></span>

<span data-ttu-id="91cee-126">Методы тестирования, отладки и устранения неполадок надстроек Office зависят от платформы.</span><span class="sxs-lookup"><span data-stu-id="91cee-126">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="91cee-127">Дополнительные сведения см. в статьях [Отладка надстроек Office в Visual Studio](debug-office-add-ins-in-visual-studio.md) и [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="91cee-127">For more information, see [Debug Office Add-ins in Visual Studio](debug-office-add-ins-in-visual-studio.md) and [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="91cee-128">Публикация надстройки</span><span class="sxs-lookup"><span data-stu-id="91cee-128">Publish the add-in</span></span>

<span data-ttu-id="91cee-129">Надстройка Office состоит из веб-приложения и файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="91cee-129">An Office Add-in consists of a web application and a manifest file.</span></span> <span data-ttu-id="91cee-130">Веб-приложение определяет пользовательский интерфейс и функции надстройки, а манифест указывает расположение веб-приложения и определяет параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="91cee-130">The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.</span></span>

<span data-ttu-id="91cee-131">В процессе разработки надстройки в Visual Studio эта надстройка запускается на локальном веб-сервере (`localhost`).</span><span class="sxs-lookup"><span data-stu-id="91cee-131">While you're developing your add-in in Visual Studio, your add-in runs on your local web server (`localhost`).</span></span> <span data-ttu-id="91cee-132">Если надстройка работает нужным образом и вы готовы опубликовать ее для доступа других пользователей, выполните следующие действия:</span><span class="sxs-lookup"><span data-stu-id="91cee-132">When your add-in is working as desired and you're ready to publish it for other users to access, you'll need to complete the following steps:</span></span>

1. <span data-ttu-id="91cee-133">Разверните веб-приложение на веб-сервере или в службе веб-хостинга (например, Microsoft Azure).</span><span class="sxs-lookup"><span data-stu-id="91cee-133">Deploy the web application to a web server or web hosting service (for example, Microsoft Azure).</span></span>
2. <span data-ttu-id="91cee-134">Обновите манифест, указав URL-адрес развернутого приложения.</span><span class="sxs-lookup"><span data-stu-id="91cee-134">Update the manifest to specify the URL of the deployed application.</span></span> 
3. <span data-ttu-id="91cee-135">Выберите метод [развертывания надстройки Office](../publish/publish.md) и следуйте инструкциям, чтобы опубликовать файл манифеста.</span><span class="sxs-lookup"><span data-stu-id="91cee-135">Choose the method you'd like to use to [deploy your Office Add-in](../publish/publish.md), and follow the instructions to publish the manifest file.</span></span>

## <a name="see-also"></a><span data-ttu-id="91cee-136">См. также</span><span class="sxs-lookup"><span data-stu-id="91cee-136">See also</span></span>

- [<span data-ttu-id="91cee-137">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="91cee-137">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="91cee-138">Основные принципы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="91cee-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="91cee-139">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="91cee-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="91cee-140">Проектирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="91cee-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="91cee-141">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="91cee-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="91cee-142">Публикация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="91cee-142">Publish Office Add-ins</span></span>](../publish/publish.md)
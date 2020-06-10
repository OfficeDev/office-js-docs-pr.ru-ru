---
title: Настройка среды разработки
description: Настройка среды разработки для создания надстроек Office
ms.date: 04/03/2020
localization_priority: Normal
ms.openlocfilehash: f44f8e48aec402f0ffa6327732613a902ea0cfe6
ms.sourcegitcommit: 19312a54f47a17988ffa86359218a504713f9f09
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/10/2020
ms.locfileid: "44679355"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="0047b-103">Настройка среды разработки</span><span class="sxs-lookup"><span data-stu-id="0047b-103">Set up your development environment</span></span>

<span data-ttu-id="0047b-104">Это руководство поможет вам настроить средства для создания надстроек Office, выполнив следующие краткие руководства по началу.</span><span class="sxs-lookup"><span data-stu-id="0047b-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="0047b-105">Вам потребуется установить средства из приведенного ниже списка.</span><span class="sxs-lookup"><span data-stu-id="0047b-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="0047b-106">Если у вас уже есть эти компоненты, вы можете начать краткий запуск, например, на [панели быстрого запуска Excel](../quickstarts/excel-quickstart-react.md).</span><span class="sxs-lookup"><span data-stu-id="0047b-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="0047b-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="0047b-107">Node.js</span></span>
- <span data-ttu-id="0047b-108">npm</span><span class="sxs-lookup"><span data-stu-id="0047b-108">npm</span></span>
- <span data-ttu-id="0047b-109">Учетная запись Office 365 (версия для подписки на Office)</span><span class="sxs-lookup"><span data-stu-id="0047b-109">An Office 365 (the subscription version of Office) account</span></span>
- <span data-ttu-id="0047b-110">Любой редактор кода</span><span class="sxs-lookup"><span data-stu-id="0047b-110">A code editor of your choice</span></span>

<span data-ttu-id="0047b-111">В этом руководстве предполагается, что вы знаете, как использовать средство командной строки.</span><span class="sxs-lookup"><span data-stu-id="0047b-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="0047b-112">Install Node. js</span><span class="sxs-lookup"><span data-stu-id="0047b-112">Install Node.js</span></span>

<span data-ttu-id="0047b-113">Node. js — это среда выполнения JavaScript, которая должна быть разработана для современных надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="0047b-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="0047b-114">Установите Node. js, [загрузив последнюю рекомендуемую версию со своего веб-сайта](https://nodejs.org).</span><span class="sxs-lookup"><span data-stu-id="0047b-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="0047b-115">Следуйте инструкциям по установке для вашей операционной системы.</span><span class="sxs-lookup"><span data-stu-id="0047b-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="0047b-116">Установка NPM</span><span class="sxs-lookup"><span data-stu-id="0047b-116">Install npm</span></span>

<span data-ttu-id="0047b-117">NPM — это реестр программного обеспечения с открытым кодом, из которого загружаются пакеты, используемые при разработке надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="0047b-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="0047b-118">Чтобы установить NPM, выполните следующую команду в командной строке:</span><span class="sxs-lookup"><span data-stu-id="0047b-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="0047b-119">Чтобы проверить, установлен ли у вас NPM, и просмотреть установленную версию, выполните следующую команду в командной строке:</span><span class="sxs-lookup"><span data-stu-id="0047b-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="0047b-120">Вы можете использовать диспетчер версий узла, чтобы можно было переключаться между несколькими версиями Node. js и NPM, но это не является обязательным.</span><span class="sxs-lookup"><span data-stu-id="0047b-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="0047b-121">Для получения дополнительных сведений о том, как это сделать, [обратитесь к разделу инструкции NPM](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span><span class="sxs-lookup"><span data-stu-id="0047b-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="0047b-122">Получение Office 365</span><span class="sxs-lookup"><span data-stu-id="0047b-122">Get Office 365</span></span>

<span data-ttu-id="0047b-123">Если у вас еще нет учетной записи Office 365, вы можете оформить бесплатную возобновляемую подписку на Office 365 на 90 дней, присоединившись к [программе для разработчиков Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="0047b-123">If you don't already have an Office 365 account, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="0047b-124">Установка редактора кода</span><span class="sxs-lookup"><span data-stu-id="0047b-124">Install a code editor</span></span>

<span data-ttu-id="0047b-125">Для создания веб-частей можно использовать любой редактор кода или интерфейс IDE, поддерживающий клиентскую разработку, например:</span><span class="sxs-lookup"><span data-stu-id="0047b-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="0047b-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="0047b-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- <span data-ttu-id="0047b-127">[Atom](https://atom.io);</span><span class="sxs-lookup"><span data-stu-id="0047b-127">[Atom](https://atom.io)</span></span>
- [<span data-ttu-id="0047b-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="0047b-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="0047b-129">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="0047b-129">Next steps</span></span>

<span data-ttu-id="0047b-130">Попробуйте создать собственную надстройку или воспользоваться лабораториями скриптов, чтобы испытать встроенные примеры.</span><span class="sxs-lookup"><span data-stu-id="0047b-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="0047b-131">Создание надстройки Office</span><span class="sxs-lookup"><span data-stu-id="0047b-131">Create an Office add-in</span></span>

<span data-ttu-id="0047b-132">Вы можете быстро создать простую надстройку для Excel, OneNote, Outlook, PowerPoint, Project или Word с помощью [5-минутного краткого руководства по началу работы](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="0047b-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](/office/dev/add-ins/).</span></span> <span data-ttu-id="0047b-133">Если вы уже ознакомились с кратким руководством и хотите создать более сложную надстройку, воспользуйтесь [учебником](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="0047b-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](/office/dev/add-ins/).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="0047b-134">Изучение API с помощью Script Lab</span><span class="sxs-lookup"><span data-stu-id="0047b-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="0047b-135">Изучите библиотеку встроенных примеров в [Script Lab](explore-with-script-lab.md), чтобы ознакомиться с возможностями API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="0047b-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="0047b-136">См. также</span><span class="sxs-lookup"><span data-stu-id="0047b-136">See also</span></span>

- [<span data-ttu-id="0047b-137">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0047b-137">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="0047b-138">Основные принципы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0047b-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="0047b-139">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0047b-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="0047b-140">Проектирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0047b-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="0047b-141">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0047b-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="0047b-142">Публикация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0047b-142">Publish Office Add-ins</span></span>](../publish/publish.md)

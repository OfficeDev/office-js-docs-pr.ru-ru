---
title: Настройка среды разработки
description: Настройка среды разработчика для создания надстройки Office.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 1dd0cc6bb035a0274e36fe9916dcd2481bdf0b39
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234130"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="5b5d8-103">Настройка среды разработки</span><span class="sxs-lookup"><span data-stu-id="5b5d8-103">Set up your development environment</span></span>

<span data-ttu-id="5b5d8-104">Это руководство поможет вам настроить средства для создания надстройки Office, следуя нашим кратким руководствам или руководствам.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="5b5d8-105">Вам потребуется установить средства из приведенного ниже списка.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="5b5d8-106">Если у вас уже установлены эти приложения, вы можете начать быстрое начало работы, например, это краткое начало [Excel React.](../quickstarts/excel-quickstart-react.md)</span><span class="sxs-lookup"><span data-stu-id="5b5d8-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="5b5d8-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="5b5d8-107">Node.js</span></span>
- <span data-ttu-id="5b5d8-108">npm</span><span class="sxs-lookup"><span data-stu-id="5b5d8-108">npm</span></span>
- <span data-ttu-id="5b5d8-109">Учетная запись Microsoft 365, включаемая версию Office по подписке</span><span class="sxs-lookup"><span data-stu-id="5b5d8-109">A Microsoft 365 account which includes the subscription version of Office</span></span>
- <span data-ttu-id="5b5d8-110">Редактор кода по вашему выбору</span><span class="sxs-lookup"><span data-stu-id="5b5d8-110">A code editor of your choice</span></span>

<span data-ttu-id="5b5d8-111">В этом руководстве предполагается, что вы знаете, как использовать средство командной строки.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="5b5d8-112">Установите Node.js.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-112">Install Node.js</span></span>

<span data-ttu-id="5b5d8-113">Node.js является среде запуска JavaScript, необходимо разрабатывать современные надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="5b5d8-114">Установите Node.js, [скачав последнюю рекомендуемую версию с веб-сайта.](https://nodejs.org)</span><span class="sxs-lookup"><span data-stu-id="5b5d8-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="5b5d8-115">Следуйте инструкциям по установке операционной системы.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="5b5d8-116">Установка npm</span><span class="sxs-lookup"><span data-stu-id="5b5d8-116">Install npm</span></span>

<span data-ttu-id="5b5d8-117">npm — это реестр программного обеспечения с открытым кодом, из которого можно скачать пакеты, используемые при разработке надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="5b5d8-118">Чтобы установить npm, в командной строке запустите следующую команду:</span><span class="sxs-lookup"><span data-stu-id="5b5d8-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="5b5d8-119">Чтобы проверить, установлен ли npm, и увидеть установленную версию, в командной строке запустите следующую команду:</span><span class="sxs-lookup"><span data-stu-id="5b5d8-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="5b5d8-120">Вы можете использовать диспетчер версий Node, чтобы разрешить переключение между несколькими версиями Node.js npm, но это не является строго необходимым.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="5b5d8-121">Подробные сведения о том, как это сделать, см. в [инструкциях npm.](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)</span><span class="sxs-lookup"><span data-stu-id="5b5d8-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-microsoft-365"></a><span data-ttu-id="5b5d8-122">Получить Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="5b5d8-122">Get Microsoft 365</span></span>

<span data-ttu-id="5b5d8-123">Если у вас еще нет учетной записи Microsoft 365, вы можете получить бесплатную 90-дневную возобновляемую подписку на Microsoft 365, которая включает все приложения Office, присоединившись к программе для разработчиков [Microsoft 365.](https://developer.microsoft.com/office/dev-program)</span><span class="sxs-lookup"><span data-stu-id="5b5d8-123">If you don't already have a Microsoft 365 account, you can get a free, 90-day renewable Microsoft 365 subscription that includes all Office apps by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="5b5d8-124">Установка редактора кода</span><span class="sxs-lookup"><span data-stu-id="5b5d8-124">Install a code editor</span></span>

<span data-ttu-id="5b5d8-125">Для создания веб-частей можно использовать любой редактор кода или интерфейс IDE, поддерживающий клиентскую разработку, например:</span><span class="sxs-lookup"><span data-stu-id="5b5d8-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="5b5d8-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="5b5d8-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- <span data-ttu-id="5b5d8-127">[Atom](https://atom.io);</span><span class="sxs-lookup"><span data-stu-id="5b5d8-127">[Atom](https://atom.io)</span></span>
- <span data-ttu-id="5b5d8-128">[Webstorm](https://www.jetbrains.com/webstorm).</span><span class="sxs-lookup"><span data-stu-id="5b5d8-128">[Webstorm](https://www.jetbrains.com/webstorm)</span></span>

## <a name="next-steps"></a><span data-ttu-id="5b5d8-129">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="5b5d8-129">Next steps</span></span>

<span data-ttu-id="5b5d8-130">Попробуйте создать собственную надстройку или воспользуйтесь Script Lab, чтобы попробовать встроенные примеры.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="5b5d8-131">Создание надстройки Office</span><span class="sxs-lookup"><span data-stu-id="5b5d8-131">Create an Office Add-in</span></span>

<span data-ttu-id="5b5d8-132">Вы можете быстро создать простую надстройку для Excel, OneNote, Outlook, PowerPoint, Project или Word с помощью [5-минутного краткого руководства по началу работы](../index.yml).</span><span class="sxs-lookup"><span data-stu-id="5b5d8-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.yml).</span></span> <span data-ttu-id="5b5d8-133">Если вы уже ознакомились с кратким руководством и хотите создать более сложную надстройку, воспользуйтесь [учебником](../index.yml).</span><span class="sxs-lookup"><span data-stu-id="5b5d8-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.yml).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="5b5d8-134">Изучение API с помощью Script Lab</span><span class="sxs-lookup"><span data-stu-id="5b5d8-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="5b5d8-135">Изучите библиотеку встроенных примеров в [Script Lab](explore-with-script-lab.md), чтобы ознакомиться с возможностями API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="5b5d8-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="5b5d8-136">См. также</span><span class="sxs-lookup"><span data-stu-id="5b5d8-136">See also</span></span>

- [<span data-ttu-id="5b5d8-137">Основные принципы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5b5d8-137">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="5b5d8-138">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5b5d8-138">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="5b5d8-139">Проектирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5b5d8-139">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="5b5d8-140">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5b5d8-140">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="5b5d8-141">Публикация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5b5d8-141">Publish Office Add-ins</span></span>](../publish/publish.md)
- [<span data-ttu-id="5b5d8-142">Сведения о программе для разработчиков Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="5b5d8-142">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
---
title: Создание первой надстройки Outlook
description: Узнайте, как создать простую надстройку для области задач Outlook, используя API JS для Office.
ms.date: 03/04/2020
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: bdef3ba43d8bb016f5d79ed3cafc25a6d89c345d
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185451"
---
# <a name="build-your-first-outlook-add-in"></a><span data-ttu-id="9d18b-103">Создание первой надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="9d18b-103">Build your first Outlook add-in</span></span>

<span data-ttu-id="9d18b-104">В этой статье вы ознакомитесь с процессом создания надстройки для области задач Outlook, отображающей минимум одно свойство выбранного сообщения.</span><span class="sxs-lookup"><span data-stu-id="9d18b-104">In this article, you'll walk through the process of building an Outlook task pane add-in that displays at least one property of a selected message.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="9d18b-105">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="9d18b-105">Create the add-in</span></span>

<span data-ttu-id="9d18b-106">Можно создать надстройку Office с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office) или Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="9d18b-106">You can create an Office Add-in by using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) or Visual Studio.</span></span> <span data-ttu-id="9d18b-107">Генератор Yeoman создает проект Node.js, которым можно управлять с помощью Visual Studio Code или любого другого редактора, а Visual Studio создает решение Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="9d18b-107">The Yeoman generator creates a Node.js project that can be managed with Visual Studio Code or any other editor, whereas Visual Studio creates a Visual Studio solution.</span></span>  <span data-ttu-id="9d18b-108">Выберите вкладку с нужным вариантом и следуйте инструкциям, чтобы создать надстройку и протестировать ее локально.</span><span class="sxs-lookup"><span data-stu-id="9d18b-108">Select the tab for the one you'd like to use and then follow the instructions to create your add-in and test it locally.</span></span>

# <a name="yeoman-generator"></a>[<span data-ttu-id="9d18b-109">Генератор Yeoman</span><span class="sxs-lookup"><span data-stu-id="9d18b-109">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="9d18b-110">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="9d18b-110">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

- <span data-ttu-id="9d18b-111">[Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="9d18b-111">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

- <span data-ttu-id="9d18b-112">Последняя версия [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="9d18b-112">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="9d18b-113">Даже если вы уже установили генератор Yeoman, рекомендуем обновить пакет до последней версии из npm.</span><span class="sxs-lookup"><span data-stu-id="9d18b-113">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="9d18b-114">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="9d18b-114">Create the add-in project</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="9d18b-115">**Выберите тип проекта** - `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="9d18b-115">**Choose a project type** - `Office Add-in Task Pane project`</span></span>

    - <span data-ttu-id="9d18b-116">**Выберите тип сценария** - `Javascript`</span><span class="sxs-lookup"><span data-stu-id="9d18b-116">**Choose a script type** - `Javascript`</span></span>

    - <span data-ttu-id="9d18b-117">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="9d18b-117">**What do you want to name your add-in?**</span></span> - `My Office Add-in`

    - <span data-ttu-id="9d18b-118">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="9d18b-118">**Which Office client application would you like to support?**</span></span> - `Outlook`

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-outlook.png)
    
    <span data-ttu-id="9d18b-120">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="9d18b-120">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. <span data-ttu-id="9d18b-121">Перейдите в корневую папку проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="9d18b-121">Navigate to the root folder of the web application project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="explore-the-project"></a><span data-ttu-id="9d18b-122">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="9d18b-122">Explore the project</span></span>

<span data-ttu-id="9d18b-123">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="9d18b-123">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="9d18b-124">Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-124">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="9d18b-125">Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="9d18b-125">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="9d18b-126">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="9d18b-126">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="9d18b-127">Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задачи и Outlook.</span><span class="sxs-lookup"><span data-stu-id="9d18b-127">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and Outlook.</span></span>

### <a name="update-the-code"></a><span data-ttu-id="9d18b-128">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="9d18b-128">Update the code</span></span>

1. <span data-ttu-id="9d18b-129">Откройте в редакторе кода файл **./src/taskpane/taskpane.html** и замените весь элемент `<main>` (внутри элемента `<body>`) приведенной ниже разметкой.</span><span class="sxs-lookup"><span data-stu-id="9d18b-129">In your code editor, open the file **./src/taskpane/taskpane.html** and replace the entire `<main>` element (within the `<body>` element) with the following markup.</span></span> <span data-ttu-id="9d18b-130">Эта новая разметка добавляет метку в том месте, где скрипт **./src/taskpane/taskpane.js** запишет данные.</span><span class="sxs-lookup"><span data-stu-id="9d18b-130">This new markup adds a label where the script in **./src/taskpane/taskpane.js** will write data.</span></span>

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. <span data-ttu-id="9d18b-131">Откройте файл **./src/taskpane/taskpane.js** в редакторе кода и добавьте следующий код в функцию `run`.</span><span class="sxs-lookup"><span data-stu-id="9d18b-131">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="9d18b-132">В этом коде используется API JavaScript для Office для получения ссылки на текущее сообщение и записи его свойства `subject` в область задач.</span><span class="sxs-lookup"><span data-stu-id="9d18b-132">This code uses the Office JavaScript API to get a reference to the current message and write its `subject` property value to the task pane.</span></span>

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a><span data-ttu-id="9d18b-133">Проверка</span><span class="sxs-lookup"><span data-stu-id="9d18b-133">Try it out</span></span>

> [!NOTE]
> <span data-ttu-id="9d18b-134">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-134">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="9d18b-135">Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="9d18b-135">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

1. <span data-ttu-id="9d18b-136">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="9d18b-136">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="9d18b-137">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="9d18b-137">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="9d18b-138">Чтобы загрузить неопубликованную надстройку в Outlook, следуйте инструкциями из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="9d18b-138">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="9d18b-139">В Outlook выберите или откройте сообщение.</span><span class="sxs-lookup"><span data-stu-id="9d18b-139">In Outlook, select or open a message.</span></span>

1. <span data-ttu-id="9d18b-140">Выберите вкладку **Главная** (или вкладку **Сообщения**, если вы открыли сообщение в новом окне), а затем нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-140">Choose the **Home** tab (or the **Message** tab if you opened the message in a new window), and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана с окном сообщения в Outlook, где выделена кнопка надстройки](../images/quick-start-button-1.png)

1. <span data-ttu-id="9d18b-142">Прокрутите область задачи в самый низ и перейдите по ссылке **Выполнить**, чтобы написать тему сообщения в области задач.</span><span class="sxs-lookup"><span data-stu-id="9d18b-142">Scroll to the bottom of the task pane and choose the **Run** link to write the message subject to the task pane.</span></span>

    ![Снимок экрана: область задач надстройки с выделенной ссылкой "Выполнить"](../images/quick-start-task-pane-2.png)

    ![Снимок экрана: область задач надстройки с темой сообщения](../images/quick-start-task-pane-3.png)

### <a name="next-steps"></a><span data-ttu-id="9d18b-145">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="9d18b-145">Next steps</span></span>

<span data-ttu-id="9d18b-146">Поздравляем! Вы успешно создали свою первую надстройку для области задач Outlook!</span><span class="sxs-lookup"><span data-stu-id="9d18b-146">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="9d18b-147">Теперь воспользуйтесь [руководством по надстройкам Outlook](../tutorials/outlook-tutorial.md), чтобы узнать больше о возможностях надстроек Outlook и создать более сложную надстройку.</span><span class="sxs-lookup"><span data-stu-id="9d18b-147">Next, learn more about the capabilities of an Outlook add-in and build a more complex add-in by following along with the [Outlook add-in tutorial](../tutorials/outlook-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="9d18b-148">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="9d18b-148">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="9d18b-149">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="9d18b-149">Prerequisites</span></span>

- <span data-ttu-id="9d18b-150">[Visual Studio 2019](https://www.visualstudio.com/vs/) с установленной рабочей нагрузкой **Разработка надстроек для Office и SharePoint**</span><span class="sxs-lookup"><span data-stu-id="9d18b-150">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="9d18b-151">Если вы уже установили Visual Studio 2019, [используйте установщик Visual Studio](/visualstudio/install/modify-visual-studio), чтобы убедиться, что также установлена рабочая нагрузка **Разработка надстроек для Office и SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="9d18b-151">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span>

- <span data-ttu-id="9d18b-152">Office 365</span><span class="sxs-lookup"><span data-stu-id="9d18b-152">Office 365</span></span>

    > [!NOTE]
    > <span data-ttu-id="9d18b-153">Если у вас нет подписки на Office 365, вы можете получить бесплатную подписку, зарегистрировавшись в [программе для разработчиков Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="9d18b-153">If you do not have an Office 365 subscription, you can get a free one by signing up for the [Office 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="9d18b-154">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="9d18b-154">Create the add-in project</span></span>

1. <span data-ttu-id="9d18b-155">В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="9d18b-155">On the Visual Studio menu bar, choose **File** > **New** > **Project**.</span></span>

1. <span data-ttu-id="9d18b-156">В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, затем выберите **Надстройки** > **Веб-надстройка Outlook**.</span><span class="sxs-lookup"><span data-stu-id="9d18b-156">In the list of project types under **Visual C#** or **Visual Basic**, expand **Office/SharePoint**, choose **Add-ins**, and then choose **Outlook Web Add-in** as the project type.</span></span>

1. <span data-ttu-id="9d18b-157">Укажите имя проекта и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="9d18b-157">Name the project, and then choose **OK**.</span></span>

1. <span data-ttu-id="9d18b-158">Visual Studio создаст решение, и два соответствующих проекта появятся в **обозревателе решений**.</span><span class="sxs-lookup"><span data-stu-id="9d18b-158">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="9d18b-159">Файл **MessageRead.html** откроется в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="9d18b-159">The **MessageRead.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="9d18b-160">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="9d18b-160">Explore the Visual Studio solution</span></span>

<span data-ttu-id="9d18b-161">После завершения работы мастера Visual Studio создает решение, которое содержит два проекта.</span><span class="sxs-lookup"><span data-stu-id="9d18b-161">When you've completed the wizard, Visual Studio creates a solution that contains two projects.</span></span>

|<span data-ttu-id="9d18b-162">**Проект**</span><span class="sxs-lookup"><span data-stu-id="9d18b-162">**Project**</span></span>|<span data-ttu-id="9d18b-163">**Описание**</span><span class="sxs-lookup"><span data-stu-id="9d18b-163">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9d18b-164">Проект надстройки</span><span class="sxs-lookup"><span data-stu-id="9d18b-164">Add-in project</span></span>|<span data-ttu-id="9d18b-165">Содержит только XML-файл манифеста со всеми параметрами надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-165">Contains only an XML manifest file, which contains all the settings that describe your add-in.</span></span> <span data-ttu-id="9d18b-166">Эти параметры помогают ведущему приложению Office определить условия активации и место отображения надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-166">These settings help the Office host determine when your add-in should be activated and where the add-in should appear.</span></span> <span data-ttu-id="9d18b-167">Visual Studio создает этот файл автоматически, чтобы вы могли сразу запускать проект и использовать надстройку.</span><span class="sxs-lookup"><span data-stu-id="9d18b-167">Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately.</span></span> <span data-ttu-id="9d18b-168">Вы можете изменить эти параметры в любой момент, отредактировав XML-файл.</span><span class="sxs-lookup"><span data-stu-id="9d18b-168">You can change these settings any time by modifying the XML file.</span></span>|
|<span data-ttu-id="9d18b-169">Проект веб-приложения</span><span class="sxs-lookup"><span data-stu-id="9d18b-169">Web application project</span></span>|<span data-ttu-id="9d18b-p109">Содержит страницы контента надстройки, включающие все файлы и ссылки на файлы, необходимые для разработки страниц HTML и JavaScript с поддержкой Office. При разработке надстройки Visual Studio размещает веб-приложение на локальном сервере IIS. Для публикации надстройки этот проект веб-приложения нужно развернуть на веб-сервере.</span><span class="sxs-lookup"><span data-stu-id="9d18b-p109">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish the add-in, you'll need to deploy this web application project to a web server.</span></span>|

### <a name="update-the-code"></a><span data-ttu-id="9d18b-173">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="9d18b-173">Update the code</span></span>

1. <span data-ttu-id="9d18b-174">Файл **MessageRead.html** содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-174">**MessageRead.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="9d18b-175">Замените элемент `<body>` в **MessageRead.html** приведенной ниже частью кода и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="9d18b-175">In **MessageRead.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```HTML
    <body class="ms-font-m ms-welcome">
        <div class="ms-Fabric content-main">
            <h1 class="ms-font-xxl">Message properties</h1>
            <table class="ms-Table ms-Table--selectable">
                <thead>
                    <tr>
                        <th>Property</th>
                        <th>Value</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>Id</strong></td>
                        <td class="prop-val"><code><label id="item-id"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Subject</strong></td>
                        <td class="prop-val"><code><label id="item-subject"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Message Id</strong></td>
                        <td class="prop-val"><code><label id="item-internetMessageId"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>From</strong></td>
                        <td class="prop-val"><code><label id="item-from"></label></code></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </body>
    ```

1. <span data-ttu-id="9d18b-176">Откройте файл **MessageRead.js** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="9d18b-176">Open the file **MessageRead.js** in the root of the web application project.</span></span> <span data-ttu-id="9d18b-177">Этот файл содержит скрипт надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-177">This file specifies the script for the add-in.</span></span> <span data-ttu-id="9d18b-178">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="9d18b-178">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.onReady(function () {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                loadItemProps(Office.context.mailbox.item);
            });
        });

        function loadItemProps(item) {
            // Write message property values to the task pane
            $('#item-id').text(item.itemId);
            $('#item-subject').text(item.subject);
            $('#item-internetMessageId').text(item.internetMessageId);
            $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        }
    })();
    ```

1. <span data-ttu-id="9d18b-179">Откройте файл **MessageRead.css** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="9d18b-179">Open the file **MessageRead.css** in the root of the web application project.</span></span> <span data-ttu-id="9d18b-180">Этот файл определяет специальные стили надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-180">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="9d18b-181">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="9d18b-181">Replace the entire contents with the following code and save the file.</span></span>

    ```CSS
    html,
    body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    td.prop-val {
        word-break: break-all;
    }

    .content-main {
        margin: 10px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="9d18b-182">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="9d18b-182">Update the manifest</span></span>

1. <span data-ttu-id="9d18b-p113">Откройте XML-файл манифеста в проекте надстройки. Этот файл определяет параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-p113">Open the XML manifest file in the Add-in project. This file defines the add-in's settings and capabilities.</span></span>

1. <span data-ttu-id="9d18b-p114">Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="9d18b-p114">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

1. <span data-ttu-id="9d18b-187">Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="9d18b-187">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="9d18b-188">Замените его на текст `My Office Add-in`.</span><span class="sxs-lookup"><span data-stu-id="9d18b-188">Replace it with `My Office Add-in`.</span></span>

1. <span data-ttu-id="9d18b-189">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="9d18b-189">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="9d18b-190">Замените его на текст `My First Outlook add-in`.</span><span class="sxs-lookup"><span data-stu-id="9d18b-190">Replace it with `My First Outlook add-in`.</span></span>

1. <span data-ttu-id="9d18b-191">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="9d18b-191">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>Northwind Traders</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="My First Outlook add-in"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="9d18b-192">Проверка</span><span class="sxs-lookup"><span data-stu-id="9d18b-192">Try it out</span></span>

1. <span data-ttu-id="9d18b-193">Протестируйте созданную в Visual Studio надстройку Outlook, нажав F5 или кнопку **Запуск**.</span><span class="sxs-lookup"><span data-stu-id="9d18b-193">Using Visual Studio, test the newly created Outlook add-in by pressing F5 or choosing the **Start** button.</span></span> <span data-ttu-id="9d18b-194">Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="9d18b-194">The add-in will be hosted locally on IIS.</span></span>

1. <span data-ttu-id="9d18b-195">В диалоговом окне **Подключение к учетной записи электронной почты Exchange** введите адрес электронной почты и пароль для вашей [учетной записи Майкрософт](https://account.microsoft.com/account) и нажмите кнопку **Подключить**.</span><span class="sxs-lookup"><span data-stu-id="9d18b-195">In the **Connect to Exchange email account** dialog box, enter the email address and password for your [Microsoft account](https://account.microsoft.com/account) and then choose **Connect**.</span></span> <span data-ttu-id="9d18b-196">Когда в браузере откроется страница входа в Outlook.com, войдите в свою учетную запись электронной почты с теми же учетными данными, которые были введены ранее.</span><span class="sxs-lookup"><span data-stu-id="9d18b-196">When the Outlook.com login page opens in a browser, sign in to your email account with the same credentials as you entered previously.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9d18b-197">Если диалоговое окно **Подключение к учетной записи электронной почты Exchange** повторно предлагает выполнить вход, для учетных записей в вашем клиенте Office 365, возможно, отключена обычная проверка подлинности.</span><span class="sxs-lookup"><span data-stu-id="9d18b-197">If the **Connect to Exchange email account** dialog box repeatedly prompts you to sign in, Basic Auth may be disabled for accounts on your Office 365 tenant.</span></span> <span data-ttu-id="9d18b-198">Чтобы протестировать эту надстройку, вместо этого выполните вход с помощью [учетной записи Майкрософт](https://account.microsoft.com/account).</span><span class="sxs-lookup"><span data-stu-id="9d18b-198">To test this add-in, sign in using a [Microsoft account](https://account.microsoft.com/account) instead.</span></span>

1. <span data-ttu-id="9d18b-199">В Outlook в Интернете выберите или откройте сообщение.</span><span class="sxs-lookup"><span data-stu-id="9d18b-199">In Outlook on the web, select or open a message.</span></span>

1. <span data-ttu-id="9d18b-200">В сообщении найдите многоточие, чтобы перейти в меню переполнения, содержащее кнопку надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-200">Within the message, locate the ellipsis for the overflow menu containing the add-in's button.</span></span>

    ![Снимок экрана: окно сообщения в Outlook в Интернете, в котором выделено многоточие](../images/quick-start-button-owa-1.png)

1. <span data-ttu-id="9d18b-202">Найдите кнопку надстройки в меню переполнения.</span><span class="sxs-lookup"><span data-stu-id="9d18b-202">Within the overflow menu, locate the add-in's button.</span></span>

    ![Снимок экрана с окном сообщения в Outlook в Интернете, где выделена кнопка надстройки](../images/quick-start-button-owa-2.png)

1. <span data-ttu-id="9d18b-204">Нажмите кнопку, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d18b-204">Click the button to open the add-in's task pane.</span></span>

    ![Снимок экрана: область задач надстройки в Outlook в Интернете со свойствами сообщения](../images/quick-start-task-pane-owa-1.png)

    > [!NOTE]
    > <span data-ttu-id="9d18b-206">Если область задач не загружается, проверьте ее, открыв в браузере на том же компьютере.</span><span class="sxs-lookup"><span data-stu-id="9d18b-206">If the task pane doesn't load, try to verify by opening it in a browser on the same machine.</span></span>

### <a name="next-steps"></a><span data-ttu-id="9d18b-207">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="9d18b-207">Next steps</span></span>

<span data-ttu-id="9d18b-208">Поздравляем! Вы успешно создали свою первую надстройку для области задач Outlook!</span><span class="sxs-lookup"><span data-stu-id="9d18b-208">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="9d18b-209">Теперь изучите дополнительные сведения о [разработке надстроек Office с помощью Visual Studio](../develop/develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="9d18b-209">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---

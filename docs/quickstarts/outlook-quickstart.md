---
title: Создание первой надстройки Outlook
description: Узнайте, как создать простую надстройку для области задач Outlook, используя API JS для Office.
ms.date: 06/10/2021
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 75a2b4d5a3ec995971b6cca1e47c77572d274555
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076604"
---
# <a name="build-your-first-outlook-add-in"></a><span data-ttu-id="243fe-103">Создание первой надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="243fe-103">Build your first Outlook add-in</span></span>

<span data-ttu-id="243fe-104">В этой статье вы ознакомитесь с процессом создания надстройки для области задач Outlook, отображающей минимум одно свойство выбранного сообщения.</span><span class="sxs-lookup"><span data-stu-id="243fe-104">In this article, you'll walk through the process of building an Outlook task pane add-in that displays at least one property of a selected message.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="243fe-105">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="243fe-105">Create the add-in</span></span>

<span data-ttu-id="243fe-106">Можно создать надстройку Office с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office) или Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="243fe-106">You can create an Office Add-in by using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) or Visual Studio.</span></span> <span data-ttu-id="243fe-107">Генератор Yeoman создает проект Node.js, которым можно управлять с помощью Visual Studio Code или любого другого редактора, а Visual Studio создает решение Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="243fe-107">The Yeoman generator creates a Node.js project that can be managed with Visual Studio Code or any other editor, whereas Visual Studio creates a Visual Studio solution.</span></span> <span data-ttu-id="243fe-108">Выберите вкладку с нужным вариантом и следуйте инструкциям, чтобы создать надстройку и протестировать ее локально.</span><span class="sxs-lookup"><span data-stu-id="243fe-108">Select the tab for the one you'd like to use and then follow the instructions to create your add-in and test it locally.</span></span>

# <a name="yeoman-generator"></a>[<span data-ttu-id="243fe-109">Генератор Yeoman</span><span class="sxs-lookup"><span data-stu-id="243fe-109">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="243fe-110">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="243fe-110">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

- <span data-ttu-id="243fe-111">[Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="243fe-111">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

- <span data-ttu-id="243fe-112">Последняя версия [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="243fe-112">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="243fe-113">Даже если вы уже установили генератор Yeoman, рекомендуем обновить пакет до последней версии из npm.</span><span class="sxs-lookup"><span data-stu-id="243fe-113">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="243fe-114">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="243fe-114">Create the add-in project</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="243fe-115">**Выберите тип проекта** - `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="243fe-115">**Choose a project type** - `Office Add-in Task Pane project`</span></span>

    - <span data-ttu-id="243fe-116">**Выберите тип сценария** - `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="243fe-116">**Choose a script type** - `JavaScript`</span></span>

    - <span data-ttu-id="243fe-117">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="243fe-117">**What do you want to name your add-in?**</span></span> - `My Office Add-in`

    - <span data-ttu-id="243fe-118">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="243fe-118">**Which Office client application would you like to support?**</span></span> - `Outlook`

    ![Снимок экрана: запросы и ответы для генератора Yeoman в интерфейсе командной строки.](../images/yo-office-outlook-1.png)

    <span data-ttu-id="243fe-120">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="243fe-120">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. <span data-ttu-id="243fe-121">Перейдите в корневую папку проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="243fe-121">Navigate to the root folder of the web application project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="explore-the-project"></a><span data-ttu-id="243fe-122">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="243fe-122">Explore the project</span></span>

<span data-ttu-id="243fe-123">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="243fe-123">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span>

- <span data-ttu-id="243fe-124">Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="243fe-124">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="243fe-125">Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="243fe-125">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="243fe-126">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="243fe-126">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="243fe-127">Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задачи и Outlook.</span><span class="sxs-lookup"><span data-stu-id="243fe-127">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and Outlook.</span></span>

### <a name="update-the-code"></a><span data-ttu-id="243fe-128">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="243fe-128">Update the code</span></span>

1. <span data-ttu-id="243fe-129">Откройте в редакторе кода файл **./src/taskpane/taskpane.html** и замените весь элемент `<main>` (внутри элемента `<body>`) приведенной ниже разметкой.</span><span class="sxs-lookup"><span data-stu-id="243fe-129">In your code editor, open the file **./src/taskpane/taskpane.html** and replace the entire `<main>` element (within the `<body>` element) with the following markup.</span></span> <span data-ttu-id="243fe-130">Эта новая разметка добавляет метку в том месте, где скрипт **./src/taskpane/taskpane.js** запишет данные.</span><span class="sxs-lookup"><span data-stu-id="243fe-130">This new markup adds a label where the script in **./src/taskpane/taskpane.js** will write data.</span></span>

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. <span data-ttu-id="243fe-131">Откройте файл **./src/taskpane/taskpane.js** в редакторе кода и добавьте следующий код в функцию `run`.</span><span class="sxs-lookup"><span data-stu-id="243fe-131">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="243fe-132">В этом коде используется API JavaScript для Office для получения ссылки на текущее сообщение и записи его свойства `subject` в область задач.</span><span class="sxs-lookup"><span data-stu-id="243fe-132">This code uses the Office JavaScript API to get a reference to the current message and write its `subject` property value to the task pane.</span></span>

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a><span data-ttu-id="243fe-133">Проверка</span><span class="sxs-lookup"><span data-stu-id="243fe-133">Try it out</span></span>

> [!NOTE]
> <span data-ttu-id="243fe-134">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="243fe-134">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="243fe-135">Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="243fe-135">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="243fe-136">Кроме того, вам может потребоваться запустить командную строку или терминал с правами администратора, чтобы внести изменения.</span><span class="sxs-lookup"><span data-stu-id="243fe-136">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

1. <span data-ttu-id="243fe-137">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="243fe-137">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="243fe-138">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен) и будет загружена ваша неопубликованная надстройка.</span><span class="sxs-lookup"><span data-stu-id="243fe-138">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="243fe-139">Чтобы загрузить неопубликованную надстройку в Outlook, следуйте инструкциями из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="243fe-139">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="243fe-140">В Outlook просмотрите сообщение в [области чтения](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) или откройте сообщение в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="243fe-140">In Outlook, view a message in the [Reading Pane](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0), or open the message in its own window.</span></span>

1. <span data-ttu-id="243fe-141">Выберите вкладку **Главная** (или вкладку **Сообщения**, если вы открыли сообщение в новом окне), а затем нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="243fe-141">Choose the **Home** tab (or the **Message** tab if you opened the message in a new window), and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана с окном сообщения в Outlook с выделенной кнопкой ленты надстройки.](../images/quick-start-button-1.png)

    > [!NOTE]
    > <span data-ttu-id="243fe-143">Если сообщение об ошибке "Не удается открыть эту надстройку с localhost" появляется в области задач, выполните действия, описанные в [статье по устранению неполадок](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).</span><span class="sxs-lookup"><span data-stu-id="243fe-143">If you receive the error "We can't open this add-in from localhost" in the task pane, follow the steps outlined in the [troubleshooting article](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).</span></span>

1. <span data-ttu-id="243fe-144">Прокрутите область задачи в самый низ и перейдите по ссылке **Выполнить**, чтобы написать тему сообщения в области задач.</span><span class="sxs-lookup"><span data-stu-id="243fe-144">Scroll to the bottom of the task pane and choose the **Run** link to write the message subject to the task pane.</span></span>

    ![Снимок экрана: область задач надстройки с выделенной ссылкой "Выполнить".](../images/quick-start-task-pane-2.png)

    ![Снимок экрана: область задач надстройки с темой сообщения.](../images/quick-start-task-pane-3.png)

### <a name="next-steps"></a><span data-ttu-id="243fe-147">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="243fe-147">Next steps</span></span>

<span data-ttu-id="243fe-148">Поздравляем! Вы успешно создали свою первую надстройку для области задач Outlook!</span><span class="sxs-lookup"><span data-stu-id="243fe-148">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="243fe-149">Теперь воспользуйтесь [руководством по надстройкам Outlook](../tutorials/outlook-tutorial.md), чтобы узнать больше о возможностях надстроек Outlook и создать более сложную надстройку.</span><span class="sxs-lookup"><span data-stu-id="243fe-149">Next, learn more about the capabilities of an Outlook add-in and build a more complex add-in by following along with the [Outlook add-in tutorial](../tutorials/outlook-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="243fe-150">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="243fe-150">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="243fe-151">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="243fe-151">Prerequisites</span></span>

- <span data-ttu-id="243fe-152">[Visual Studio 2019](https://www.visualstudio.com/vs/) с установленной рабочей нагрузкой **Разработка надстроек для Office и SharePoint**</span><span class="sxs-lookup"><span data-stu-id="243fe-152">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="243fe-153">Если вы уже установили Visual Studio 2019, [используйте установщик Visual Studio](/visualstudio/install/modify-visual-studio), чтобы убедиться, что также установлена рабочая нагрузка **Разработка надстроек для Office и SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="243fe-153">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span>

- <span data-ttu-id="243fe-154">Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="243fe-154">Microsoft 365</span></span>

    > [!NOTE]
    > <span data-ttu-id="243fe-155">Если у вас нет подписки на Microsoft 365, вы можете получить бесплатную подписку, зарегистрировавшись в [программе для разработчиков Microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="243fe-155">If you do not have a Microsoft 365 subscription, you can get a free one by signing up for the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="243fe-156">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="243fe-156">Create the add-in project</span></span>

1. <span data-ttu-id="243fe-157">В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="243fe-157">On the Visual Studio menu bar, choose **File** > **New** > **Project**.</span></span>

1. <span data-ttu-id="243fe-158">В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, затем выберите **Надстройки** > **Веб-надстройка Outlook**.</span><span class="sxs-lookup"><span data-stu-id="243fe-158">In the list of project types under **Visual C#** or **Visual Basic**, expand **Office/SharePoint**, choose **Add-ins**, and then choose **Outlook Web Add-in** as the project type.</span></span>

1. <span data-ttu-id="243fe-159">Укажите имя проекта и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="243fe-159">Name the project, and then choose **OK**.</span></span>

1. <span data-ttu-id="243fe-160">Visual Studio создаст решение, и два соответствующих проекта появятся в **обозревателе решений**.</span><span class="sxs-lookup"><span data-stu-id="243fe-160">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="243fe-161">Файл **MessageRead.html** откроется в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="243fe-161">The **MessageRead.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="243fe-162">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="243fe-162">Explore the Visual Studio solution</span></span>

<span data-ttu-id="243fe-163">После завершения работы мастера Visual Studio создает решение, которое содержит два проекта.</span><span class="sxs-lookup"><span data-stu-id="243fe-163">When you've completed the wizard, Visual Studio creates a solution that contains two projects.</span></span>

|<span data-ttu-id="243fe-164">**Проект**</span><span class="sxs-lookup"><span data-stu-id="243fe-164">**Project**</span></span>|<span data-ttu-id="243fe-165">**Описание**</span><span class="sxs-lookup"><span data-stu-id="243fe-165">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="243fe-166">Проект надстройки</span><span class="sxs-lookup"><span data-stu-id="243fe-166">Add-in project</span></span>|<span data-ttu-id="243fe-p108">Содержит только XML-файл манифеста со всеми параметрами надстройки. Эти параметры помогают приложению Office определить условия активации и место отображения надстройки. Visual Studio создает этот файл автоматически, чтобы вы могли сразу запускать проект и использовать надстройку. Вы можете изменить эти параметры в любой момент, отредактировав XML-файл.</span><span class="sxs-lookup"><span data-stu-id="243fe-p108">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office application determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You can change these settings any time by modifying the XML file.</span></span>|
|<span data-ttu-id="243fe-171">Проект веб-приложения</span><span class="sxs-lookup"><span data-stu-id="243fe-171">Web application project</span></span>|<span data-ttu-id="243fe-p109">Содержит страницы контента надстройки, включающие все файлы и ссылки на файлы, необходимые для разработки страниц HTML и JavaScript с поддержкой Office. При разработке надстройки Visual Studio размещает веб-приложение на локальном сервере IIS. Для публикации надстройки этот проект веб-приложения нужно развернуть на веб-сервере.</span><span class="sxs-lookup"><span data-stu-id="243fe-p109">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish the add-in, you'll need to deploy this web application project to a web server.</span></span>|

### <a name="update-the-code"></a><span data-ttu-id="243fe-175">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="243fe-175">Update the code</span></span>

1. <span data-ttu-id="243fe-176">Файл **MessageRead.html** содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="243fe-176">**MessageRead.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="243fe-177">Замените элемент `<body>` в **MessageRead.html** приведенной ниже частью кода и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="243fe-177">In **MessageRead.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

1. <span data-ttu-id="243fe-178">Откройте файл **MessageRead.js** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="243fe-178">Open the file **MessageRead.js** in the root of the web application project.</span></span> <span data-ttu-id="243fe-179">Этот файл содержит скрипт надстройки.</span><span class="sxs-lookup"><span data-stu-id="243fe-179">This file specifies the script for the add-in.</span></span> <span data-ttu-id="243fe-180">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="243fe-180">Replace the entire contents with the following code and save the file.</span></span>

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

1. <span data-ttu-id="243fe-181">Откройте файл **MessageRead.css** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="243fe-181">Open the file **MessageRead.css** in the root of the web application project.</span></span> <span data-ttu-id="243fe-182">Этот файл определяет специальные стили надстройки.</span><span class="sxs-lookup"><span data-stu-id="243fe-182">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="243fe-183">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="243fe-183">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="243fe-184">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="243fe-184">Update the manifest</span></span>

1. <span data-ttu-id="243fe-p113">Откройте XML-файл манифеста в проекте надстройки. Этот файл определяет параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="243fe-p113">Open the XML manifest file in the Add-in project. This file defines the add-in's settings and capabilities.</span></span>

1. <span data-ttu-id="243fe-p114">Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="243fe-p114">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

1. <span data-ttu-id="243fe-189">Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="243fe-189">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="243fe-190">Замените его на текст `My Office Add-in`.</span><span class="sxs-lookup"><span data-stu-id="243fe-190">Replace it with `My Office Add-in`.</span></span>

1. <span data-ttu-id="243fe-191">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="243fe-191">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="243fe-192">Замените его на текст `My First Outlook add-in`.</span><span class="sxs-lookup"><span data-stu-id="243fe-192">Replace it with `My First Outlook add-in`.</span></span>

1. <span data-ttu-id="243fe-193">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="243fe-193">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="My First Outlook add-in"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="243fe-194">Проверка</span><span class="sxs-lookup"><span data-stu-id="243fe-194">Try it out</span></span>

1. <span data-ttu-id="243fe-195">Протестируйте созданную в Visual Studio надстройку Outlook, нажав F5 или кнопку **Запуск**.</span><span class="sxs-lookup"><span data-stu-id="243fe-195">Using Visual Studio, test the newly created Outlook add-in by pressing F5 or choosing the **Start** button.</span></span> <span data-ttu-id="243fe-196">Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="243fe-196">The add-in will be hosted locally on IIS.</span></span>

1. <span data-ttu-id="243fe-197">В диалоговом окне **Подключение к учетной записи электронной почты Exchange** введите адрес электронной почты и пароль для вашей [учетной записи Майкрософт](https://account.microsoft.com/account) и нажмите кнопку **Подключить**.</span><span class="sxs-lookup"><span data-stu-id="243fe-197">In the **Connect to Exchange email account** dialog box, enter the email address and password for your [Microsoft account](https://account.microsoft.com/account) and then choose **Connect**.</span></span> <span data-ttu-id="243fe-198">Когда в браузере откроется страница входа в Outlook.com, войдите в свою учетную запись электронной почты с теми же учетными данными, которые были введены ранее.</span><span class="sxs-lookup"><span data-stu-id="243fe-198">When the Outlook.com login page opens in a browser, sign in to your email account with the same credentials as you entered previously.</span></span>

    > [!NOTE]
    > <span data-ttu-id="243fe-199">Если диалоговое окно **Подключение к учетной записи электронной почты Exchange** повторно предлагает выполнить вход или вы получили сообщение с ошибкой авторизации, для учетных записей в вашем клиенте Microsoft 365, возможно, отключена обычная проверка подлинности.</span><span class="sxs-lookup"><span data-stu-id="243fe-199">If the **Connect to Exchange email account** dialog box repeatedly prompts you to sign in or you receive an error that you are unauthorized, Basic Auth may be disabled for accounts on your Microsoft 365 tenant.</span></span> <span data-ttu-id="243fe-200">Чтобы проверить надстройку, попробуйте войти снова, задав для свойства **Использование многофакторной проверки подлинности** значение True в диалогом окне "Свойства проекта веб-надстройки", или войдите с помощью [учетной записи Майкрософт](https://account.microsoft.com/account).</span><span class="sxs-lookup"><span data-stu-id="243fe-200">To test this add-in, try signing in again after setting the **Use multi-factor auth** property to True in the Web Add-in project properties dialog, or sign in using a [Microsoft account](https://account.microsoft.com/account) instead.</span></span>

1. <span data-ttu-id="243fe-201">В Outlook в Интернете выберите или откройте сообщение.</span><span class="sxs-lookup"><span data-stu-id="243fe-201">In Outlook on the web, select or open a message.</span></span>

1. <span data-ttu-id="243fe-202">В сообщении найдите многоточие, чтобы перейти в меню переполнения, содержащее кнопку надстройки.</span><span class="sxs-lookup"><span data-stu-id="243fe-202">Within the message, locate the ellipsis for the overflow menu containing the add-in's button.</span></span>

    ![Снимок экрана: окно сообщения в Outlook в Интернете, в котором выделено многоточие.](../images/quick-start-button-owa-1.png)

1. <span data-ttu-id="243fe-204">Найдите кнопку надстройки в меню переполнения.</span><span class="sxs-lookup"><span data-stu-id="243fe-204">Within the overflow menu, locate the add-in's button.</span></span>

    ![Снимок экрана с окном сообщения в Outlook в Интернете, где выделена кнопка надстройки.](../images/quick-start-button-owa-2.png)

1. <span data-ttu-id="243fe-206">Нажмите кнопку, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="243fe-206">Click the button to open the add-in's task pane.</span></span>

    ![Снимок экрана: область задач надстройки в Outlook в Интернете со свойствами сообщения.](../images/quick-start-task-pane-owa-1.png)

    > [!NOTE]
    > <span data-ttu-id="243fe-208">Если область задач не загружается, проверьте ее, открыв в браузере на том же компьютере.</span><span class="sxs-lookup"><span data-stu-id="243fe-208">If the task pane doesn't load, try to verify by opening it in a browser on the same machine.</span></span>

### <a name="next-steps"></a><span data-ttu-id="243fe-209">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="243fe-209">Next steps</span></span>

<span data-ttu-id="243fe-210">Поздравляем! Вы успешно создали свою первую надстройку для области задач Outlook!</span><span class="sxs-lookup"><span data-stu-id="243fe-210">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="243fe-211">Теперь изучите дополнительные сведения о [разработке надстроек Office с помощью Visual Studio](../develop/develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="243fe-211">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---

# <a name="build-your-first-project-add-in"></a><span data-ttu-id="74906-101">Создание первой надстройки Project</span><span class="sxs-lookup"><span data-stu-id="74906-101">Build your first Project add-in</span></span>

<span data-ttu-id="74906-102">В этой статье мы разберем, как создать надстройку Project, используя jQuery и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="74906-102">In this article, you'll walk through the process of building a Project add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="74906-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="74906-103">Prerequisites</span></span>

- [<span data-ttu-id="74906-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="74906-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="74906-105">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="74906-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a><span data-ttu-id="74906-106">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="74906-106">Create the add-in</span></span>

1. <span data-ttu-id="74906-107">Создайте на локальном диске папку и назовите ее `my-project-addin`.</span><span class="sxs-lookup"><span data-stu-id="74906-107">Create a folder on your local drive and name it `my-project-addin`.</span></span> <span data-ttu-id="74906-108">В ней вы будете создавать файлы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="74906-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="74906-109">Перейдите к новой папке.</span><span class="sxs-lookup"><span data-stu-id="74906-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-project-addin
    ```

3. <span data-ttu-id="74906-110">С помощью генератора Yeoman создайте проект надстройки Project.</span><span class="sxs-lookup"><span data-stu-id="74906-110">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="74906-111">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="74906-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="74906-112">**Выберите тип проекта:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="74906-112">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="74906-113">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="74906-113">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="74906-114">**Как вы хотите назвать надстройку?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="74906-114">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="74906-115">**Какое клиентское приложение Office должно поддерживаться?** `Project`</span><span class="sxs-lookup"><span data-stu-id="74906-115">**Which Office client application would you like to support?:** `Project`</span></span>

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-project-jquery.png)
    
    <span data-ttu-id="74906-117">После завершения работы мастера генератор создаст проект и установит поддерживающие компоненты узла.</span><span class="sxs-lookup"><span data-stu-id="74906-117">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="74906-118">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="74906-118">Update the code</span></span>

1. <span data-ttu-id="74906-119">В редакторе кода откройте файл **index.html** из корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="74906-119">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="74906-120">Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="74906-120">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="74906-121">Замените элемент `<header>` внутри элемента `<body>` на приведенную ниже разметку.</span><span class="sxs-lookup"><span data-stu-id="74906-121">Replace the `<header>` element inside the `<body>` element with the following markup.</span></span>

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

3. <span data-ttu-id="74906-122">Замените элемент `<main>` внутри элемента `<body>` приведенной ниже разметкой и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="74906-122">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
            <h3>Try it out</h3>
            <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
            <br/><br/>
            <button class="ms-Button" id="get-task">Get Task data</button>
            <br/>
            <h4>Results:</h4>
            <textarea id="result" rows="6" cols="25"></textarea>
        </div>
    </div>
    ```

4. <span data-ttu-id="74906-123">Откройте файл **app.js**, чтобы указать скрипт для надстройки.</span><span class="sxs-lookup"><span data-stu-id="74906-123">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="74906-124">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="74906-124">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        var taskGuid;

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        };

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. <span data-ttu-id="74906-125">Откройте файл **app.css** в корневой папке проекта, чтобы указать специальные стили для надстройки.</span><span class="sxs-lookup"><span data-stu-id="74906-125">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="74906-126">Замените все его содержимое следующим кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="74906-126">Replace the entire contents with the following and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="74906-127">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="74906-127">Update the manifest</span></span>

1. <span data-ttu-id="74906-128">Откройте файл **my-office-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="74906-128">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="74906-129">Элемент `ProviderName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="74906-129">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="74906-130">Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="74906-130">Replace it with your name.</span></span>

3. <span data-ttu-id="74906-131">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="74906-131">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="74906-132">Замените его строкой **Надстройка области задач для Project**.</span><span class="sxs-lookup"><span data-stu-id="74906-132">Replace it with **A task pane add-in for Project**.</span></span>

4. <span data-ttu-id="74906-133">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="74906-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="74906-134">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="74906-134">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="74906-135">Проверка</span><span class="sxs-lookup"><span data-stu-id="74906-135">Try it out</span></span>

1. <span data-ttu-id="74906-136">Создайте простой проект Project, содержащий по крайней мере одну задачу.</span><span class="sxs-lookup"><span data-stu-id="74906-136">In Project, create a simple project that has at least one task.</span></span>

2. <span data-ttu-id="74906-137">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Project.</span><span class="sxs-lookup"><span data-stu-id="74906-137">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Project.</span></span>

    - <span data-ttu-id="74906-138">Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="74906-138">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="74906-139">Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="74906-139">Project Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="74906-140">iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="74906-140">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

3. <span data-ttu-id="74906-141">Выберите задачу в Project.</span><span class="sxs-lookup"><span data-stu-id="74906-141">In Project, select a task.</span></span>

    ![Снимок экрана: план проекта в Project с одной выбранной задачей](../images/project_quickstart_addin_1.png)

4. <span data-ttu-id="74906-143">В области задач нажмите кнопку **Get Task GUID**, чтобы записать GUID задачи в поле **Results**.</span><span class="sxs-lookup"><span data-stu-id="74906-143">In the task pane, choose the **Get Task GUID** button to write the task GUID to the **Results** textbox.</span></span>

    ![Снимок экрана: план проекта в Project с одной выбранной задачей и GUID в текстовом поле области задач](../images/project_quickstart_addin_2.png)

5. <span data-ttu-id="74906-145">В области задач нажмите кнопку **Get Task data**, чтобы записать несколько свойств выбранной задачи в поле **Results**.</span><span class="sxs-lookup"><span data-stu-id="74906-145">In the task pane, choose the **Get Task data** button to write several properties of the selected task to the **Results** textbox.</span></span>

    ![Снимок экрана: план проекта в Project с одной выбранной задачей и несколькими свойствами в текстовом поле области задач](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a><span data-ttu-id="74906-147">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="74906-147">Next steps</span></span>

<span data-ttu-id="74906-148">Поздравляем, вы успешно создали надстройку Project!</span><span class="sxs-lookup"><span data-stu-id="74906-148">Congratulations, you've successfully created a Project add-in!</span></span> <span data-ttu-id="74906-149">Следующим шагом узнайте больше о возможностях надстроек Project и изучите распространенные сценарии.</span><span class="sxs-lookup"><span data-stu-id="74906-149">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="74906-150">Надстройки Project</span><span class="sxs-lookup"><span data-stu-id="74906-150">Project add-ins</span></span>](../project/project-add-ins.md)

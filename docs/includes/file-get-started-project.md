# <a name="build-your-first-project-add-in"></a><span data-ttu-id="57b09-101">Создание вашей первой надстройки Project</span><span class="sxs-lookup"><span data-stu-id="57b09-101">Build your first Project add-in</span></span>

<span data-ttu-id="57b09-102">В этой статье рассматривается процесс создания надстройки Project с прмиенением jQuery и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="57b09-102">In this article, you'll walk through the process of building a Project add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="57b09-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="57b09-103">Prerequisites</span></span>

- [<span data-ttu-id="57b09-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="57b09-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="57b09-105">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="57b09-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a><span data-ttu-id="57b09-106">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="57b09-106">Create the add-in</span></span>

1. <span data-ttu-id="57b09-p101">Создайте на локальном диске папку и назовите ее `my-project-addin`. В ней будут создаваться файлы для новой надстройки.</span><span class="sxs-lookup"><span data-stu-id="57b09-p101">Create a folder on your local drive and name it `my-project-addin`. This is where you'll create the files for your add-in.</span></span>

    ```bash
    mkdir my-project-addin
    ```

2. <span data-ttu-id="57b09-109">Перейдите к новой папке.</span><span class="sxs-lookup"><span data-stu-id="57b09-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-project-addin
    ```

3. <span data-ttu-id="57b09-p102">Используйте генератор Yeoman для создания проекта надстройки Project. Запустите указанную ниже команду, после чего ответьте на предлагаемые вопросы следующим образом:</span><span class="sxs-lookup"><span data-stu-id="57b09-p102">Use the Yeoman generator to create a Project add-in project. Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="57b09-112">**Выберите тип проекта:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="57b09-112">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="57b09-113">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="57b09-113">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="57b09-114">**Как вы хотите назвать надстройку?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="57b09-114">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="57b09-115">**Какое клиентское приложение Office должно поддерживаться?** `Project`</span><span class="sxs-lookup"><span data-stu-id="57b09-115">**Which Office client application would you like to support?:** `Project`</span></span>

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-project-jquery.png)
    
    <span data-ttu-id="57b09-117">После завершения работы мастера генератор создаст проект и установит поддерживающие компоненты узла.</span><span class="sxs-lookup"><span data-stu-id="57b09-117">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
4. <span data-ttu-id="57b09-118">Перейдите в корневую папку проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="57b09-118">Navigate to the root folder of the web application project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="57b09-119">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="57b09-119">Update the code</span></span>

1. <span data-ttu-id="57b09-p103">В редакторе кода откройте файл **index.html**, имеющийся в корневой папке проекта. Этот файл содержит HTML-содержимое, которое будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="57b09-p103">In your code editor, open **index.html** in the root of the project. This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="57b09-122">Замените элемент `<body>` следующей разметкой.</span><span class="sxs-lookup"><span data-stu-id="57b09-122">Replace the `<body>` element inside the  element with the following markup.</span></span>

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
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
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="57b09-p104">Откройте файл **index/src.js**, чтобы указать сценарий для надстройки. Замените все его содержимое следующим кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="57b09-p104">Open the file **src/index.js** to specify the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

4. <span data-ttu-id="57b09-p105">Откройте файл **app.css** в корневой папке проекта, чтобы указать настраиваемые стили для надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="57b09-p105">Open the file **app.css** in the root of the project to specify the custom styles for the add-in. Replace the entire contents with the following and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="57b09-127">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="57b09-127">Update the manifest</span></span>

1. <span data-ttu-id="57b09-128">Откройте файл **manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="57b09-128">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="57b09-p106">Элемент `ProviderName` содержит значение заполнителя. Замените его своим именем.</span><span class="sxs-lookup"><span data-stu-id="57b09-p106">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="57b09-p107">Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Word**.</span><span class="sxs-lookup"><span data-stu-id="57b09-p107">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Project**.</span></span>

4. <span data-ttu-id="57b09-133">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="57b09-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="57b09-134">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="57b09-134">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="57b09-135">Проверьте надстройку</span><span class="sxs-lookup"><span data-stu-id="57b09-135">Try it out</span></span>

1. <span data-ttu-id="57b09-136">В Project создайте простой проект, содержащий хотя бы одну задачу.</span><span class="sxs-lookup"><span data-stu-id="57b09-136">In Project, create a simple project that has at least one task.</span></span>

2. <span data-ttu-id="57b09-137">Следуя указаниям, касающимся платформы, которая используется для запуска надстройки, загрузите неопубликованную надстройку в Project.</span><span class="sxs-lookup"><span data-stu-id="57b09-137">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Project.</span></span>

    - <span data-ttu-id="57b09-138">Windows: [загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="57b09-138">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="57b09-139">Project Online: [Загрузка неопубликованных надстроек Office в Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="57b09-139">Project Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="57b09-140">iPad и Mac: [Загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="57b09-140">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

3. <span data-ttu-id="57b09-141">Выберите задачу в Project.</span><span class="sxs-lookup"><span data-stu-id="57b09-141">In Project, select a task.</span></span>

    ![Снимок экрана, отображающего план проекта в Project с одной выбранной задачей](../images/project_quickstart_addin_1.png)

4. <span data-ttu-id="57b09-143">В области задач нажмите на кнопку **Получить GUID задачи**, чтобы записать GUID задачи в поле **Результаты**.</span><span class="sxs-lookup"><span data-stu-id="57b09-143">In the task pane, choose the **Get Task GUID** button to write the task GUID to the **Results** textbox.</span></span>

    ![Снимок экрана, отображающего план проекта в Project с одной выбранной задачей и указанном в соответствующем поле области задач GUID задачи](../images/project_quickstart_addin_2.png)

5. <span data-ttu-id="57b09-145">В области задач нажмите на кнопку **Получить данные задачи**, чтобы записать несколько свойств выбранной задачи в поле **Результаты**.</span><span class="sxs-lookup"><span data-stu-id="57b09-145">In the task pane, choose the **Get Task data** button to write several properties of the selected task to the **Results** textbox.</span></span>

    ![Снимок экрана, отображающего план проекта в Project с одной выбранной задачей и несколькими свойствами задачи, содержащимися в соответствующем поле области задач](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a><span data-ttu-id="57b09-147">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="57b09-147">Next steps</span></span>

<span data-ttu-id="57b09-p108">Поздравляем, вы успешно создали надстройку Project! Теперь узнайте больше о возможностях надстроек Project и изучите распространенные сценарии.</span><span class="sxs-lookup"><span data-stu-id="57b09-p108">Congratulations, you've successfully created a Project add-in! Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="57b09-150">Надстройки Project</span><span class="sxs-lookup"><span data-stu-id="57b09-150">Project add-ins</span></span>](../project/project-add-ins.md)

# <a name="build-your-first-project-add-in"></a>Создание вашей первой надстройки Project

В этой статье рассматривается процесс создания надстройки Project с прмиенением jQuery и API JavaScript для Office.

## <a name="prerequisites"></a>Необходимые компоненты

- [Node.js](https://nodejs.org)

- Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a>Создание надстройки

1. Создайте на локальном диске папку и назовите ее `my-project-addin`. В ней будут создаваться файлы для новой надстройки.

    ```bash
    mkdir my-project-addin
    ```

2. Перейдите к новой папке.

    ```bash
    cd my-project-addin
    ```

3. Используйте генератор Yeoman для создания проекта надстройки Project. Запустите указанную ниже команду, после чего ответьте на предлагаемые вопросы следующим образом:

    ```bash
    yo office
    ```

    - **Выберите тип проекта:** `Office Add-in project using Jquery framework`
    - **Выберите тип сценария:** `Javascript`
    - **Как вы хотите назвать надстройку?:** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?** `Project`

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-project-jquery.png)
    
    После завершения работы мастера генератор создаст проект и установит поддерживающие компоненты узла.
    
4. Перейдите в корневую папку проекта веб-приложения.

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>Обновление кода

1. В редакторе кода откройте файл **index.html**, имеющийся в корневой папке проекта. Этот файл содержит HTML-содержимое, которое будет отображаться в области задач надстройки.

2. Замените элемент `<body>` на следующую разметку.

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

3. Откройте файл **src/index.js**, чтобы указать сценарий для надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.

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

4. Откройте файл **app.css** в корневой папке проекта, чтобы указать настраиваемые стили для надстройки. Замените все его содержимое следующим кодом и сохраните файл.

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

## <a name="update-the-manifest"></a>Обновление манифеста

1. Откройте файл **my-office-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки.

2. Элемент `ProviderName` содержит значение заполнителя. Замените его на свое имя.

3. Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Project**.

4. Сохраните файл.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a>Запуск сервера разработки

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>Проверьте надстройку

1. В Project создайте простой проект, содержащий хотя бы одну задачу.

2. Следуя указаниям, касающимся платформы, которая используется для запуска надстройки, загрузите неопубликованную надстройку в Project.

    - Windows: [Загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Project Online: [Загрузка неопубликованных надстроек Office в Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad и Mac: [Загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

3. Выберите задачу в Project.

    ![Снимок экрана, отображающего план проекта в Project с одной выбранной задачей](../images/project_quickstart_addin_1.png)

4. В области задач нажмите на кнопку **Получить GUID задачи**, чтобы записать GUID задачи в поле **Результаты**.

    ![Снимок экрана, отображающего план проекта в Project с одной выбранной задачей и указанном в соответствующем поле области задач GUID задачи](../images/project_quickstart_addin_2.png)

5. В области задач нажмите на кнопку **Получить данные задачи**, чтобы записать несколько свойств выбранной задачи в поле **Результаты**.

    ![Снимок экрана, отображающего план проекта в Project с одной выбранной задачей и несколькими свойствами задачи, содержащимися в соответствующем поле области задач](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a>Следующие шаги

Поздравляем, вы успешно создали надстройку Project! После этого, узнайте больше о возможностях надстроек Project и изучите распространенные сценарии.

> [!div class="nextstepaction"]
> [Надстройки Project](../project/project-add-ins.md)

# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="647f5-101">Создание первой надстройки OneNote</span><span class="sxs-lookup"><span data-stu-id="647f5-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="647f5-102">В этой статье мы разберем, как создать надстройку OneNote, используя jQuery и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="647f5-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="647f5-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="647f5-103">Prerequisites</span></span>

- [<span data-ttu-id="647f5-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="647f5-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="647f5-105">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="647f5-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="647f5-106">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="647f5-106">Create the add-in project</span></span>

1. <span data-ttu-id="647f5-p101">Используйте генератор Yeoman для создания проекта надстройки OneNote. Выполните следующую команду и затем ответьте на вопросы следующим образом:</span><span class="sxs-lookup"><span data-stu-id="647f5-p101">Use the Yeoman generator to create a OneNote add-in project. Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="647f5-109">**Выберите тип проекта:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="647f5-109">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="647f5-110">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="647f5-110">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="647f5-111">**Как вы хотите назвать надстройку?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="647f5-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="647f5-112">**Какое клиентское приложение Office должно поддерживаться?** `Onenote`</span><span class="sxs-lookup"><span data-stu-id="647f5-112">**Which Office client application would you like to support?:** `Onenote`</span></span>

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-onenote-jquery.png)
    
    <span data-ttu-id="647f5-114">После завершения работы мастера генератор создаст проект и установит поддерживающие компоненты узла.</span><span class="sxs-lookup"><span data-stu-id="647f5-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="647f5-115">Перейдите в корневую папку проекта.</span><span class="sxs-lookup"><span data-stu-id="647f5-115">Navigate to the root folder of the web application project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="647f5-116">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="647f5-116">Update the code</span></span>

1. <span data-ttu-id="647f5-p102">В редакторе кода откройте файл **index.html**, имеющийся в корневой папке проекта. Этот файл содержит HTML-содержимое, которое будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="647f5-p102">In your code editor, open **index.html** in the root of the project. This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="647f5-119">Замените элемент `<body>` следующей разметкой и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="647f5-119">Replace the `<body>` element inside the  element with the following markup and save the file.</span></span> 

    ```html
    <body class="ms-font-m ms-welcome">
        <header class="ms-welcome__header ms-bgColor-themeDark ms-u-fadeIn500">
            <h2 class="ms-fontSize-xxl ms-fontWeight-regular ms-fontColor-white">OneNote Add-in</h1>
        </header>
        <main id="app-body" class="ms-welcome__main">
            <br />
            <p class="ms-font-m">Enter HTML content here:</p>
            <div class="ms-TextField ms-TextField--placeholder">
                <textarea id="textBox" rows="8" cols="30"></textarea>
            </div>
            <button id="addOutline" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Add outline</span>
            </button>
        </main>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="647f5-120">Откройте файл **src\index.js**, чтобы указать сценарий для надстройки.</span><span class="sxs-lookup"><span data-stu-id="647f5-120">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="647f5-121">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="647f5-121">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    import * as OfficeHelpers from "@microsoft/office-js-helpers";

    Office.initialize = (reason) => {
        $(document).ready(() => {
            $('#addOutline').click(addOutlineToPage);
        });
    };
    
    async function addOutlineToPage() {
        try {
            await OneNote.run(async context => {
                var html = "<p>" + $("#textBox").val() + "</p>";

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.
                page.load("title");

                // Add text to the page by using the specified HTML.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log("Added outline to page " + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error);
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
                });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
    ```

4. <span data-ttu-id="647f5-p104">Откройте файл **app.css**, чтобы указать настраиваемые стили для надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="647f5-p104">Open the file **app.css** to specify the custom styles for the add-in. Replace the entire contents with the following and save the file.</span></span>

    ```css
    html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    ul, p, h1, h2, h3, h4, h5, h6 {
        margin: 0;
        padding: 0;
    }

    .ms-welcome {
        position: relative;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        min-height: 500px;
        min-width: 320px;
        overflow: auto;
        overflow-x: hidden;
    }

    .ms-welcome__header {
        min-height: 30px;
        padding: 0px;
        padding-bottom: 5px;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: center;
        -webkit-justify-content: flex-end;
        justify-content: flex-end;
    }

    .ms-welcome__header > h1 {
        margin-top: 5px;
        text-align: center;
    }

    .ms-welcome__main {
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: left;
        -webkit-flex: 1 0 0;
        flex: 1 0 0;
        padding: 30px 20px;
    }

    .ms-welcome__main > h2 {
        width: 100%;
        text-align: left;
    }

    @media (min-width: 0) and (max-width: 350px) {
        .ms-welcome__features {
            width: 100%;
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="647f5-124">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="647f5-124">Update the manifest</span></span>

1. <span data-ttu-id="647f5-125">Откройте файл **manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="647f5-125">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="647f5-p105">Элемент `ProviderName` содержит значение заполнителя. Замените его своим именем.</span><span class="sxs-lookup"><span data-stu-id="647f5-p105">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="647f5-p106">Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его на строку **Надстройка области задач для OneNote**.</span><span class="sxs-lookup"><span data-stu-id="647f5-p106">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="647f5-130">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="647f5-130">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="647f5-131">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="647f5-131">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="647f5-132">Проверка</span><span class="sxs-lookup"><span data-stu-id="647f5-132">Try it out</span></span>

1. <span data-ttu-id="647f5-133">Откройте записную книжку в [OneNote Online](https://www.onenote.com/notebooks).</span><span class="sxs-lookup"><span data-stu-id="647f5-133">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="647f5-134">Выберите **Вставка > Надстройки Office**. Откроется диалоговое окно "Надстройки Office".</span><span class="sxs-lookup"><span data-stu-id="647f5-134">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="647f5-135">Если вы вошли с помощью обычной учетной записи, выберите **Отправить надстройку** на вкладке **МОИ НАДСТРОЙКИ**.</span><span class="sxs-lookup"><span data-stu-id="647f5-135">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="647f5-136">Если вы вошли с помощью рабочей или учебной учетной записи, выберите **Отправить надстройку** на вкладке **МОЯ ОРГАНИЗАЦИЯ**.</span><span class="sxs-lookup"><span data-stu-id="647f5-136">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="647f5-137">На следующем изображении показана вкладка **МОИ НАДСТРОЙКИ** для потребительских записных книжек.</span><span class="sxs-lookup"><span data-stu-id="647f5-137">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="647f5-138">В диалоговом окне "Отправить надстройку" выберите файл **one-note-add-in-manifest.xml** в папке проекта и нажмите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="647f5-138">In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="647f5-p107">На вкладке **Главная** нажмите кнопку **Показать область задач** на ленте. Область задач надстройки откроется в iFrame рядом со страницей OneNote.</span><span class="sxs-lookup"><span data-stu-id="647f5-p107">From the **Home** tab, choose the **Show Taskpane** button in the ribbon. The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="647f5-141">Введите следующее HTML-содержимое в текстовом поле и нажмите кнопку **Добавить структуры**.</span><span class="sxs-lookup"><span data-stu-id="647f5-141">Enter some text in the text area and then choose **Add outline**.</span></span>  

    ```html
    <ol>
    <li>Item #1</li>
    <li>Item #2</li>
    <li>Item #3</li>
    <li>Item #4</li>
    </ol>
    ```

    <span data-ttu-id="647f5-142">Указанная вами структура добавляется на страницу.</span><span class="sxs-lookup"><span data-stu-id="647f5-142">The outline that you specified is added to the page.</span></span>

    ![Надстройка OneNote, созданная на основе этого пошагового руководства](../images/onenote-first-add-in-3.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="647f5-144">Средство устранения неполадок и советы</span><span class="sxs-lookup"><span data-stu-id="647f5-144">Troubleshooting and tips</span></span>

- <span data-ttu-id="647f5-p108">Для отладки надстройки можно использовать имеющиеся в браузере средства разработчика. При использовании веб-сервера Gulp и отладке в Internet Explorer или Chrome вы можете сохранить внесенные изменения в локальном расположении, а затем просто обновить iFrame надстройки.</span><span class="sxs-lookup"><span data-stu-id="647f5-p108">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="647f5-p109">Просматривая объект OneNote, вы увидите, что доступные для использования свойства имеют действительные значения. Свойства, которые необходимо загрузить, имеют значение *undefined*. Разверните узел `_proto_`, чтобы увидеть свойства, которые определены для объекта, но еще не загружены.</span><span class="sxs-lookup"><span data-stu-id="647f5-p109">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![Выгруженный объект OneNote в отладчике](../images/onenote-debug.png)

- <span data-ttu-id="647f5-p110">Если надстройка использует какие-либо HTTP-ресурсы, то вам потребуется включить смешанное содержимое в браузере. Надстройки, которые применяются в рабочей среде, должны использовать только безопасные HTTPS-ресурсы.</span><span class="sxs-lookup"><span data-stu-id="647f5-p110">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="647f5-153">Надстройки области задач можно открыть откуда угодно, но контентные надстройки вставляются только в содержимое стандартной страницы (не в заголовки, изображения, iFrames и т. д.).</span><span class="sxs-lookup"><span data-stu-id="647f5-153">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="647f5-154">Дальнейшие шаги</span><span class="sxs-lookup"><span data-stu-id="647f5-154">Next steps</span></span>

<span data-ttu-id="647f5-p111">Поздравляем, вы успешно ли создали надстройку OneNote! Узнайте подробнее об основных концепциях создания надстройки OneNote далее.</span><span class="sxs-lookup"><span data-stu-id="647f5-p111">Congratulations, you've successfully created a OneNote add-in! Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="647f5-157">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="647f5-157">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="647f5-158">См. также</span><span class="sxs-lookup"><span data-stu-id="647f5-158">See also</span></span>

- [<span data-ttu-id="647f5-159">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="647f5-159">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="647f5-160">Ссылка на API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="647f5-160">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [<span data-ttu-id="647f5-161">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="647f5-161">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="647f5-162">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="647f5-162">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

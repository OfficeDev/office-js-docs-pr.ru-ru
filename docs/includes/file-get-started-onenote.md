# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="30838-101">Создание первой надстройки OneNote</span><span class="sxs-lookup"><span data-stu-id="30838-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="30838-102">В этой статье мы разберем, как создать надстройку OneNote, используя jQuery и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="30838-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="30838-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="30838-103">Prerequisites</span></span>

- [<span data-ttu-id="30838-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="30838-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="30838-105">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="30838-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="30838-106">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="30838-106">Create the add-in project</span></span>

1. <span data-ttu-id="30838-107">Создайте на локальном диске папку и назовите ее `my-onenote-addin`.</span><span class="sxs-lookup"><span data-stu-id="30838-107">Create a folder on your local drive and name it `my-onenote-addin`.</span></span> <span data-ttu-id="30838-108">В ней вы будете создавать файлы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="30838-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="30838-109">Перейдите к новой папке.</span><span class="sxs-lookup"><span data-stu-id="30838-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-onenote-addin
    ```

3. <span data-ttu-id="30838-110">С помощью генератора Yeoman создайте проект надстройки OneNote.</span><span class="sxs-lookup"><span data-stu-id="30838-110">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="30838-111">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="30838-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="30838-112">**Выберите тип проекта:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="30838-112">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="30838-113">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="30838-113">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="30838-114">**Как вы хотите назвать надстройку?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="30838-114">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="30838-115">**Какое клиентское приложение Office должно поддерживаться?** `Onenote`</span><span class="sxs-lookup"><span data-stu-id="30838-115">**Which Office client application would you like to support?:** `Onenote`</span></span>

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-onenote-jquery.png)
    
    <span data-ttu-id="30838-117">После завершения работы мастера, генератор создаст проект и установит поддерживающие компоненты узла.</span><span class="sxs-lookup"><span data-stu-id="30838-117">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
4. <span data-ttu-id="30838-118">Перейдите в корневую папку проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="30838-118">Navigate to the root folder of the web application project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="30838-119">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="30838-119">Update the code</span></span>

1. <span data-ttu-id="30838-120">В редакторе кода откройте файл **index.html** из корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="30838-120">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="30838-121">Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="30838-121">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="30838-122">Замените элемент `<main>` внутри элемента `<body>` приведенной ниже разметкой и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="30838-122">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span> <span data-ttu-id="30838-123">Эта разметка добавляет текстовую область и кнопку, используя [компоненты Office UI Fabric](https://developer.microsoft.com/en-us/fabric#/components).</span><span class="sxs-lookup"><span data-stu-id="30838-123">This adds a text area and a button using [Office UI Fabric components](https://developer.microsoft.com/en-us/fabric#/components).</span></span>

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. <span data-ttu-id="30838-124">Откройте файл **src\index.js**, чтобы указать сценарий для надстройки.</span><span class="sxs-lookup"><span data-stu-id="30838-124">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="30838-125">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="30838-125">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="30838-126">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="30838-126">Update the manifest</span></span>

1. <span data-ttu-id="30838-127">Откройте файл **one-note-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="30838-127">Open the file **one-note-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="30838-128">Элемент `ProviderName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="30838-128">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="30838-129">Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="30838-129">Replace it with your name.</span></span>

3. <span data-ttu-id="30838-130">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="30838-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="30838-131">Замените его строкой **Надстройка области задач для OneNote**.</span><span class="sxs-lookup"><span data-stu-id="30838-131">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="30838-132">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="30838-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="30838-133">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="30838-133">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="30838-134">Проверка</span><span class="sxs-lookup"><span data-stu-id="30838-134">Try it out</span></span>

1. <span data-ttu-id="30838-135">Откройте записную книжку в [OneNote Online](https://www.onenote.com/notebooks).</span><span class="sxs-lookup"><span data-stu-id="30838-135">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="30838-136">Выберите **Вставка > Надстройки Office**. Откроется диалоговое окно "Надстройки Office".</span><span class="sxs-lookup"><span data-stu-id="30838-136">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="30838-137">Если вы вошли с помощью обычной учетной записи, выберите **Отправить надстройку** на вкладке **МОИ НАДСТРОЙКИ**.</span><span class="sxs-lookup"><span data-stu-id="30838-137">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="30838-138">Если вы вошли с помощью рабочей или учебной учетной записи, выберите **Отправить надстройку** на вкладке **МОЯ ОРГАНИЗАЦИЯ**.</span><span class="sxs-lookup"><span data-stu-id="30838-138">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="30838-139">На следующем изображении показана вкладка **МОИ НАДСТРОЙКИ** для обычных записных книжек.</span><span class="sxs-lookup"><span data-stu-id="30838-139">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="30838-140">В диалоговом окне "Отправить надстройку" выберите файл **one-note-add-in-manifest.xml** в папке проекта и нажмите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="30838-140">In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="30838-141">На вкладке **Главная** нажмите кнопку **Показать область задач** на ленте.</span><span class="sxs-lookup"><span data-stu-id="30838-141">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="30838-142">Надстройка откроется в iFrame рядом со страницей OneNote.</span><span class="sxs-lookup"><span data-stu-id="30838-142">6- The add-in opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="30838-143">Введите текст в текстовой области и нажмите кнопку **Добавить структуру**.</span><span class="sxs-lookup"><span data-stu-id="30838-143">Enter some text in the text area and then choose **Add outline**.</span></span> <span data-ttu-id="30838-144">Введенный текст будет добавлен на страницу.</span><span class="sxs-lookup"><span data-stu-id="30838-144">The text you entered is added to the page.</span></span> 

    ![Надстройка OneNote, созданная на основе этого руководства](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="30838-146">Устранение неполадок и советы</span><span class="sxs-lookup"><span data-stu-id="30838-146">Troubleshooting and tips</span></span>

- <span data-ttu-id="30838-p110">Для отладки надстройки можно использовать имеющиеся в браузере средства разработчика. При использовании веб-сервера Gulp и отладке в Internet Explorer или Chrome вы можете сохранить внесенные изменения в локальном расположении, а затем просто обновить iFrame надстройки.</span><span class="sxs-lookup"><span data-stu-id="30838-p110">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="30838-p111">Просматривая объект OneNote, вы увидите, что доступные для использования свойства имеют действительные значения. Свойства, которые необходимо загрузить, имеют значение *undefined*. Разверните узел `_proto_`, чтобы увидеть свойства, которые определены для объекта, но еще не загружены.</span><span class="sxs-lookup"><span data-stu-id="30838-p111">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![Выгруженный объект OneNote в отладчике](../images/onenote-debug.png)

- <span data-ttu-id="30838-p112">Если надстройка использует какие-либо HTTP-ресурсы, то вам потребуется включить смешанное содержимое в браузере. Надстройки, которые применяются в рабочей среде, должны использовать только безопасные HTTPS-ресурсы.</span><span class="sxs-lookup"><span data-stu-id="30838-p112">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="30838-155">Надстройки области задач можно открыть откуда угодно, но контентные надстройки вставляются только в содержимое стандартной страницы (не в заголовки, изображения, iFrames и т. д.).</span><span class="sxs-lookup"><span data-stu-id="30838-155">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="30838-156">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="30838-156">Next steps</span></span>

<span data-ttu-id="30838-157">Поздравляем, вы успешно создали надстройку OneNote!</span><span class="sxs-lookup"><span data-stu-id="30838-157">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="30838-158">Следующим шагом узнайте больше об основных понятиях, связанных с созданием надстроек OneNote.</span><span class="sxs-lookup"><span data-stu-id="30838-158">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="30838-159">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="30838-159">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="30838-160">См. также</span><span class="sxs-lookup"><span data-stu-id="30838-160">See also</span></span>

- [<span data-ttu-id="30838-161">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="30838-161">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="30838-162">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="30838-162">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/onenote-add-ins-javascript-reference?view=office-js)
- [<span data-ttu-id="30838-163">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="30838-163">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="30838-164">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="30838-164">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

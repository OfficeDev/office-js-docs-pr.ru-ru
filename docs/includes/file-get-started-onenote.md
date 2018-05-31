# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="478de-101">Создание первой надстройки OneNote</span><span class="sxs-lookup"><span data-stu-id="478de-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="478de-102">В этой статье мы разберем, как создать надстройку OneNote, используя jQuery и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="478de-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="478de-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="478de-103">Prerequisites</span></span>

- [<span data-ttu-id="478de-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="478de-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="478de-105">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="478de-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="478de-106">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="478de-106">Create the add-in project</span></span>

1. <span data-ttu-id="478de-107">Создайте на локальном диске папку и назовите ее `my-onenote-addin`.</span><span class="sxs-lookup"><span data-stu-id="478de-107">Create a folder on your local drive and name it `my-onenote-addin`.</span></span> <span data-ttu-id="478de-108">В ней вы будете создавать файлы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="478de-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="478de-109">Перейдите к новой папке.</span><span class="sxs-lookup"><span data-stu-id="478de-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-onenote-addin
    ```

3. <span data-ttu-id="478de-110">С помощью генератора Yeoman создайте проект надстройки OneNote.</span><span class="sxs-lookup"><span data-stu-id="478de-110">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="478de-111">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="478de-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="478de-112">**Вы хотите создать новую вложенную папку для проекта?:** `No`</span><span class="sxs-lookup"><span data-stu-id="478de-112">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="478de-113">**Как вы хотите назвать надстройку?:** `OneNote Add-in`</span><span class="sxs-lookup"><span data-stu-id="478de-113">**What do you want to name your add-in?:** `OneNote Add-in`</span></span>
    - <span data-ttu-id="478de-114">**Какое клиентское приложение Office должно поддерживаться?:** `OneNote`</span><span class="sxs-lookup"><span data-stu-id="478de-114">**Which Office client application would you like to support?:** `OneNote`</span></span>
    - <span data-ttu-id="478de-115">**Вы хотите создать новую надстройку?:** `Yes`</span><span class="sxs-lookup"><span data-stu-id="478de-115">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="478de-116">**Вы хотите использовать TypeScript?:** `No`</span><span class="sxs-lookup"><span data-stu-id="478de-116">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="478de-117">**Выберите платформу:** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="478de-117">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="478de-p103">Затем генератор предложит вам открыть файл **resource.html**. В нашем случае открывать его не обязательно, но можете заглянуть, если вам интересно! Выберите Yes (Да) или No (Нет), чтобы завершить работу мастера, и подождите, пока работает генератор.</span><span class="sxs-lookup"><span data-stu-id="478de-p103">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-onenote-jquery.png)


## <a name="update-the-code"></a><span data-ttu-id="478de-122">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="478de-122">Update the code</span></span>

1. <span data-ttu-id="478de-123">В редакторе кода откройте файл **index.html** из корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="478de-123">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="478de-124">Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="478de-124">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="478de-125">Замените элемент `<main>` внутри элемента `<body>` приведенной ниже разметкой и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="478de-125">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span> <span data-ttu-id="478de-126">Эта разметка добавляет текстовую область и кнопку, используя [компоненты Office UI Fabric](http://dev.office.com/fabric/components).</span><span class="sxs-lookup"><span data-stu-id="478de-126">This adds a text area and a button using [Office UI Fabric components](http://dev.office.com/fabric/components).</span></span>

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

3. <span data-ttu-id="478de-127">Откройте файл **app.js**, чтобы указать скрипт для надстройки.</span><span class="sxs-lookup"><span data-stu-id="478de-127">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="478de-128">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="478de-128">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

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

## <a name="update-the-manifest"></a><span data-ttu-id="478de-129">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="478de-129">Update the manifest</span></span>

1. <span data-ttu-id="478de-130">Откройте файл **one-note-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="478de-130">Open the file **one-note-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="478de-131">Элемент `ProviderName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="478de-131">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="478de-132">Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="478de-132">Replace it with your name.</span></span>

3. <span data-ttu-id="478de-133">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="478de-133">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="478de-134">Замените его строкой **Надстройка области задач для OneNote**.</span><span class="sxs-lookup"><span data-stu-id="478de-134">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="478de-135">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="478de-135">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="478de-136">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="478de-136">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="478de-137">Проверка</span><span class="sxs-lookup"><span data-stu-id="478de-137">Try it out</span></span>

1. <span data-ttu-id="478de-138">Откройте записную книжку в [OneNote Online](https://www.onenote.com/notebooks).</span><span class="sxs-lookup"><span data-stu-id="478de-138">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="478de-139">Выберите **Вставка > Надстройки Office**. Откроется диалоговое окно "Надстройки Office".</span><span class="sxs-lookup"><span data-stu-id="478de-139">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="478de-140">Если вы вошли с помощью обычной учетной записи, выберите **Отправить надстройку** на вкладке **МОИ НАДСТРОЙКИ**.</span><span class="sxs-lookup"><span data-stu-id="478de-140">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="478de-141">Если вы вошли с помощью рабочей или учебной учетной записи, выберите **Отправить надстройку** на вкладке **МОЯ ОРГАНИЗАЦИЯ**.</span><span class="sxs-lookup"><span data-stu-id="478de-141">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="478de-142">На следующем изображении показана вкладка **МОИ НАДСТРОЙКИ** для обычных записных книжек.</span><span class="sxs-lookup"><span data-stu-id="478de-142">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="478de-143">В диалоговом окне "Отправить надстройку" выберите файл **one-note-add-in-manifest.xml** в папке проекта и нажмите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="478de-143">In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="478de-144">На вкладке **Главная** нажмите кнопку **Показать область задач** на ленте.</span><span class="sxs-lookup"><span data-stu-id="478de-144">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="478de-145">Надстройка откроется в iFrame рядом со страницей OneNote.</span><span class="sxs-lookup"><span data-stu-id="478de-145">6- The add-in opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="478de-146">Введите текст в текстовой области и нажмите кнопку **Добавить структуру**.</span><span class="sxs-lookup"><span data-stu-id="478de-146">Enter some text in the text area and then choose **Add outline**.</span></span> <span data-ttu-id="478de-147">Введенный текст будет добавлен на страницу.</span><span class="sxs-lookup"><span data-stu-id="478de-147">The text you entered is added to the page.</span></span> 

    ![Надстройка OneNote, созданная на основе этого руководства](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="478de-149">Устранение неполадок и советы</span><span class="sxs-lookup"><span data-stu-id="478de-149">Troubleshooting and tips</span></span>

- <span data-ttu-id="478de-p111">Для отладки надстройки можно использовать имеющиеся в браузере средства разработчика. При использовании веб-сервера Gulp и отладке в Internet Explorer или Chrome вы можете сохранить внесенные изменения в локальном расположении, а затем просто обновить iFrame надстройки.</span><span class="sxs-lookup"><span data-stu-id="478de-p111">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="478de-p112">Просматривая объект OneNote, вы увидите, что доступные для использования свойства имеют действительные значения. Свойства, которые необходимо загрузить, имеют значение *undefined*. Разверните узел `_proto_`, чтобы увидеть свойства, которые определены для объекта, но еще не загружены.</span><span class="sxs-lookup"><span data-stu-id="478de-p112">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![Выгруженный объект OneNote в отладчике](../images/onenote-debug.png)

- <span data-ttu-id="478de-p113">Если надстройка использует какие-либо HTTP-ресурсы, то вам потребуется включить смешанное содержимое в браузере. Надстройки, которые применяются в рабочей среде, должны использовать только безопасные HTTPS-ресурсы.</span><span class="sxs-lookup"><span data-stu-id="478de-p113">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="478de-158">Надстройки области задач можно открыть откуда угодно, но контентные надстройки вставляются только в содержимое стандартной страницы (не в заголовки, изображения, iFrames и т. д.).</span><span class="sxs-lookup"><span data-stu-id="478de-158">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="478de-159">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="478de-159">Next steps</span></span>

<span data-ttu-id="478de-160">Поздравляем, вы успешно создали надстройку OneNote!</span><span class="sxs-lookup"><span data-stu-id="478de-160">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="478de-161">Следующим шагом узнайте больше об основных понятиях, связанных с созданием надстроек OneNote.</span><span class="sxs-lookup"><span data-stu-id="478de-161">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="478de-162">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="478de-162">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="478de-163">См. также</span><span class="sxs-lookup"><span data-stu-id="478de-163">See also</span></span>

- [<span data-ttu-id="478de-164">Обзор создания кода с помощью API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="478de-164">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="478de-165">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="478de-165">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="478de-166">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="478de-166">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="478de-167">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="478de-167">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

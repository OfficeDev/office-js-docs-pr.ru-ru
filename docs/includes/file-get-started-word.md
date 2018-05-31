# <a name="build-your-first-word-add-in"></a><span data-ttu-id="8f776-101">Создание первой надстройки Word</span><span class="sxs-lookup"><span data-stu-id="8f776-101">Build your first Word add-in</span></span>

<span data-ttu-id="8f776-102">_Относится к: Word 2016, Word для iPad, Word для Mac_</span><span class="sxs-lookup"><span data-stu-id="8f776-102">_Applies to: Word 2016, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="8f776-103">В этой статье мы разберем, как создать надстройку Word, используя jQuery и API JavaScript для Word.</span><span class="sxs-lookup"><span data-stu-id="8f776-103">In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="8f776-104">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="8f776-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="8f776-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="8f776-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="8f776-106">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="8f776-106">Prerequisites</span></span>

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="8f776-107">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="8f776-107">Create the add-in project</span></span>

1. <span data-ttu-id="8f776-108">В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="8f776-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="8f776-109">В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, а затем выберите **Надстройки** > **Веб-надстройка Word**.</span><span class="sxs-lookup"><span data-stu-id="8f776-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="8f776-110">Укажите имя проекта и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="8f776-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="8f776-p101">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="8f776-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="8f776-113">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="8f776-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="8f776-114">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="8f776-114">Update the code</span></span>

1. <span data-ttu-id="8f776-115">Файл **Home.html** содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="8f776-116">В файле **Home.html** замените элемент `<body>` на приведенную ниже часть кода и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="8f776-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>    
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion"/>
    </body>
    ```

2. <span data-ttu-id="8f776-117">Откройте файл **Home.js** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="8f776-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="8f776-118">Этот файл содержит скрипт надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="8f776-119">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="8f776-119">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="8f776-120">Откройте файл **Home.css** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="8f776-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="8f776-121">Этот файл определяет специальные стили надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="8f776-122">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="8f776-122">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="8f776-123">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="8f776-123">Update the manifest</span></span>

1. <span data-ttu-id="8f776-124">Откройте XML-файл манифеста в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-124">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="8f776-125">Этот файл определяет параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="8f776-126">Элемент `ProviderName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="8f776-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="8f776-127">Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="8f776-127">Replace it with your name.</span></span>

3. <span data-ttu-id="8f776-128">Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="8f776-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="8f776-129">Замените его на строку **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="8f776-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="8f776-130">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="8f776-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="8f776-131">Замените его строкой **Надстройка области задач для Word**.</span><span class="sxs-lookup"><span data-stu-id="8f776-131">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="8f776-132">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="8f776-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="8f776-133">Проверка</span><span class="sxs-lookup"><span data-stu-id="8f776-133">Try it out</span></span>

1. <span data-ttu-id="8f776-p109">Протестируйте новую надстройку Word в Visual Studio, нажав клавишу F5 или кнопку **Запустить**, чтобы запустить Word с кнопкой надстройки **Show Taskpane** (Показать область задач) на ленте. Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="8f776-p109">Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="8f776-136">В Word выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-136">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: приложение Word с выделенной кнопкой "Показать область задач"](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="8f776-138">В области задач нажмите любую кнопку, чтобы добавить стандартный текст в документ.</span><span class="sxs-lookup"><span data-stu-id="8f776-138">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Снимок экрана: приложение Word с загруженной надстройкой, добавляющей стандартный текст.](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="8f776-140">Любой редактор</span><span class="sxs-lookup"><span data-stu-id="8f776-140">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="8f776-141">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="8f776-141">Prerequisites</span></span>

- [<span data-ttu-id="8f776-142">Node.js</span><span class="sxs-lookup"><span data-stu-id="8f776-142">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="8f776-143">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="8f776-143">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a><span data-ttu-id="8f776-144">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="8f776-144">Create the add-in project</span></span>

1. <span data-ttu-id="8f776-145">Создайте на локальном диске папку и назовите ее `my-word-addin`.</span><span class="sxs-lookup"><span data-stu-id="8f776-145">Create a folder on your local drive and name it `my-word-addin`.</span></span> <span data-ttu-id="8f776-146">В ней вы будете создавать файлы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-146">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="8f776-147">Перейдите к новой папке.</span><span class="sxs-lookup"><span data-stu-id="8f776-147">Navigate to your new folder.</span></span>

    ```bash
    cd my-word-addin
    ```

3. <span data-ttu-id="8f776-148">С помощью генератора Yeoman создайте проект надстройки Word.</span><span class="sxs-lookup"><span data-stu-id="8f776-148">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="8f776-149">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="8f776-149">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="8f776-150">**Вы хотите создать новую вложенную папку для проекта?:** `No`</span><span class="sxs-lookup"><span data-stu-id="8f776-150">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="8f776-151">**Как вы хотите назвать надстройку?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="8f776-151">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="8f776-152">**Какое клиентское приложение Office должно поддерживаться?:** `Word`</span><span class="sxs-lookup"><span data-stu-id="8f776-152">**Which Office client application would you like to support?:** `Word`</span></span>
    - <span data-ttu-id="8f776-153">**Вы хотите создать новую надстройку?:** `Yes`</span><span class="sxs-lookup"><span data-stu-id="8f776-153">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="8f776-154">**Вы хотите использовать TypeScript?:** `No`</span><span class="sxs-lookup"><span data-stu-id="8f776-154">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="8f776-155">**Выберите платформу:** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="8f776-155">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="8f776-p112">Затем генератор предложит вам открыть файл **resource.html**. В нашем случае открывать его не обязательно, но можете заглянуть, если вам интересно! Выберите Yes (Да) или No (Нет), чтобы завершить работу мастера, и подождите, пока работает генератор.</span><span class="sxs-lookup"><span data-stu-id="8f776-p112">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-word-jquery.png)

### <a name="update-the-code"></a><span data-ttu-id="8f776-160">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="8f776-160">Update the code</span></span>

1. <span data-ttu-id="8f776-161">В редакторе кода откройте файл **index.html** из корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="8f776-161">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="8f776-162">Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-162">This file contains the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="8f776-163">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="8f776-163">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="8f776-164">В этой надстройке будет три кнопки. При нажатии любой из них в документ будет добавляться стандартный текст.</span><span class="sxs-lookup"><span data-stu-id="8f776-164">This add-in will display three buttons and when any of the buttons are chosen, boilerplate text will be added to the document.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <title>Boilerplate text app</title>
            <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="app.js" type="text/javascript"></script>
            <link href="app.css" rel="stylesheet" type="text/css" />
        </head>
        <body>
            <div id="content-header">
                <div class="padding">
                    <h1>Welcome</h1>
                </div>
            </div>    
            <div id="content-main">
                <div class="padding">
                    <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                    <br />
                    <h3>Try it out</h3>
                    <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                    <br /><br />
                    <button id="checkhov">Add quote from Anton Chekhov</button>
                    <br /><br />
                    <button id="proverb">Add Chinese proverb</button>
                </div>
            </div>
            <br />
            <div id="supportedVersion"/>
        </body>
    </html>
    ```

2. <span data-ttu-id="8f776-165">Откройте файл **app.js**, чтобы указать скрипт для надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-165">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="8f776-166">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="8f776-166">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="8f776-167">Этот скрипт содержит код инициализации, а также код, вносящий изменения в документ Word, вставляя текст при нажатии кнопки.</span><span class="sxs-lookup"><span data-stu-id="8f776-167">This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen.</span></span> 

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="8f776-168">Откройте файл **app.css** в корневой папке проекта, чтобы указать специальные стили для надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-168">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="8f776-169">Замените все его содержимое на приведенный ниже код и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="8f776-169">Replace the entire contents with the following and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="8f776-170">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="8f776-170">Update the manifest</span></span>

1. <span data-ttu-id="8f776-171">Откройте файл **my-office-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-171">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="8f776-172">Элемент `ProviderName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="8f776-172">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="8f776-173">Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="8f776-173">Replace it with your name.</span></span>

3. <span data-ttu-id="8f776-174">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="8f776-174">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="8f776-175">Замените его строкой **Надстройка области задач для Word**.</span><span class="sxs-lookup"><span data-stu-id="8f776-175">Replace it with **A task pane add-in for Word**.</span></span>

4. <span data-ttu-id="8f776-176">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="8f776-176">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="8f776-177">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="8f776-177">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="8f776-178">Проверка</span><span class="sxs-lookup"><span data-stu-id="8f776-178">Try it out</span></span>

1. <span data-ttu-id="8f776-179">Следуйте инструкциям для нужной платформы, чтобы загрузить неопубликованную надстройку в Word.</span><span class="sxs-lookup"><span data-stu-id="8f776-179">To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.</span></span>

    - <span data-ttu-id="8f776-180">Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="8f776-180">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="8f776-181">Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="8f776-181">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="8f776-182">iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="8f776-182">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="8f776-183">В Word выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="8f776-183">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: приложение Word с выделенной кнопкой "Показать область задач"](../images/word-quickstart-addin-2.png)

3. <span data-ttu-id="8f776-185">В области задач нажмите любую кнопку, чтобы добавить стандартный текст в документ.</span><span class="sxs-lookup"><span data-stu-id="8f776-185">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Снимок экрана: приложение Word с загруженной надстройкой, добавляющей стандартный текст.](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a><span data-ttu-id="8f776-187">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="8f776-187">Next steps</span></span>

<span data-ttu-id="8f776-188">Поздравляем, вы успешно создали надстройку Word с помощью jQuery!</span><span class="sxs-lookup"><span data-stu-id="8f776-188">Congratulations, you've successfully created a Word add-in using jQuery!</span></span> <span data-ttu-id="8f776-189">Чтобы узнать больше о возможностях надстроек Word и создать более сложную надстройку, воспользуйтесь учебным пособием по использованию надстроек Word.</span><span class="sxs-lookup"><span data-stu-id="8f776-189">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="8f776-190">Учебное пособие по использованию надстроек Word</span><span class="sxs-lookup"><span data-stu-id="8f776-190">Word add-in tutorial</span></span>](../tutorials/word-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="8f776-191">См. также</span><span class="sxs-lookup"><span data-stu-id="8f776-191">See also</span></span>

* [<span data-ttu-id="8f776-192">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="8f776-192">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* [<span data-ttu-id="8f776-193">Примеры кода надстроек Word</span><span class="sxs-lookup"><span data-stu-id="8f776-193">Word add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [<span data-ttu-id="8f776-194">Справочник по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="8f776-194">Word JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)

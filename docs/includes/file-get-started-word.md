# <a name="build-your-first-word-add-in"></a><span data-ttu-id="5a136-101">Создание первой надстройки Word</span><span class="sxs-lookup"><span data-stu-id="5a136-101">Build your first Word add-in</span></span>

<span data-ttu-id="5a136-102">_Относится к: Word 2016, Word для iPad, Word для Mac_</span><span class="sxs-lookup"><span data-stu-id="5a136-102">_Applies to: Word 2016, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="5a136-103">В этой статье мы разберем, как создать надстройку Word, используя jQuery и API JavaScript для Word.</span><span class="sxs-lookup"><span data-stu-id="5a136-103">In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="5a136-104">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="5a136-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="5a136-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="5a136-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="5a136-106">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="5a136-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="5a136-107">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="5a136-107">Create the add-in project</span></span>

1. <span data-ttu-id="5a136-108">В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="5a136-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="5a136-109">В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, а затем выберите **Надстройки** > **Веб-надстройка Word**.</span><span class="sxs-lookup"><span data-stu-id="5a136-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="5a136-110">Укажите имя проекта и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="5a136-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="5a136-p101">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="5a136-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="5a136-113">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="5a136-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="5a136-114">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="5a136-114">Update the code</span></span>

1. <span data-ttu-id="5a136-115">Файл **Home.html** содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="5a136-116">В файле **Home.html** замените элемент `<body>` на приведенную ниже часть кода и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5a136-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

2. <span data-ttu-id="5a136-117">Откройте файл **Home.js** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="5a136-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="5a136-118">Этот файл содержит скрипт надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="5a136-119">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5a136-119">Replace the entire contents with the following code and save the file.</span></span>

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

3. <span data-ttu-id="5a136-120">Откройте файл **Home.css** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="5a136-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="5a136-121">Этот файл определяет специальные стили надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="5a136-122">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5a136-122">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="5a136-123">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="5a136-123">Update the manifest</span></span>

1. <span data-ttu-id="5a136-124">Откройте XML-файл манифеста в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-124">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="5a136-125">Этот файл определяет параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="5a136-126">Элемент `ProviderName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="5a136-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="5a136-127">Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="5a136-127">Replace it with your name.</span></span>

3. <span data-ttu-id="5a136-128">Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="5a136-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="5a136-129">Замените его на строку **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="5a136-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="5a136-130">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="5a136-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="5a136-131">Замените его строкой **Надстройка области задач для Word**.</span><span class="sxs-lookup"><span data-stu-id="5a136-131">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="5a136-132">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5a136-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="5a136-133">Проверка</span><span class="sxs-lookup"><span data-stu-id="5a136-133">Try it out</span></span>

1. <span data-ttu-id="5a136-p109">Протестируйте новую надстройку Word в Visual Studio, нажав клавишу F5 или кнопку **Запустить**, чтобы запустить Word с кнопкой надстройки **Show Taskpane** (Показать область задач) на ленте. Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="5a136-p109">Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="5a136-136">В Word выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-136">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: приложение Word с выделенной кнопкой "Показать область задач"](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="5a136-138">В области задач нажмите любую кнопку, чтобы добавить стандартный текст в документ.</span><span class="sxs-lookup"><span data-stu-id="5a136-138">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Снимок экрана: приложение Word с загруженной надстройкой, добавляющей стандартный текст.](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="5a136-140">Любой редактор</span><span class="sxs-lookup"><span data-stu-id="5a136-140">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="5a136-141">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="5a136-141">Prerequisites</span></span>

- [<span data-ttu-id="5a136-142">Node.js</span><span class="sxs-lookup"><span data-stu-id="5a136-142">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="5a136-143">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="5a136-143">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a><span data-ttu-id="5a136-144">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="5a136-144">Create the add-in project</span></span>

1. <span data-ttu-id="5a136-145">Создайте на локальном диске папку и назовите ее `my-word-addin`.</span><span class="sxs-lookup"><span data-stu-id="5a136-145">Create a folder on your local drive and name it `my-word-addin`.</span></span> <span data-ttu-id="5a136-146">В ней вы будете создавать файлы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-146">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="5a136-147">Перейдите к новой папке.</span><span class="sxs-lookup"><span data-stu-id="5a136-147">Navigate to your new folder.</span></span>

    ```bash
    cd my-word-addin
    ```

3. <span data-ttu-id="5a136-148">С помощью генератора Yeoman создайте проект надстройки Word.</span><span class="sxs-lookup"><span data-stu-id="5a136-148">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="5a136-149">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="5a136-149">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="5a136-150">**Выберите тип проекта:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="5a136-150">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="5a136-151">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="5a136-151">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="5a136-152">**Как вы хотите назвать надстройку?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="5a136-152">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="5a136-153">**Какое клиентское приложение Office должно поддерживаться?** `Word`</span><span class="sxs-lookup"><span data-stu-id="5a136-153">**Which Office client application would you like to support?:** `Word`</span></span>

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-word-jquery.png)
    
    <span data-ttu-id="5a136-155">После завершения работы мастера, генератор создаст проект и установит поддерживающие компоненты узла.</span><span class="sxs-lookup"><span data-stu-id="5a136-155">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

### <a name="update-the-code"></a><span data-ttu-id="5a136-156">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="5a136-156">Update the code</span></span>

1. <span data-ttu-id="5a136-157">В редакторе кода откройте файл **index.html** из корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="5a136-157">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="5a136-158">Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-158">This file contains the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="5a136-159">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5a136-159">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="5a136-160">В этой надстройке будет три кнопки. При нажатии любой из них в документ будет добавляться стандартный текст.</span><span class="sxs-lookup"><span data-stu-id="5a136-160">This add-in will display three buttons and when any of the buttons are chosen, boilerplate text will be added to the document.</span></span>

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

2. <span data-ttu-id="5a136-161">Откройте файл **app.js**, чтобы указать скрипт для надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-161">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="5a136-162">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5a136-162">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="5a136-163">Этот скрипт содержит код инициализации, а также код, вносящий изменения в документ Word, вставляя текст при нажатии кнопки.</span><span class="sxs-lookup"><span data-stu-id="5a136-163">This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen.</span></span> 

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

3. <span data-ttu-id="5a136-164">Откройте файл **app.css** в корневой папке проекта, чтобы указать специальные стили для надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-164">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="5a136-165">Замените все его содержимое следующим кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5a136-165">Replace the entire contents with the following and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="5a136-166">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="5a136-166">Update the manifest</span></span>

1. <span data-ttu-id="5a136-167">Откройте файл **my-office-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-167">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="5a136-168">Элемент `ProviderName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="5a136-168">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="5a136-169">Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="5a136-169">Replace it with your name.</span></span>

3. <span data-ttu-id="5a136-170">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="5a136-170">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="5a136-171">Замените его строкой **Надстройка области задач для Word**.</span><span class="sxs-lookup"><span data-stu-id="5a136-171">Replace it with **A task pane add-in for Word**.</span></span>

4. <span data-ttu-id="5a136-172">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5a136-172">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="5a136-173">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="5a136-173">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="5a136-174">Проверка</span><span class="sxs-lookup"><span data-stu-id="5a136-174">Try it out</span></span>

1. <span data-ttu-id="5a136-175">Следуйте инструкциям для нужной платформы, чтобы загрузить неопубликованную надстройку в Word.</span><span class="sxs-lookup"><span data-stu-id="5a136-175">To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.</span></span>

    - <span data-ttu-id="5a136-176">Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="5a136-176">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="5a136-177">Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="5a136-177">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="5a136-178">iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="5a136-178">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="5a136-179">В Word выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="5a136-179">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: приложение Word с выделенной кнопкой "Показать область задач"](../images/word-quickstart-addin-2.png)

3. <span data-ttu-id="5a136-181">В области задач нажмите любую кнопку, чтобы добавить стандартный текст в документ.</span><span class="sxs-lookup"><span data-stu-id="5a136-181">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Снимок экрана: приложение Word с загруженной надстройкой, добавляющей стандартный текст.](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a><span data-ttu-id="5a136-183">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="5a136-183">Next steps</span></span>

<span data-ttu-id="5a136-184">Поздравляем, вы успешно создали надстройку Word с помощью jQuery!</span><span class="sxs-lookup"><span data-stu-id="5a136-184">Congratulations, you've successfully created a Word add-in using jQuery!</span></span> <span data-ttu-id="5a136-185">Чтобы узнать больше о возможностях надстроек Word и создать более сложную надстройку, воспользуйтесь руководством по использованию надстроек Word.</span><span class="sxs-lookup"><span data-stu-id="5a136-185">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5a136-186">Руководство по надстройкам Word</span><span class="sxs-lookup"><span data-stu-id="5a136-186">Word add-in tutorial</span></span>](../tutorials/word-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="5a136-187">См. также</span><span class="sxs-lookup"><span data-stu-id="5a136-187">See also</span></span>

* [<span data-ttu-id="5a136-188">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="5a136-188">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* [<span data-ttu-id="5a136-189">Примеры кода надстроек Word</span><span class="sxs-lookup"><span data-stu-id="5a136-189">Word add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [<span data-ttu-id="5a136-190">Справочник по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="5a136-190">Word JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)

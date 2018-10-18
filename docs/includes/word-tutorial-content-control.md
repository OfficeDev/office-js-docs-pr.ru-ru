<span data-ttu-id="2d9b0-101">На этом этапе руководства мы рассмотрим создание элементов управления форматированным текстом в документе, а также вставку и замену содержимого этих элементов.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-101">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span> 

> [!NOTE]
> <span data-ttu-id="2d9b0-p101">На этой странице описывается отдельный этап из руководства по надстройкам Word. Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Word](../tutorials/word-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

<span data-ttu-id="2d9b0-104">Прежде чем приступать к этому этапу руководства, рекомендуем создать элементы управления форматированным текстом и управлять ими через пользовательский интерфейс Word, чтобы получить представление об этих элементах и их свойствах.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-104">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="2d9b0-105">Дополнительные сведения см. в статье [Создание форм, предназначенных для заполнения или печати в приложении Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span><span class="sxs-lookup"><span data-stu-id="2d9b0-105">For details, see [Create forms that users complete or print in Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

> [!NOTE]
> <span data-ttu-id="2d9b0-106">Существует несколько типов элементов управления содержимым, которые можно добавить в документ Word через пользовательский интерфейс. Однако в настоящее время Word.js поддерживает только элементы управления форматированным текстом.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-106">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>


## <a name="create-a-content-control"></a><span data-ttu-id="2d9b0-107">Создание элемента управления содержимым</span><span class="sxs-lookup"><span data-stu-id="2d9b0-107">Create a content control</span></span>

1. <span data-ttu-id="2d9b0-108">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-108">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="2d9b0-109">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-109">Open the file index.html.</span></span>
3. <span data-ttu-id="2d9b0-110">Под элементом `div`, содержащим кнопку `replace-text`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="2d9b0-110">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. <span data-ttu-id="2d9b0-111">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-111">Open the app.js file.</span></span>

5. <span data-ttu-id="2d9b0-112">Под строкой, назначающей обработчик нажатия кнопки `insert-table`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="2d9b0-112">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="2d9b0-113">Добавьте приведенную ниже функцию под функцией `insertTable`.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-113">Below the `insertTable` function, add the following function:</span></span>

    ```js
    function createContentControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to create a content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. <span data-ttu-id="2d9b0-p103">Замените `TODO1` на приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="2d9b0-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="2d9b0-116">Этот код заключает фразу "Office 365" в элемент управления содержимым.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-116">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="2d9b0-117">Для простоты предполагается, что такая строка существует и пользователь выделил ее.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-117">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="2d9b0-118">Свойство `ContentControl.title` задает видимый заголовок элемента управления содержимым.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-118">The `ContentControl.title` property specifies the visible title of the content control.</span></span> 
   - <span data-ttu-id="2d9b0-119">Свойство `ContentControl.tag` задает тег, с помощью которого можно получить ссылку на элемент управления содержимым путем вызова метода `ContentControlCollection.getByTag`, который будет использоваться в последующей функции.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-119">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span> 
   - <span data-ttu-id="2d9b0-120">Свойство `ContentControl.appearance` задает внешний вид элемента управления.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-120">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="2d9b0-121">Значение Tags указывает, что элемент управления будет заключен в открывающие и закрывающие теги, а открывающий тег будет содержать заголовок элемента управления содержимым.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-121">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="2d9b0-122">Другие возможные значения: BoundingBox и None.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-122">Other possible values are "BoundingBox" and "None".</span></span>
   - <span data-ttu-id="2d9b0-123">Свойство `ContentControl.color` задает цвет тегов или рамки ограничивающего прямоугольника.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-123">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="2d9b0-124">Замена содержимого элемента управления</span><span class="sxs-lookup"><span data-stu-id="2d9b0-124">Replace the content of the content control</span></span>

1. <span data-ttu-id="2d9b0-125">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-125">Open the file index.html.</span></span>
2. <span data-ttu-id="2d9b0-126">Под элементом `div`, содержащим кнопку `create-content-control`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="2d9b0-126">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

3. <span data-ttu-id="2d9b0-127">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-127">Open the app.js file.</span></span>

4. <span data-ttu-id="2d9b0-128">Под строкой, назначающей обработчик нажатия кнопки `create-content-control`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="2d9b0-128">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. <span data-ttu-id="2d9b0-129">Добавьте приведенную ниже функцию под функцией `createContentControl`.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-129">Below the `createContentControl` function, add the following function:</span></span>

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. <span data-ttu-id="2d9b0-130">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-130">Replace `TODO1` with the following code.</span></span> 
    > [!NOTE]
    > <span data-ttu-id="2d9b0-131">Метод возвращает `ContentControlCollection` всех элементов управления содержимым указанного тега.`ContentControlCollection.getByTag`</span><span class="sxs-lookup"><span data-stu-id="2d9b0-131">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="2d9b0-132">Мы используем `getFirst` чтобы получить ссылку на требуемый элемент управления.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-132">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="2d9b0-133">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="2d9b0-133">Test the add-in</span></span>

1. <span data-ttu-id="2d9b0-134">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-134">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="2d9b0-135">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-135">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
     > [!NOTE]
     > <span data-ttu-id="2d9b0-136">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-136">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="2d9b0-137">Для этого необходимо завершить процесс сервера, чтобы появился запрос и вы могли ввести команду сборки.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-137">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="2d9b0-138">После сборки перезапустите сервер.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-138">After the build, restart the server.</span></span> <span data-ttu-id="2d9b0-139">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-139">The next few steps carry out this process.</span></span>
2. <span data-ttu-id="2d9b0-140">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-140">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="2d9b0-141">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-141">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="2d9b0-142">Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-142">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="2d9b0-143">В области задач нажмите кнопку **Insert Paragraph** (Вставить абзац), чтобы убедиться, что в начале документа есть абзац с фразой "Office 365".</span><span class="sxs-lookup"><span data-stu-id="2d9b0-143">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>
6. <span data-ttu-id="2d9b0-144">Выделите фразу "Office 365" в добавленном абзаце, а затем нажмите кнопку **Create Content Control** (Создать элемент управления содержимым).</span><span class="sxs-lookup"><span data-stu-id="2d9b0-144">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="2d9b0-145">Обратите внимание, что фраза заключена в теги с меткой Service Name.</span><span class="sxs-lookup"><span data-stu-id="2d9b0-145">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>
7. <span data-ttu-id="2d9b0-146">Нажмите кнопку **Rename Service** (Переименовать службу) и обратите внимание, что текст элемента управления содержимым меняется на "Fabrikam Online Productivity Suite".</span><span class="sxs-lookup"><span data-stu-id="2d9b0-146">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Руководство по Word: создание элемента управления содержимым и изменение его текста](../images/word-tutorial-content-control.png)

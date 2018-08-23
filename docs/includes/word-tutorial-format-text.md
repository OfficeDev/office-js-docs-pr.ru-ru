<span data-ttu-id="9e3b0-101">На этом этапе руководства мы изменим шрифт текста и применим к нему как встроенные, так и пользовательские стили.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-101">In this step of the tutorial, you'll change the font of text, and use both built-in and custom styles on the text.</span></span>

> [!NOTE]
> <span data-ttu-id="9e3b0-p101">На этой странице описывается отдельный этап из руководства по надстройкам Word. Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Word](../tutorials/word-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="9e3b0-104">Применение встроенного стиля к тексту</span><span class="sxs-lookup"><span data-stu-id="9e3b0-104">Apply a built-in style to text</span></span>

1. <span data-ttu-id="9e3b0-105">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="9e3b0-106">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-106">Open the file index.html.</span></span>
3. <span data-ttu-id="9e3b0-107">Под элементом `div`, содержащим кнопку `insert-paragraph`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="9e3b0-107">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="9e3b0-108">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-108">Open the app.js file.</span></span>

5. <span data-ttu-id="9e3b0-109">Под строкой, назначающей обработчик нажатия кнопки `insert-paragraph`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="9e3b0-109">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="9e3b0-110">Под функцией `insertParagraph` добавьте следующую функцию:</span><span class="sxs-lookup"><span data-stu-id="9e3b0-110">Just below the `insertParagraph` function, add the following function:</span></span>

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

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

7. <span data-ttu-id="9e3b0-111">Замените `TODO1` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-111">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="9e3b0-112">Обратите внимание, что этот код применяет стиль к абзацу, но стили также можно применять к диапазонам текста.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-112">Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="9e3b0-113">Применение пользовательского стиля к тексту</span><span class="sxs-lookup"><span data-stu-id="9e3b0-113">Apply a custom style to text</span></span>

1. <span data-ttu-id="9e3b0-114">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-114">Open the file index.html.</span></span>
2. <span data-ttu-id="9e3b0-115">Под элементом `div`, содержащим кнопку `apply-style`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="9e3b0-115">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="9e3b0-116">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-116">Open the app.js file.</span></span>

4. <span data-ttu-id="9e3b0-117">Под строкой, назначающей обработчик нажатия кнопки `apply-style`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="9e3b0-117">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="9e3b0-118">Добавьте приведенную ниже функцию под функцией `applyStyle`.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-118">Below the `applyStyle` function, add the following function:</span></span>

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

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

7. <span data-ttu-id="9e3b0-119">Замените `TODO1` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-119">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="9e3b0-120">Обратите внимание, что этот код применяет пользовательский стиль, который еще не существует.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-120">Note that the code applies a custom style that does not exist yet.</span></span> <span data-ttu-id="9e3b0-121">Мы создадим стиль с именем **MyCustomStyle** во время [тестирования настройки](#test-the-add-in).</span><span class="sxs-lookup"><span data-stu-id="9e3b0-121">You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## <a name="change-the-font-of-text"></a><span data-ttu-id="9e3b0-122">Изменение шрифта для текста</span><span class="sxs-lookup"><span data-stu-id="9e3b0-122">Change the font of text</span></span>

1. <span data-ttu-id="9e3b0-123">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-123">Open the file index.html.</span></span>
2. <span data-ttu-id="9e3b0-124">Под элементом `div`, содержащим кнопку `apply-custom-style`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="9e3b0-124">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="9e3b0-125">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-125">Open the app.js file.</span></span>

4. <span data-ttu-id="9e3b0-126">Под строкой, назначающей обработчик нажатия кнопки `apply-custom-style`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="9e3b0-126">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="9e3b0-127">Добавьте приведенную ниже функцию под функцией `applyCustomStyle`.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-127">Below the `applyCustomStyle` function, add the following function:</span></span>

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

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

7. <span data-ttu-id="9e3b0-128">Замените `TODO1` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-128">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="9e3b0-129">Обратите внимание, что этот код получает ссылку на второй абзац с помощью метода `ParagraphCollection.getFirst`, привязанного к методу `Paragraph.getNext`.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-129">Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="9e3b0-130">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="9e3b0-130">Test the add-in</span></span>

1. <span data-ttu-id="9e3b0-131">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="9e3b0-132">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="9e3b0-133">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="9e3b0-134">Для этого необходимо завершить процесс сервера, чтобы появился запрос и вы могли ввести команду сборки.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-134">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="9e3b0-135">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-135">After the build, you restart the server.</span></span> <span data-ttu-id="9e3b0-136">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-136">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="9e3b0-137">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="9e3b0-138">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-138">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="9e3b0-139">Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-139">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="9e3b0-140">Убедитесь, что в тексте есть по крайней мере три абзаца.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-140">Be sure there are at least three paragraphs in the document.</span></span> <span data-ttu-id="9e3b0-141">Вы можете три раза нажать кнопку **Insert Paragraph** (Вставить абзац).</span><span class="sxs-lookup"><span data-stu-id="9e3b0-141">You can choose **Insert Paragraph** three times.</span></span> <span data-ttu-id="9e3b0-142">*Внимательно проверьте, нет ли в конце документа пустого абзаца. Если он есть, удалите его.*</span><span class="sxs-lookup"><span data-stu-id="9e3b0-142">*Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>
6. <span data-ttu-id="9e3b0-143">В Word создайте пользовательский стиль с именем "MyCustomStyle".</span><span class="sxs-lookup"><span data-stu-id="9e3b0-143">In Word, create a custom style named "MyCustomStyle".</span></span> <span data-ttu-id="9e3b0-144">Его форматирование может быть любым.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-144">It can have any formatting that you want.</span></span>
7. <span data-ttu-id="9e3b0-145">Нажмите кнопку **Apply Style** (Применить стиль).</span><span class="sxs-lookup"><span data-stu-id="9e3b0-145">Choose the **Apply Style** button.</span></span> <span data-ttu-id="9e3b0-146">К первому абзацу будет применен встроенный стиль **Сильная ссылка**.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-146">The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>
8. <span data-ttu-id="9e3b0-147">Нажмите кнопку **Apply Custom Style** (Применить пользовательский стиль).</span><span class="sxs-lookup"><span data-stu-id="9e3b0-147">Choose the **Apply Custom Style** button.</span></span> <span data-ttu-id="9e3b0-148">К последнему абзацу будет применен созданный вами стиль.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-148">The last paragraph will be styled with your custom style.</span></span> <span data-ttu-id="9e3b0-149">Если ничего не происходит, возможно, последний абзац пуст.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-149">(If nothing seems to happen, the last paragraph might be blank.</span></span> <span data-ttu-id="9e3b0-150">Если это так, добавьте в него какой-нибудь текст.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-150">If so, add some text to it.)</span></span>
9. <span data-ttu-id="9e3b0-151">Нажмите кнопку **Change Font** (Изменить шрифт).</span><span class="sxs-lookup"><span data-stu-id="9e3b0-151">Choose the **Change Font** button.</span></span> <span data-ttu-id="9e3b0-152">Шрифт второго абзаца изменится на полужирный Courier New с размером 18.</span><span class="sxs-lookup"><span data-stu-id="9e3b0-152">The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Руководство по Word: применение стилей и шрифта](../images/word-tutorial-apply-styles-and-font.png)

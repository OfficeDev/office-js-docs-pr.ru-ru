<span data-ttu-id="4f900-101">На этом этапе руководства мы рассмотрим вставку изображений, HTML-кода и таблиц в документ.</span><span class="sxs-lookup"><span data-stu-id="4f900-101">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

> [!NOTE]
> <span data-ttu-id="4f900-p101">На этой странице описывается отдельный этап из руководства по надстройкам Word. Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Word](../tutorials/word-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="4f900-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="insert-an-image"></a><span data-ttu-id="4f900-104">Вставка изображения</span><span class="sxs-lookup"><span data-stu-id="4f900-104">Insert an image</span></span>

1. <span data-ttu-id="4f900-105">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="4f900-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="4f900-106">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="4f900-106">Open the file index.html.</span></span>
3. <span data-ttu-id="4f900-107">Под элементом `div`, содержащим кнопку `replace-text`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="4f900-107">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-image">Insert Image</button>            
    </div>
    ```

4. <span data-ttu-id="4f900-108">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="4f900-108">Open the app.js file.</span></span>

5. <span data-ttu-id="4f900-109">Добавьте приведенную ниже строку сразу под строкой use-strict в верхней части файла.</span><span class="sxs-lookup"><span data-stu-id="4f900-109">Near the top of the file, just below the use-strict line, add the following line.</span></span> <span data-ttu-id="4f900-110">Эта строка импортирует переменную из другого файла.</span><span class="sxs-lookup"><span data-stu-id="4f900-110">This line imports a variable from another file.</span></span> <span data-ttu-id="4f900-111">Переменная представляет собой строку с кодировкой Base 64, кодирующую изображение.</span><span class="sxs-lookup"><span data-stu-id="4f900-111">The variable is a base 64 string that encodes an image.</span></span> <span data-ttu-id="4f900-112">Чтобы просмотреть закодированную строку, откройте файл base64Image.js в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="4f900-112">To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ``` 

5. <span data-ttu-id="4f900-113">Под строкой, назначающей обработчик нажатия кнопки `replace-text`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="4f900-113">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

6. <span data-ttu-id="4f900-114">Добавьте приведенную ниже функцию под функцией `replaceText`.</span><span class="sxs-lookup"><span data-stu-id="4f900-114">Below the `replaceText` function, add the following function:</span></span>

    ```js
    function insertImage() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert an image.

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

7. <span data-ttu-id="4f900-115">Замените `TODO1` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="4f900-115">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="4f900-116">Обратите внимание, что эта строка вставляет изображение с кодировкой Base 64 в конце документа.</span><span class="sxs-lookup"><span data-stu-id="4f900-116">Note that this line inserts the base 64 encoded image at the end of the document.</span></span> <span data-ttu-id="4f900-117">У объекта `Paragraph` также есть метод `insertInlinePictureFromBase64` и другие методы `insert*`.</span><span class="sxs-lookup"><span data-stu-id="4f900-117">(The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods.</span></span> <span data-ttu-id="4f900-118">Пример представлен в следующем разделе, посвященном вставке HTML.</span><span class="sxs-lookup"><span data-stu-id="4f900-118">See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ``` 

## <a name="insert-html"></a><span data-ttu-id="4f900-119">Вставка HTML</span><span class="sxs-lookup"><span data-stu-id="4f900-119">Insert HTML</span></span>

1. <span data-ttu-id="4f900-120">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="4f900-120">Open the file index.html.</span></span>
2. <span data-ttu-id="4f900-121">Под элементом `div`, содержащим кнопку `insert-image`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="4f900-121">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-html">Insert HTML</button>            
    </div>
    ```

3. <span data-ttu-id="4f900-122">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="4f900-122">Open the app.js file.</span></span>

4. <span data-ttu-id="4f900-123">Под строкой, назначающей обработчик нажатия кнопки `insert-image`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="4f900-123">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="4f900-124">Добавьте приведенную ниже функцию под функцией `insertImage`.</span><span class="sxs-lookup"><span data-stu-id="4f900-124">Below the `insertImage` function, add the following function:</span></span>

    ```js
    function insertHTML() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert a string of HTML.

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

6. <span data-ttu-id="4f900-p104">Замените `TODO1` на приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="4f900-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="4f900-127">Первая строка добавляет пустой абзац в конце документа.</span><span class="sxs-lookup"><span data-stu-id="4f900-127">The first line adds a blank paragraph to the end of the document.</span></span> 
   - <span data-ttu-id="4f900-128">Вторая команда вставляет строку HTML-кода в конце абзаца. В частности, вставляются два абзаца, в одном из которых используется шрифт Verdana, а в другом — стандартный стиль документа Word.</span><span class="sxs-lookup"><span data-stu-id="4f900-128">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document.</span></span> <span data-ttu-id="4f900-129">Как видно по вышеописанному методу `insertImage`, у объекта `context.document.body` также есть методы `insert*`.</span><span class="sxs-lookup"><span data-stu-id="4f900-129">(As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ``` 

## <a name="insert-table"></a><span data-ttu-id="4f900-130">Вставка таблицы</span><span class="sxs-lookup"><span data-stu-id="4f900-130">Insert Table</span></span>

1. <span data-ttu-id="4f900-131">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="4f900-131">Open the file index.html.</span></span>
3. <span data-ttu-id="4f900-132">Под элементом `div`, содержащим кнопку `insert-html`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="4f900-132">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-table">Insert Table</button>            
    </div>
    ```

4. <span data-ttu-id="4f900-133">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="4f900-133">Open the app.js file.</span></span>

5. <span data-ttu-id="4f900-134">Под строкой, назначающей обработчик нажатия кнопки `insert-html`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="4f900-134">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

6. <span data-ttu-id="4f900-135">Добавьте приведенную ниже функцию под функцией `insertHTML`.</span><span class="sxs-lookup"><span data-stu-id="4f900-135">Below the `insertHTML` function, add the following function:</span></span>

    ```js
    function insertTable() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

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

7. <span data-ttu-id="4f900-136">Замените `TODO1` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="4f900-136">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="4f900-137">Обратите внимание, что в этой строке используется метод `ParapgraphCollection.getFirst`, чтобы получить ссылку на первый абзац, а затем — метод `Paragraph.getNext`, чтобы получить ссылку на второй абзац.</span><span class="sxs-lookup"><span data-stu-id="4f900-137">Note that this line uses the `ParapgraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ``` 

8. <span data-ttu-id="4f900-p107">Замените `TODO2` на приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="4f900-p107">Replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="4f900-140">Первые два параметра метода `insertTable` задают количество строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="4f900-140">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>
   - <span data-ttu-id="4f900-141">Третий параметр указывает, где вставить таблицу (в данном случае — после абзаца).</span><span class="sxs-lookup"><span data-stu-id="4f900-141">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>
   - <span data-ttu-id="4f900-142">Четвертый параметр представляет собой двумерный массив, задающий значения ячеек таблицы.</span><span class="sxs-lookup"><span data-stu-id="4f900-142">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>
   - <span data-ttu-id="4f900-143">К таблице применяется простой стиль по умолчанию, но метод `insertTable` возвращает объект `Table` со множеством элементов, некоторые из которых используются для настройки стиля таблицы.</span><span class="sxs-lookup"><span data-stu-id="4f900-143">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

     ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="4f900-144">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="4f900-144">Test the add-in</span></span>


1. <span data-ttu-id="4f900-145">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="4f900-145">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="4f900-146">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="4f900-146">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="4f900-147">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="4f900-147">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="4f900-148">Для этого необходимо завершить процесс сервера, чтобы появился запрос и вы могли ввести команду сборки.</span><span class="sxs-lookup"><span data-stu-id="4f900-148">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="4f900-149">После сборки перезапустите сервер.</span><span class="sxs-lookup"><span data-stu-id="4f900-149">After the build, restart the server.</span></span> <span data-ttu-id="4f900-150">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="4f900-150">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="4f900-151">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="4f900-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="4f900-152">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="4f900-152">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="4f900-153">Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="4f900-153">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="4f900-154">В области задач нажмите кнопку **Insert Paragraph** (Вставить абзац) не менее трех раз, чтобы убедиться, что в документе есть несколько абзацев.</span><span class="sxs-lookup"><span data-stu-id="4f900-154">In the taskpane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>
6. <span data-ttu-id="4f900-155">Нажмите кнопку **Insert Image** (Вставить изображение) и обратите внимание, что изображение вставляется в конце документа.</span><span class="sxs-lookup"><span data-stu-id="4f900-155">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>
7. <span data-ttu-id="4f900-156">Нажмите кнопку **Insert HTML** (Вставить HTML) и обратите внимание, что в конце документа вставляются два абзаца, в первом из которых используется шрифт Verdana.</span><span class="sxs-lookup"><span data-stu-id="4f900-156">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>
8. <span data-ttu-id="4f900-157">Нажмите кнопку **Insert Table** (Вставить таблицу) и обратите внимание, что после второго абзаца вставляется таблица.</span><span class="sxs-lookup"><span data-stu-id="4f900-157">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Руководство по Word: вставка изображения, HTML-кода и таблицы](../images/word-tutorial-insert-image-html-table.png)

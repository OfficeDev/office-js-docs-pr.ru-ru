<span data-ttu-id="3e905-101">Когда таблица достаточно длинная, при прокрутке строка заголовков может исчезать с экрана.</span><span class="sxs-lookup"><span data-stu-id="3e905-101">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="3e905-102">В этом разделе учебника мы расскажем, как закрепить строку заголовков созданной ранее таблицы, чтобы она была видна, даже когда пользователь прокручивает лист.</span><span class="sxs-lookup"><span data-stu-id="3e905-102">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span> 

> [!NOTE]
> <span data-ttu-id="3e905-103">Это один из разделов руководства по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="3e905-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="3e905-104">Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Excel](../tutorials/excel-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="3e905-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="freeze-the-tables-header-row"></a><span data-ttu-id="3e905-105">Закрепление строки заголовков таблицы</span><span class="sxs-lookup"><span data-stu-id="3e905-105">Freeze the table's header row</span></span>

1. <span data-ttu-id="3e905-106">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="3e905-106">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="3e905-107">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="3e905-107">Open the file index.html.</span></span>
3. <span data-ttu-id="3e905-108">Под элементом `div`, содержащим кнопку `create-chart`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="3e905-108">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="freeze-header">Freeze Header</button>            
    </div>
    ```

4. <span data-ttu-id="3e905-109">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="3e905-109">Open the app.js file.</span></span>

5. <span data-ttu-id="3e905-110">Под строкой, назначающей обработчик нажатия кнопки `create-chart`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="3e905-110">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="3e905-111">Под функцией `createChart` добавьте следующую функцию:</span><span class="sxs-lookup"><span data-stu-id="3e905-111">Below the `createChart` function add the following function:</span></span>

    ```js
    function freezeHeader() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to keep the header visible when the user scrolls.

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

7. <span data-ttu-id="3e905-p103">Замените `TODO1` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="3e905-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="3e905-114">Коллекция `Worksheet.freezePanes` — это набор закрепленных строк, которые не исчезают с экрана при прокрутке листа.</span><span class="sxs-lookup"><span data-stu-id="3e905-114">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>
   - <span data-ttu-id="3e905-p104">Метод `freezeRows` принимает в качестве параметра количество строк сверху, которые необходимо закрепить. Мы передаем значение `1`, чтобы закрепить первую строку.</span><span class="sxs-lookup"><span data-stu-id="3e905-p104">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="3e905-117">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="3e905-117">Test the add-in</span></span>

1. <span data-ttu-id="3e905-118">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="3e905-118">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="3e905-119">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="3e905-119">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3e905-120">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="3e905-120">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="3e905-121">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="3e905-121">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="3e905-122">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="3e905-122">After the build, you restart the server.</span></span> <span data-ttu-id="3e905-123">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="3e905-123">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="3e905-124">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="3e905-124">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="3e905-125">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="3e905-125">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="3e905-126">Повторно загрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**.</span><span class="sxs-lookup"><span data-stu-id="3e905-126">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="3e905-127">Если таблица на листе, удалите ее.</span><span class="sxs-lookup"><span data-stu-id="3e905-127">If the table is in the worksheet, delete it.</span></span>
7. <span data-ttu-id="3e905-128">В области задач нажмите кнопку **Create Table** (Создать таблицу).</span><span class="sxs-lookup"><span data-stu-id="3e905-128">In the taskpane, choose **Create Table**.</span></span> 
8. <span data-ttu-id="3e905-129">Нажмите кнопку **Freeze Header** (Закрепить заголовок).</span><span class="sxs-lookup"><span data-stu-id="3e905-129">Choose the **Freeze Header** button.</span></span>
9. <span data-ttu-id="3e905-130">Прокрутите лист вниз, чтобы убедиться, что заголовок таблицы по-прежнему остается на экране, даже когда более высокие строки исчезают.</span><span class="sxs-lookup"><span data-stu-id="3e905-130">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Учебник Excel | Закрепление заголовка](../images/excel-tutorial-freeze-header.png)

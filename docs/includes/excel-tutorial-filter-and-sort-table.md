<span data-ttu-id="af076-101">Из этого раздела руководства вы узнаете, как отфильтровать и отсортировать созданную ранее таблицу.</span><span class="sxs-lookup"><span data-stu-id="af076-101">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

> [!NOTE]
> <span data-ttu-id="af076-102">Это один из разделов руководства по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="af076-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="af076-103">Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Excel](../tutorials/excel-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="af076-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="filter-the-table"></a><span data-ttu-id="af076-104">Фильтрация таблицы</span><span class="sxs-lookup"><span data-stu-id="af076-104">Filter the table</span></span>

1. <span data-ttu-id="af076-105">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="af076-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="af076-106">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="af076-106">Open the file index.html.</span></span>
3. <span data-ttu-id="af076-107">Под элементом `div`, содержащим кнопку `create-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="af076-107">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. <span data-ttu-id="af076-108">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="af076-108">Open the app.js file.</span></span>

5. <span data-ttu-id="af076-109">Под строкой, назначающей обработчик нажатия кнопки `create-table`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="af076-109">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="af076-110">Под функцией `createTable` добавьте следующую функцию:</span><span class="sxs-lookup"><span data-stu-id="af076-110">Just below the `createTable` function, add the following function:</span></span>

    ```js
    function filterTable() {
        Excel.run(function (context) {

            // TODO1: Queue commands to filter out all expense categories except
            //        Groceries and Education.

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

7. <span data-ttu-id="af076-p102">Замените `TODO1` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="af076-p102">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="af076-113">Код получает ссылку на столбец, который нужно отфильтровать, передавая имя столбца методу `getItem`, а не передавая его индекс методу `getItemAt`, как это делает метод `createTable`.</span><span class="sxs-lookup"><span data-stu-id="af076-113">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does.</span></span> <span data-ttu-id="af076-114">Так как пользователи могут перемещать столбцы, по заданному индексу может располагаться уже другой столбец.</span><span class="sxs-lookup"><span data-stu-id="af076-114">Since users can move table columns, the column at a given index might change after the table is created.</span></span> <span data-ttu-id="af076-115">Следовательно, для получения ссылки безопаснее использовать имя столбца.</span><span class="sxs-lookup"><span data-stu-id="af076-115">Hence, it is safer to use the column name to get a reference to the column.</span></span> <span data-ttu-id="af076-116">Мы спокойно использовали метод `getItemAt` в предыдущем разделе, потому что мы использовали его в методе, который создает таблицу, и пользователь никак не мог переместить столбец.</span><span class="sxs-lookup"><span data-stu-id="af076-116">We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>
   - <span data-ttu-id="af076-117">Метод `applyValuesFilter` является одним из нескольких методов фильтрации объекта `Filter`.</span><span class="sxs-lookup"><span data-stu-id="af076-117">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

## <a name="sort-the-table"></a><span data-ttu-id="af076-118">Сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="af076-118">Sort the table</span></span>

1. <span data-ttu-id="af076-119">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="af076-119">Open the file index.html.</span></span>
2. <span data-ttu-id="af076-120">Под элементом `div`, содержащим кнопку `filter-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="af076-120">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. <span data-ttu-id="af076-121">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="af076-121">Open the app.js file.</span></span>

4. <span data-ttu-id="af076-122">Под строкой, назначающей обработчик нажатия кнопки `filter-table`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="af076-122">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="af076-123">Под функцией `filterTable` добавьте приведенную ниже функцию.</span><span class="sxs-lookup"><span data-stu-id="af076-123">Below the `filterTable` function add the following function.</span></span>

    ```js
    function sortTable() {
        Excel.run(function (context) {

            // TODO1: Queue commands to sort the table by Merchant name.

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

7. <span data-ttu-id="af076-p104">Замените `TODO1` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="af076-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="af076-126">Код создает массив объектов `SortField`, состоящий из одного элемента, так как надстройка сортирует таблицу только по столбцу Merchant.</span><span class="sxs-lookup"><span data-stu-id="af076-126">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>
   - <span data-ttu-id="af076-127">Свойство `key` объекта `SortField` — это отсчитываемый от нуля индекс столбца, по которому необходимо сортировать таблицу.</span><span class="sxs-lookup"><span data-stu-id="af076-127">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>
   - <span data-ttu-id="af076-128">Элемент `sort` объекта `Table` — это объект `TableSort`, а не метод.</span><span class="sxs-lookup"><span data-stu-id="af076-128">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="af076-129">Объекты `SortField` передаются методу `apply` объекта `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="af076-129">The `SortField`s are passed the `TableSort` object's `apply` method.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="af076-130">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="af076-130">Test the add-in</span></span>

1. <span data-ttu-id="af076-131">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="af076-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="af076-132">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="af076-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="af076-133">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="af076-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="af076-134">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="af076-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="af076-135">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="af076-135">After the build, you restart the server.</span></span> <span data-ttu-id="af076-136">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="af076-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="af076-137">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="af076-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="af076-138">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="af076-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="af076-139">Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="af076-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="af076-140">Если по той или иной причине на открытом листе нет таблицы, нажмите в области задач кнопку **Create Table** (Создать таблицу).</span><span class="sxs-lookup"><span data-stu-id="af076-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table**.</span></span>
6. <span data-ttu-id="af076-141">Нажмите кнопки **Filter Table** (Фильтровать таблицу) и **Sort Table** (Сортировать таблицу) в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="af076-141">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Учебник Excel | Фильтрация и сортировка таблицы](../images/excel-tutorial-filter-and-sort-table.png)

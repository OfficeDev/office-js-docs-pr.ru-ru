<span data-ttu-id="97487-101">На этом этапе руководства мы создадим диаграмму, используя данные из ранее созданной таблицы, а затем отформатируем эту диаграмму.</span><span class="sxs-lookup"><span data-stu-id="97487-101">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

> [!NOTE]
> <span data-ttu-id="97487-102">Это один из разделов руководства по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="97487-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="97487-103">Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Excel](../tutorials/excel-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="97487-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="chart-table-data"></a><span data-ttu-id="97487-104">Табличные данные диаграммы</span><span class="sxs-lookup"><span data-stu-id="97487-104">Chart table data</span></span>

1. <span data-ttu-id="97487-105">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="97487-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="97487-106">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="97487-106">Open the file index.html.</span></span>
3. <span data-ttu-id="97487-107">Под элементом `div`, содержащим кнопку `sort-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="97487-107">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-chart">Create Chart</button>            
    </div>
    ```

4. <span data-ttu-id="97487-108">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="97487-108">Open the app.js file.</span></span>

5. <span data-ttu-id="97487-109">Под строкой, назначающей обработчик нажатия кнопки `sort-chart`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="97487-109">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="97487-110">Под функцией `sortTable` добавьте приведенную ниже функцию.</span><span class="sxs-lookup"><span data-stu-id="97487-110">Below the `sortTable` function add the following function.</span></span>

    ```js
    function createChart() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to get the range of data to be charted.

            // TODO2: Queue command to create the chart and define its type.

            // TODO3: Queue commands to position and format the chart.

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

7. <span data-ttu-id="97487-p102">Замените `TODO1` приведенным ниже кодом. Обратите внимание на то, что для исключения строки заголовков в коде вместо метода `getRange` используется метод `Table.getDataBodyRange`, чтобы получить нужный диапазон данных для диаграммы.</span><span class="sxs-lookup"><span data-stu-id="97487-p102">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const dataRange = expensesTable.getDataBodyRange();
    ``` 

8. <span data-ttu-id="97487-113">Замените `TODO2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="97487-113">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="97487-114">Обратите внимание на следующие параметры:</span><span class="sxs-lookup"><span data-stu-id="97487-114">Note the following parameters:</span></span>
   - <span data-ttu-id="97487-p104">Первый параметр метода `add` задает тип диаграммы. Существует несколько десятков типов.</span><span class="sxs-lookup"><span data-stu-id="97487-p104">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span> 
   - <span data-ttu-id="97487-117">Второй параметр задает диапазон данных, включаемых в диаграмму.</span><span class="sxs-lookup"><span data-stu-id="97487-117">The second parameter specifies the range of data to include in the chart.</span></span> 
   - <span data-ttu-id="97487-118">Третий параметр определяет, как следует отображать на диаграмме ряд точек данных из таблицы: по строкам или по столбцам.</span><span class="sxs-lookup"><span data-stu-id="97487-118">The third parameter determines whether a series of data points from the table should be charted rowwise or columnwise.</span></span> <span data-ttu-id="97487-119">Значение `auto` сообщает Excel, что следует выбрать оптимальный способ.</span><span class="sxs-lookup"><span data-stu-id="97487-119">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    let chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ``` 

9. <span data-ttu-id="97487-120">Замените `TODO3` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="97487-120">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="97487-121">Большая часть этого кода не требует объяснений.</span><span class="sxs-lookup"><span data-stu-id="97487-121">Most of this code is self-explanatory.</span></span> <span data-ttu-id="97487-122">Примечание.</span><span class="sxs-lookup"><span data-stu-id="97487-122">Note:</span></span>
   - <span data-ttu-id="97487-123">Параметры метода `setPosition` задают левую верхнюю и правую нижнюю ячейки области листа, которые должны содержать диаграмму.</span><span class="sxs-lookup"><span data-stu-id="97487-123">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="97487-124">Excel может настраивать такие параметры, как ширина линий, чтобы диаграмма хорошо выглядела в выделенном для нее пространстве.</span><span class="sxs-lookup"><span data-stu-id="97487-124">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   - <span data-ttu-id="97487-125">"Ряд" — это набор точек данных из столбца таблицы.</span><span class="sxs-lookup"><span data-stu-id="97487-125">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="97487-126">Так как в таблице есть только один нестроковый столбец, Excel делает вывод, что это единственный столбец точек данных для диаграммы.</span><span class="sxs-lookup"><span data-stu-id="97487-126">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="97487-127">Он рассматривает другие столбцы как метки диаграммы.</span><span class="sxs-lookup"><span data-stu-id="97487-127">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="97487-128">Следовательно, в диаграмме будет только один ряд, обозначенный индексом 0.</span><span class="sxs-lookup"><span data-stu-id="97487-128">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="97487-129">К нему следует добавить метку "Значение в €".</span><span class="sxs-lookup"><span data-stu-id="97487-129">This is the one to label with "Value in €".</span></span> 

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="97487-130">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="97487-130">Test the add-in</span></span>


1. <span data-ttu-id="97487-131">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="97487-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="97487-132">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="97487-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="97487-133">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="97487-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="97487-134">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="97487-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="97487-135">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="97487-135">After the build, you restart the server.</span></span> <span data-ttu-id="97487-136">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="97487-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="97487-137">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="97487-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="97487-138">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="97487-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="97487-139">Повторно загрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="97487-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="97487-140">Если по той или иной причине на открытом листе нет таблицы, нажмите в области задач кнопку **Create Table** (Создать таблицу), а затем — кнопки **Filter Table** (Фильтровать таблицу) и **Sort Table** (Сортировать таблицу) в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="97487-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>
6. <span data-ttu-id="97487-141">Нажмите кнопку **Create Chart** (Создать диаграмму).</span><span class="sxs-lookup"><span data-stu-id="97487-141">Choose the **Create Chart** button.</span></span> <span data-ttu-id="97487-142">Будет создана диаграмма, включающая только данные из отфильтрованных строк.</span><span class="sxs-lookup"><span data-stu-id="97487-142">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="97487-143">Метки точек данных в нижней части диаграммы отсортированы согласно заданному для нее порядку, то есть по именам продавцов в обратном алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="97487-143">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Руководство по Excel: создание диаграммы](../images/excel-tutorial-create-chart.png)

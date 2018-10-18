<span data-ttu-id="2b00f-101">На этом этапе руководства мы проверим программным способом, поддерживает ли надстройка текущую версию Excel, установленную у пользователя, а также добавим таблицу на лист, заполним ее данными и отформатируем.</span><span class="sxs-lookup"><span data-stu-id="2b00f-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

> [!NOTE]
> <span data-ttu-id="2b00f-102">Это один из разделов руководства по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="2b00f-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="2b00f-103">Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Excel](../tutorials/excel-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="2b00f-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="2b00f-104">Написание кода надстройки</span><span class="sxs-lookup"><span data-stu-id="2b00f-104">Code the add-in</span></span>

1. <span data-ttu-id="2b00f-105">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="2b00f-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="2b00f-106">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2b00f-106">Open the file index.html.</span></span>
3. <span data-ttu-id="2b00f-107">Замените `TODO1` на следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="2b00f-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="2b00f-108">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2b00f-108">Open the app.js file.</span></span>
5. <span data-ttu-id="2b00f-109">Замените `TODO1` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="2b00f-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="2b00f-110">Этот код определяет, поддерживает ли установленная у пользователя версия Excel ту версию файла Excel.js, которая включает все API, используемые в этой серии руководств.</span><span class="sxs-lookup"><span data-stu-id="2b00f-110">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="2b00f-111">В рабочей надстройке можно использовать текст условного блока, чтобы скрыть или отключить пользовательский интерфейс, где вызываются неподдерживаемые API.</span><span class="sxs-lookup"><span data-stu-id="2b00f-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="2b00f-112">При этом пользователь по-прежнему сможет использовать те части надстройки, которые поддерживаются в его версии Excel.</span><span class="sxs-lookup"><span data-stu-id="2b00f-112">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    } 
    ```

6. <span data-ttu-id="2b00f-113">Замените `TODO2` на следующий код:</span><span class="sxs-lookup"><span data-stu-id="2b00f-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="2b00f-114">Замените `TODO3` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="2b00f-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="2b00f-115">Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="2b00f-115">Note the following:</span></span>
   - <span data-ttu-id="2b00f-116">Бизнес-логика Excel.js будет добавлена в функцию, передаваемую методу `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="2b00f-116">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="2b00f-117">Эта логика выполняется не сразу.</span><span class="sxs-lookup"><span data-stu-id="2b00f-117">This logic does not execute immediately.</span></span> <span data-ttu-id="2b00f-118">Вместо этого она добавляется в очередь ожидания команд.</span><span class="sxs-lookup"><span data-stu-id="2b00f-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="2b00f-119">Метод `context.sync` отправляет все команды из очереди в Excel для выполнения.</span><span class="sxs-lookup"><span data-stu-id="2b00f-119">The `context.sync` method sends all queued commands to Excel for execution.</span></span>
   - <span data-ttu-id="2b00f-120">За методом `Excel.run` следует блок `catch`.</span><span class="sxs-lookup"><span data-stu-id="2b00f-120">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="2b00f-121">Рекомендуется всегда следовать этой методике.</span><span class="sxs-lookup"><span data-stu-id="2b00f-121">This is a best practice that you should always follow.</span></span> 

    ```js
    function createTable() {
        Excel.run(function (context) {
            
            // TODO4: Queue table creation logic here.

            // TODO5: Queue commands to populate the table with data.

            // TODO6: Queue commands to format the table.

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

8. <span data-ttu-id="2b00f-p106">Замените `TODO4` на приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="2b00f-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="2b00f-124">код создает таблицу с помощью метода `add` коллекции таблиц на листе, которая всегда существует, даже если она пуста.</span><span class="sxs-lookup"><span data-stu-id="2b00f-124">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="2b00f-125">Это стандартный способ создания объектов Excel.js.</span><span class="sxs-lookup"><span data-stu-id="2b00f-125">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="2b00f-126">API конструкторов классов не существуют, а для создания объекта Excel никогда не следует использовать оператор `new`.</span><span class="sxs-lookup"><span data-stu-id="2b00f-126">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="2b00f-127">Вместо этого следует добавить его к объекту родительской коллекции.</span><span class="sxs-lookup"><span data-stu-id="2b00f-127">Instead, you add to a parent collection object.</span></span> 
   - <span data-ttu-id="2b00f-128">Первый параметр метода `add` — это диапазон, содержащий только первую строку, а не весь диапазон таблицы, который мы в конечном итоге будем использовать.</span><span class="sxs-lookup"><span data-stu-id="2b00f-128">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="2b00f-129">Это связано с тем, что при заполнении строк данных (на следующем этапе) надстройка добавляет к таблице новые строки, а не записывает их в ячейки имеющихся строк.</span><span class="sxs-lookup"><span data-stu-id="2b00f-129">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="2b00f-130">Такой шаблон более распространен, так как количество строк в таблице часто неизвестно на момент ее создания.</span><span class="sxs-lookup"><span data-stu-id="2b00f-130">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span> 
   - <span data-ttu-id="2b00f-131">Имена таблиц должны быть уникальными в рамках всей книги, а не только одного листа.</span><span class="sxs-lookup"><span data-stu-id="2b00f-131">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ``` 

9. <span data-ttu-id="2b00f-p109">Замените `TODO5` на приведенный ниже код. Примечание:</span><span class="sxs-lookup"><span data-stu-id="2b00f-p109">Replace `TODO5` with the following code. Note:</span></span>
   - <span data-ttu-id="2b00f-134">значения ячеек диапазона задаются с помощью массива массивов.</span><span class="sxs-lookup"><span data-stu-id="2b00f-134">The cell values of a range are set with an array of arrays.</span></span>
   - <span data-ttu-id="2b00f-135">Новые строки создаются в таблице путем вызова метода `add` коллекции ее строк.</span><span class="sxs-lookup"><span data-stu-id="2b00f-135">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="2b00f-136">Вы можете добавить несколько строк в одном вызове метода `add`, включив несколько массивов значений ячеек в родительский массив, передаваемый в качестве второго параметра.</span><span class="sxs-lookup"><span data-stu-id="2b00f-136">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

    ```js
    expensesTable.getHeaderRowRange().values = 
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    ``` 

10. <span data-ttu-id="2b00f-p111">Замените `TODO6` на приведенный ниже код. Примечание:</span><span class="sxs-lookup"><span data-stu-id="2b00f-p111">Replace `TODO6` with the following code. Note:</span></span>
   - <span data-ttu-id="2b00f-139">код получает ссылку на столбец **Сумма**, передавая его индекс (с отсчетом от нуля) в метод `getItemAt` коллекции столбцов таблицы.</span><span class="sxs-lookup"><span data-stu-id="2b00f-139">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span> 

     > [!NOTE]
     > <span data-ttu-id="2b00f-140">У объектов коллекций Excel.js (например, `TableCollection`, `WorksheetCollection` и `TableColumnCollection`) есть свойство `items`, представляющее собой массив дочерних типов объектов (например, `Table`, `Worksheet` или `TableColumn`). Однако сам объект `*Collection` не является массивом.</span><span class="sxs-lookup"><span data-stu-id="2b00f-140">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="2b00f-141">Затем код форматирует диапазон столбца **Сумма** как денежные суммы в евро с точностью до второго знака после запятой.</span><span class="sxs-lookup"><span data-stu-id="2b00f-141">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 
   - <span data-ttu-id="2b00f-142">Напоследок он обеспечивает достаточные ширину столбцов и высоту строк для размещения самого длинного (или самого высокого) элемента данных.</span><span class="sxs-lookup"><span data-stu-id="2b00f-142">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="2b00f-143">Обратите внимание, что код должен привести объекты `Range` к нужному формату.</span><span class="sxs-lookup"><span data-stu-id="2b00f-143">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="2b00f-144">`TableColumn` У объектов `TableColumn` и `TableRow` нет свойств формата.</span><span class="sxs-lookup"><span data-stu-id="2b00f-144">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="2b00f-145">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="2b00f-145">Test the add-in</span></span>

1. <span data-ttu-id="2b00f-146">Откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="2b00f-146">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="2b00f-147">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="2b00f-147">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
3. <span data-ttu-id="2b00f-148">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="2b00f-148">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="2b00f-149">Загрузите неопубликованную надстройку одним из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="2b00f-149">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="2b00f-150">Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="2b00f-150">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="2b00f-151">Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="2b00f-151">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="2b00f-152">iPad и Mac: [загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="2b00f-152">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="2b00f-153">В меню **Главная** выберите пункт **Показать область задач**.</span><span class="sxs-lookup"><span data-stu-id="2b00f-153">On the **Home** menu, choose **Show Taskpane**.</span></span>
6. <span data-ttu-id="2b00f-154">В области задач нажмите кнопку **Создать таблицу**.</span><span class="sxs-lookup"><span data-stu-id="2b00f-154">In the taskpane, choose **Create Table**.</span></span>

    ![Руководство по Excel: создание таблицы](../images/excel-tutorial-create-table.png)

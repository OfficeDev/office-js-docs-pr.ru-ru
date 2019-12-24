---
title: Руководство по надстройкам Excel
description: В этом руководстве показана разработка надстройки Excel, которая создает, заполняет, фильтрует и сортирует данные таблиц, создает диаграммы, закрепляет заголовки таблиц, защищает листы и открывает диалоговые окна.
ms.date: 11/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 48f4decc0cadddecd5669b960238ddd3381f0932
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851420"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="a72b6-103">Учебник: Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="a72b6-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="a72b6-104">С помощью данного учебника вы сможете создать надстройку области задач Excel, которая выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="a72b6-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="a72b6-105">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-105">Creates a table</span></span>
> * <span data-ttu-id="a72b6-106">Фильтрация и сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="a72b6-107">Создание графика</span><span class="sxs-lookup"><span data-stu-id="a72b6-107">Creates a chart</span></span>
> * <span data-ttu-id="a72b6-108">Закрепление заголовка таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-108">Freezes a table header</span></span>
> * <span data-ttu-id="a72b6-109">Защита листа</span><span class="sxs-lookup"><span data-stu-id="a72b6-109">Protects a worksheet</span></span>
> * <span data-ttu-id="a72b6-110">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="a72b6-110">Opens a dialog</span></span>

> [!TIP]
> <span data-ttu-id="a72b6-111">Если вы уже выполнили краткое руководство по [надстройке области задач Excel](../quickstarts/excel-quickstart-jquery.md) и хотите использовать этот проект в качестве отправной точки для этого руководства, перейдите непосредственно к разделу [Создание таблицы](#create-a-table) , чтобы запустить это руководство.</span><span class="sxs-lookup"><span data-stu-id="a72b6-111">If you've already completed the [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md) quick start, and want to use that project as a starting point for this tutorial, go directly to the [Create a table](#create-a-table) section to start this tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a72b6-112">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="a72b6-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="a72b6-113">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="a72b6-113">Create your add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="a72b6-114">**Выберите тип проекта:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="a72b6-114">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="a72b6-115">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="a72b6-115">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="a72b6-116">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="a72b6-116">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="a72b6-117">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="a72b6-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Генератор Yeoman](../images/yo-office-excel.png)

<span data-ttu-id="a72b6-119">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="a72b6-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a><span data-ttu-id="a72b6-120">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-120">Create a table</span></span>

<span data-ttu-id="a72b6-121">На этом этапе руководства мы проверим программным способом, поддерживает ли надстройка текущую версию Excel, установленную у пользователя, а также добавим таблицу на лист, заполним ее данными и отформатируем.</span><span class="sxs-lookup"><span data-stu-id="a72b6-121">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="a72b6-122">Написание кода надстройки</span><span class="sxs-lookup"><span data-stu-id="a72b6-122">Code the add-in</span></span>

1. <span data-ttu-id="a72b6-123">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="a72b6-123">Open the project in your code editor.</span></span>

2. <span data-ttu-id="a72b6-124">Откройте файл **./src/TaskPane/TaskPane.HTML**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-124">Open the file **./src/taskpane/taskpane.html**.</span></span>  <span data-ttu-id="a72b6-125">Этот файл содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="a72b6-125">This file contains the HTML markup for the task pane.</span></span>

3. <span data-ttu-id="a72b6-126">Нахождение `<main>` элемента и удаление всех строк, которые отображаются после открывающего `<main>` тега и перед закрывающим `</main>` тегом.</span><span class="sxs-lookup"><span data-stu-id="a72b6-126">Locate the `<main>` element and delete all lines that appear after the opening `<main>` tag and before the closing `</main>` tag.</span></span>

4. <span data-ttu-id="a72b6-127">Добавьте следующую разметку сразу после открывающего `<main>` тега:</span><span class="sxs-lookup"><span data-stu-id="a72b6-127">Add the following markup immediately after the opening `<main>` tag:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

5. <span data-ttu-id="a72b6-128">Откройте файл **./СРК/таскпане/таскпане.ЖС**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-128">Open the file **./src/taskpane/taskpane.js**.</span></span> <span data-ttu-id="a72b6-129">Этот файл содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и ведущим приложением Office.</span><span class="sxs-lookup"><span data-stu-id="a72b6-129">This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

6. <span data-ttu-id="a72b6-130">Удалите все ссылки на `run` кнопку и `run()` функцию, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="a72b6-130">Remove all references to the `run` button and the `run()` function by doing the following:</span></span>

    - <span data-ttu-id="a72b6-131">Откройте и удалите строку `document.getElementById("run").onclick = run;`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-131">Locate and delete the line `document.getElementById("run").onclick = run;`.</span></span>

    - <span data-ttu-id="a72b6-132">Искать и удалить функцию целиком `run()` .</span><span class="sxs-lookup"><span data-stu-id="a72b6-132">Locate and delete the entire `run()` function.</span></span>

7. <span data-ttu-id="a72b6-133">В вызове `Office.onReady` метода откройте строку `if (info.host === Office.HostType.Excel) {` и добавьте следующий код сразу после этой строки.</span><span class="sxs-lookup"><span data-stu-id="a72b6-133">Within the `Office.onReady` method call, locate the line `if (info.host === Office.HostType.Excel) {` and add the following code immediately after that line.</span></span> <span data-ttu-id="a72b6-134">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-134">Note:</span></span>

    - <span data-ttu-id="a72b6-135">Первая часть этого кода определяет, поддерживает ли версия Excel для пользователя версию файла Excel. js, которая включает все API, которые будет использовать эта серия учебных курсов.</span><span class="sxs-lookup"><span data-stu-id="a72b6-135">The first part of this code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="a72b6-136">В рабочей надстройке можно использовать текст условного блока, чтобы скрыть или отключить пользовательский интерфейс, где вызываются неподдерживаемые API.</span><span class="sxs-lookup"><span data-stu-id="a72b6-136">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="a72b6-137">При этом пользователь по-прежнему сможет использовать те части надстройки, которые поддерживаются в его версии Excel.</span><span class="sxs-lookup"><span data-stu-id="a72b6-137">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    - <span data-ttu-id="a72b6-138">Во второй части этого кода добавляется обработчик событий для `create-table` кнопки.</span><span class="sxs-lookup"><span data-stu-id="a72b6-138">The second part of this code adds an event handler for the `create-table` button.</span></span>

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

8. <span data-ttu-id="a72b6-139">Добавьте указанную ниже функцию в конец файла.</span><span class="sxs-lookup"><span data-stu-id="a72b6-139">Add the following function to the end of the file.</span></span> <span data-ttu-id="a72b6-140">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-140">Note:</span></span>

    - <span data-ttu-id="a72b6-p106">Бизнес-логика Excel.js будет добавлена в функцию, передаваемую методу `Excel.run`. Эта логика выполняется не сразу. Вместо этого она добавляется в очередь ожидания команд.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p106">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

    - <span data-ttu-id="a72b6-144">Метод `context.sync` отправляет все команды из очереди в Excel для выполнения.</span><span class="sxs-lookup"><span data-stu-id="a72b6-144">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

    - <span data-ttu-id="a72b6-p107">За методом `Excel.run` следует блок `catch`. Рекомендуется всегда следовать этой методике.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p107">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

    ```js
    function createTable() {
        Excel.run(function (context) {

            // TODO1: Queue table creation logic here.

            // TODO2: Queue commands to populate the table with data.

            // TODO3: Queue commands to format the table.

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

9. <span data-ttu-id="a72b6-147">Замените `TODO1` в `createTable()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-147">Within the `createTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="a72b6-148">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-148">Note:</span></span>

    - <span data-ttu-id="a72b6-p109">код создает таблицу с помощью метода `add` коллекции таблиц на листе, которая всегда существует, даже если она пуста. Это стандартный способ создания объектов Excel.js. API конструкторов классов не существуют, а для создания объекта Excel никогда не следует использовать оператор `new`. Вместо этого следует добавить его к объекту родительской коллекции.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p109">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

    - <span data-ttu-id="a72b6-p110">Первый параметр метода `add`— это диапазон, содержащий только первую строку, а не весь диапазон таблицы, который мы в конечном итоге будем использовать. Это связано с тем, что при заполнении строк данных (на следующем этапе) надстройка добавляет к таблице новые строки, а не записывает их в ячейки имеющихся строк. Такой шаблон более распространен, так как количество строк в таблице часто неизвестно на момент ее создания.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p110">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

    - <span data-ttu-id="a72b6-156">Имена таблиц должны быть уникальными в рамках всей книги, а не только одного листа.</span><span class="sxs-lookup"><span data-stu-id="a72b6-156">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

10. <span data-ttu-id="a72b6-157">Замените `TODO2` в `createTable()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-157">Within the `createTable()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="a72b6-158">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-158">Note:</span></span>

    - <span data-ttu-id="a72b6-159">значения ячеек диапазона задаются с помощью массива массивов.</span><span class="sxs-lookup"><span data-stu-id="a72b6-159">The cell values of a range are set with an array of arrays.</span></span>

    - <span data-ttu-id="a72b6-p112">Новые строки создаются в таблице путем вызова метода `add` коллекции ее строк. Вы можете добавить несколько строк в одном вызове метода `add`, включив несколько массивов значений ячеек в родительский массив, передаваемый в качестве второго параметра.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p112">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

11. <span data-ttu-id="a72b6-162">Замените `TODO3` в `createTable()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-162">Within the `createTable()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="a72b6-163">Примечание:</span><span class="sxs-lookup"><span data-stu-id="a72b6-163">Note:</span></span>

    - <span data-ttu-id="a72b6-164">код получает ссылку на столбец **Сумма**, передавая его индекс (с отсчетом от нуля) в метод `getItemAt` коллекции столбцов таблицы.</span><span class="sxs-lookup"><span data-stu-id="a72b6-164">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

        > [!NOTE]
        > <span data-ttu-id="a72b6-165">У объектов коллекций Excel.js (например, `TableCollection`, `WorksheetCollection` и `TableColumnCollection`) есть свойство `items`, представляющее собой массив дочерних типов объектов (например, `Table`, `Worksheet` или `TableColumn`). Однако сам объект `*Collection` не является массивом.</span><span class="sxs-lookup"><span data-stu-id="a72b6-165">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

    - <span data-ttu-id="a72b6-166">Затем код форматирует диапазон столбца **Сумма** как денежные суммы в евро с точностью до второго знака после запятой.</span><span class="sxs-lookup"><span data-stu-id="a72b6-166">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

    - <span data-ttu-id="a72b6-p114">Напоследок он обеспечивает достаточные ширину столбцов и высоту строк для размещения самого длинного (или самого высокого) элемента данных. Обратите внимание, что код должен привести объекты `Range` к нужному формату. У объектов `TableColumn` и `TableRow` нет свойств формата.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p114">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

12. <span data-ttu-id="a72b6-170">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="a72b6-170">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="a72b6-171">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="a72b6-171">Test the add-in</span></span>

1. <span data-ttu-id="a72b6-172">Выполните указанные ниже действия, чтобы запустить локальный веб-сервер и загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="a72b6-172">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a72b6-173">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="a72b6-173">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="a72b6-174">Если вам будет предложено установить сертификат после того, как вы запустите одну из указанных ниже команд, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="a72b6-174">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="a72b6-175">Если вы тестируете надстройку на Mac, выполните следующую команду в корневом каталоге проекта, прежде чем продолжить.</span><span class="sxs-lookup"><span data-stu-id="a72b6-175">If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding.</span></span> <span data-ttu-id="a72b6-176">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="a72b6-176">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="a72b6-177">Чтобы протестировать надстройку в Excel, выполните следующую команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="a72b6-177">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="a72b6-178">При этом запустится локальный веб-сервер (если он еще не запущен) и откроется Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="a72b6-178">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="a72b6-179">Чтобы протестировать надстройку в Excel в Интернете, выполните следующую команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="a72b6-179">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="a72b6-180">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="a72b6-180">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="a72b6-181">Чтобы использовать надстройку, откройте новый документ в Excel в Интернете и затем Загрузка неопубликованных свою надстройку, следуя инструкциям в статье [Загрузка неопубликованных Office Add-ins in Office in Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="a72b6-181">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

2. <span data-ttu-id="a72b6-182">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="a72b6-182">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

3. <span data-ttu-id="a72b6-184">В области задач нажмите кнопку **CREATE TABLE (создать таблицу** ).</span><span class="sxs-lookup"><span data-stu-id="a72b6-184">In the task pane, choose the **Create Table** button.</span></span>

    ![Руководство по Excel: создание таблицы](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="a72b6-186">Фильтрация и сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-186">Filter and sort a table</span></span>

<span data-ttu-id="a72b6-187">Из этого раздела руководства вы узнаете, как отфильтровать и отсортировать созданную ранее таблицу.</span><span class="sxs-lookup"><span data-stu-id="a72b6-187">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="a72b6-188">Фильтрация таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-188">Filter the table</span></span>

1. <span data-ttu-id="a72b6-189">Откройте файл **./src/TaskPane/TaskPane.HTML**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-189">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="a72b6-190">Нахождение `<button>` элемента для `create-table` кнопки и добавление приведенной ниже разметки после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-190">Locate the `<button>` element for the `create-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="filter-table">Filter Table</button><br/><br/>
    ```

3. <span data-ttu-id="a72b6-191">Откройте файл **./СРК/таскпане/таскпане.ЖС**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-191">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="a72b6-192">В вызове `Office.onReady` метода укажите строку, которая назначает обработчик нажатия `create-table` кнопки, и добавьте следующий код после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-192">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("filter-table").onclick = filterTable;
    ```

5. <span data-ttu-id="a72b6-193">Добавьте указанную ниже функцию в конец файла:</span><span class="sxs-lookup"><span data-stu-id="a72b6-193">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="a72b6-194">Замените `TODO1` в `filterTable()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-194">Within the `filterTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="a72b6-195">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-195">Note:</span></span>

   - <span data-ttu-id="a72b6-p120">Код получает ссылку на столбец, который нужно отфильтровать, передавая имя столбца методу `getItem`, а не передавая его индекс методу `getItemAt`, как это делает метод `createTable`. Так как пользователи могут перемещать столбцы, по заданному индексу может располагаться уже другой столбец. Следовательно, для получения ссылки безопаснее использовать имя столбца. Мы спокойно использовали метод `getItemAt` в предыдущем разделе, потому что мы использовали его в методе, который создает таблицу, и пользователь никак не мог переместить столбец.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p120">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="a72b6-200">Метод `applyValuesFilter` является одним из нескольких методов фильтрации объекта `Filter`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-200">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="a72b6-201">Сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-201">Sort the table</span></span>

1. <span data-ttu-id="a72b6-202">Откройте файл **./src/TaskPane/TaskPane.HTML**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-202">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="a72b6-203">Нахождение `<button>` элемента для `filter-table` кнопки и добавление приведенной ниже разметки после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-203">Locate the `<button>` element for the `filter-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

3. <span data-ttu-id="a72b6-204">Откройте файл **./СРК/таскпане/таскпане.ЖС**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-204">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="a72b6-205">В вызове `Office.onReady` метода укажите строку, которая назначает обработчик нажатия `filter-table` кнопки, и добавьте следующий код после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-205">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `filter-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

5. <span data-ttu-id="a72b6-206">Добавьте указанную ниже функцию в конец файла:</span><span class="sxs-lookup"><span data-stu-id="a72b6-206">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="a72b6-207">Замените `TODO1` в `sortTable()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-207">Within the `sortTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="a72b6-208">Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="a72b6-208">Note:</span></span>

   - <span data-ttu-id="a72b6-209">Код создает массив объектов `SortField`, состоящий из одного элемента, так как надстройка сортирует таблицу только по столбцу Merchant.</span><span class="sxs-lookup"><span data-stu-id="a72b6-209">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="a72b6-210">Свойство `key` объекта `SortField` — это отсчитываемый от нуля индекс столбца, по которому необходимо сортировать таблицу.</span><span class="sxs-lookup"><span data-stu-id="a72b6-210">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="a72b6-211">Элемент `sort` объекта `Table` — это объект `TableSort`, а не метод.</span><span class="sxs-lookup"><span data-stu-id="a72b6-211">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="a72b6-212">Объекты `SortField` передаются методу `apply` объекта `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-212">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ```

7. <span data-ttu-id="a72b6-213">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="a72b6-213">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="a72b6-214">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="a72b6-214">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="a72b6-215">Если область задач надстройки еще не открыта в Excel, перейдите на вкладку **Главная** и нажмите кнопку **Показать** область задач на ленте, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="a72b6-215">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="a72b6-216">Если таблица, добавленная ранее в этом руководстве, отсутствует на открытом листе, нажмите кнопку **CREATE TABLE (создать таблицу** ) в области задач.</span><span class="sxs-lookup"><span data-stu-id="a72b6-216">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button in the task pane.</span></span>

4. <span data-ttu-id="a72b6-217">Нажмите кнопку **Фильтрация таблицы** и кнопку **Сортировка таблицы** в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="a72b6-217">Choose the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

    ![Учебник Excel - Фильтрация и сортировка таблицы](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart"></a><span data-ttu-id="a72b6-219">Создание диаграммы</span><span class="sxs-lookup"><span data-stu-id="a72b6-219">Create a chart</span></span>

<span data-ttu-id="a72b6-220">На этом этапе руководства мы создадим диаграмму, используя данные из ранее созданной таблицы, а затем отформатируем эту диаграмму.</span><span class="sxs-lookup"><span data-stu-id="a72b6-220">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="a72b6-221">Создание диаграммы с помощью таблицы данных</span><span class="sxs-lookup"><span data-stu-id="a72b6-221">Chart a chart using table data</span></span>

1. <span data-ttu-id="a72b6-222">Откройте файл **./src/TaskPane/TaskPane.HTML**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-222">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="a72b6-223">Нахождение `<button>` элемента для `sort-table` кнопки и добавление приведенной ниже разметки после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-223">Locate the `<button>` element for the `sort-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="create-chart">Create Chart</button><br/><br/>
    ```

3. <span data-ttu-id="a72b6-224">Откройте файл **./СРК/таскпане/таскпане.ЖС**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-224">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="a72b6-225">В вызове `Office.onReady` метода укажите строку, которая назначает обработчик нажатия `sort-table` кнопки, и добавьте следующий код после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-225">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `sort-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("create-chart").onclick = createChart;
    ```

5. <span data-ttu-id="a72b6-226">Добавьте указанную ниже функцию в конец файла:</span><span class="sxs-lookup"><span data-stu-id="a72b6-226">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="a72b6-227">Замените `TODO1` в `createChart()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-227">Within the `createChart()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="a72b6-228">Обратите внимание на то, что для исключения строки заголовков в коде вместо метода `Table.getDataBodyRange` используется метод `getRange`, чтобы получить нужный диапазон данных для диаграммы.</span><span class="sxs-lookup"><span data-stu-id="a72b6-228">Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

7. <span data-ttu-id="a72b6-229">Замените `TODO2` в `createChart()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-229">Within the `createChart()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="a72b6-230">Обратите внимание на следующие параметры:</span><span class="sxs-lookup"><span data-stu-id="a72b6-230">Note the following parameters:</span></span>

   - <span data-ttu-id="a72b6-p125">Первый параметр метода `add` задает тип диаграммы. Существует несколько десятков типов.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p125">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="a72b6-233">Второй параметр задает диапазон данных, включаемых в диаграмму.</span><span class="sxs-lookup"><span data-stu-id="a72b6-233">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="a72b6-234">Третий параметр определяет, как следует отображать на диаграмме ряд точек данных из таблицы: по строкам или по столбцам.</span><span class="sxs-lookup"><span data-stu-id="a72b6-234">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="a72b6-235">Значение `auto` сообщает Excel, что следует выбрать оптимальный способ.</span><span class="sxs-lookup"><span data-stu-id="a72b6-235">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

8. <span data-ttu-id="a72b6-236">Замените `TODO3` в `createChart()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-236">Within the `createChart()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="a72b6-237">Большая часть этого кода не требует объяснений.</span><span class="sxs-lookup"><span data-stu-id="a72b6-237">Most of this code is self-explanatory.</span></span> <span data-ttu-id="a72b6-238">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-238">Note:</span></span>
   
   - <span data-ttu-id="a72b6-p128">Параметры метода `setPosition` задают левую верхнюю и правую нижнюю ячейки области листа, которые должны содержать диаграмму. Excel может настраивать такие параметры, как ширина линий, чтобы диаграмма хорошо выглядела в выделенном для нее пространстве.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p128">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="a72b6-p129">"Ряд" — это набор точек данных из столбца таблицы. Так как в таблице есть только один нестроковый столбец, Excel делает вывод, что это единственный столбец точек данных для диаграммы. Он рассматривает другие столбцы как метки диаграммы. Следовательно, в диаграмме будет только один ряд, обозначенный индексом 0. К нему следует добавить метку "Значение в €".</span><span class="sxs-lookup"><span data-stu-id="a72b6-p129">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

9. <span data-ttu-id="a72b6-246">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="a72b6-246">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="a72b6-247">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="a72b6-247">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="a72b6-248">Если область задач надстройки еще не открыта в Excel, перейдите на вкладку **Главная** и нажмите кнопку **Показать** область задач на ленте, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="a72b6-248">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="a72b6-249">Если таблица, добавленная ранее в этом руководстве, отсутствует на открытом листе, нажмите кнопку **CREATE TABLE (создать таблицу** ), а затем кнопку **таблицы фильтра** и кнопку **Сортировка таблицы** в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="a72b6-249">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button, and then the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

4. <span data-ttu-id="a72b6-p130">Нажмите кнопку **Create Chart** (Создать диаграмму). Будет создана диаграмма, включающая только данные из отфильтрованных строк. Метки точек данных в нижней части диаграммы отсортированы согласно заданному для нее порядку, то есть по именам продавцов в обратном алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p130">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Руководство по Excel - Создание диаграммы](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="a72b6-254">Закрепление заголовка таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-254">Freeze a table header</span></span>

<span data-ttu-id="a72b6-p131">Когда таблица достаточно длинная, при прокрутке строка заголовков может исчезать с экрана. В этом разделе учебника мы расскажем, как закрепить строку заголовков созданной ранее таблицы, чтобы она была видна, даже когда пользователь прокручивает лист.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p131">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="a72b6-257">Закрепление строки заголовков таблицы</span><span class="sxs-lookup"><span data-stu-id="a72b6-257">Freeze the table's header row</span></span>

1. <span data-ttu-id="a72b6-258">Откройте файл **./src/TaskPane/TaskPane.HTML**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-258">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="a72b6-259">Нахождение `<button>` элемента для `create-chart` кнопки и добавление приведенной ниже разметки после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-259">Locate the `<button>` element for the `create-chart` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="freeze-header">Freeze Header</button><br/><br/>
    ```

3. <span data-ttu-id="a72b6-260">Откройте файл **./СРК/таскпане/таскпане.ЖС**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-260">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="a72b6-261">В вызове `Office.onReady` метода укажите строку, которая назначает обработчик нажатия `create-chart` кнопки, и добавьте следующий код после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-261">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-chart` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("freeze-header").onclick = freezeHeader;
    ```

5. <span data-ttu-id="a72b6-262">Добавьте указанную ниже функцию в конец файла:</span><span class="sxs-lookup"><span data-stu-id="a72b6-262">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="a72b6-263">Замените `TODO1` в `freezeHeader()` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-263">Within the `freezeHeader()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="a72b6-264">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-264">Note:</span></span>

   - <span data-ttu-id="a72b6-265">Коллекция `Worksheet.freezePanes` — это набор закрепленных строк, которые не исчезают с экрана при прокрутке листа.</span><span class="sxs-lookup"><span data-stu-id="a72b6-265">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="a72b6-p133">Метод `freezeRows` принимает в качестве параметра количество строк сверху, которые необходимо закрепить. Мы передаем значение `1`, чтобы закрепить первую строку.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p133">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

7. <span data-ttu-id="a72b6-268">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="a72b6-268">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="a72b6-269">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="a72b6-269">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="a72b6-270">Если область задач надстройки еще не открыта в Excel, перейдите на вкладку **Главная** и нажмите кнопку **Показать** область задач на ленте, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="a72b6-270">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="a72b6-271">Если таблица, добавленная ранее в этом руководстве, присутствует на листе, удалите ее.</span><span class="sxs-lookup"><span data-stu-id="a72b6-271">If the table you added previously in this tutorial is present in the worksheet, delete it.</span></span>

4. <span data-ttu-id="a72b6-272">В области задач нажмите кнопку **CREATE TABLE (создать таблицу** ).</span><span class="sxs-lookup"><span data-stu-id="a72b6-272">In the task pane, choose the **Create Table** button.</span></span>

5. <span data-ttu-id="a72b6-273">В области задач нажмите кнопку **Закрепить заголовок** .</span><span class="sxs-lookup"><span data-stu-id="a72b6-273">In the task pane, choose the **Freeze Header** button.</span></span>

6. <span data-ttu-id="a72b6-274">Прокрутите лист вниз достаточно для того, чтобы увидеть, что заголовок таблицы остается видимым в верхней части, даже когда верхние строки прокручивается.</span><span class="sxs-lookup"><span data-stu-id="a72b6-274">Scroll down the worksheet far enough to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Учебник Excel - Закрепление заголовка](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="a72b6-276">Защита листа</span><span class="sxs-lookup"><span data-stu-id="a72b6-276">Protect a worksheet</span></span>

<span data-ttu-id="a72b6-277">На данном этапе, описанном в руководстве, вы добавите на ленту еще одну кнопку, при нажатии которой будет выполнена определенная вами функция включения или выключения защиты листа.</span><span class="sxs-lookup"><span data-stu-id="a72b6-277">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="a72b6-278">Настройка манифеста для добавления второй кнопки на ленту</span><span class="sxs-lookup"><span data-stu-id="a72b6-278">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="a72b6-279">Откройте файл манифеста **./манифест.ксмл**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-279">Open the manifest file **./manifest.xml**.</span></span>

2. <span data-ttu-id="a72b6-280">Нахождение `<Control>` элемента.</span><span class="sxs-lookup"><span data-stu-id="a72b6-280">Locate the `<Control>` element.</span></span> <span data-ttu-id="a72b6-281">Этот элемент определяет кнопку **Show Taskpane** (Показать область задач) на вкладке **Главная**, которую вы используете для запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="a72b6-281">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="a72b6-282">Мы добавим вторую кнопку в эту же группу на ленте **Главная**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-282">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="a72b6-283">Добавьте приведенный ниже код между закрывающим тегом элемента управления (`</Control>`) и закрывающим тегом группы (`</Group>`).</span><span class="sxs-lookup"><span data-stu-id="a72b6-283">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="a72b6-284">В XML-файле, который вы только что добавили в файл `TODO1` манифеста, замените на строку, которая дает кнопке идентификатор, уникальный в этом файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a72b6-284">Within the XML you just added to the manifest file, replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="a72b6-285">Так как кнопка будет включать и выключать защиту листа, укажите "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="a72b6-285">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="a72b6-286">Когда вы закончите, открывающий тег `Control` элемента должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="a72b6-286">When you are done, the opening tag for the `Control` element should look like this:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="a72b6-287">Следующие три элемента `TODO` устанавливают "resid", или идентификаторы ресурса.</span><span class="sxs-lookup"><span data-stu-id="a72b6-287">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="a72b6-288">Ресурс должен быть строкой, и вы создадите эти три строки на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="a72b6-288">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="a72b6-289">Сейчас вам нужно присвоить идентификаторы ресурсам.</span><span class="sxs-lookup"><span data-stu-id="a72b6-289">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="a72b6-290">Подпись кнопки должна иметь значение "Toggle Protection", но *идентификатор* этой строки должен быть "ProtectionButtonLabel", поэтому `Label` элемент должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="a72b6-290">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the `Label` element should look like this:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="a72b6-291">Элемент `SuperTip` определяет подсказку для кнопки.</span><span class="sxs-lookup"><span data-stu-id="a72b6-291">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="a72b6-292">Заголовок этой подсказки должен совпадать с названием кнопки, поэтому мы используем тот же ИД ресурса — "ProtectionButtonLabel".</span><span class="sxs-lookup"><span data-stu-id="a72b6-292">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="a72b6-293">Описание подсказки будет следующим: "Click to turn protection of the worksheet on and off" (Нажмите для включения или выключения защиты листа).</span><span class="sxs-lookup"><span data-stu-id="a72b6-293">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="a72b6-294">У `ID` должно быть значение "ProtectionButtonToolTip".</span><span class="sxs-lookup"><span data-stu-id="a72b6-294">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="a72b6-295">Таким образом, при завершении `SuperTip` элемент должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="a72b6-295">So, when you are done, the `SuperTip` element should look like this:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="a72b6-p138">В рабочей надстройке не нужно использовать один и тот же значок для двух разных кнопок, но сейчас мы предлагаем сделать это для простоты. Поэтому код `Icon` в новом теге `Control` представляет собой лишь копию элемента `Icon` из существующего тега `Control`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p138">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="a72b6-298">Для элемента `Action` в исходном элементе `Control`, уже присутствующем в манифесте, задан тип `ShowTaskpane`, но новая кнопка будет не открывать область задач, а выполнять специальную функцию, которую вы создадите позже.</span><span class="sxs-lookup"><span data-stu-id="a72b6-298">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="a72b6-299">Поэтому замените `TODO5` на `ExecuteFunction`(тип действия для кнопок, запускающих специальные функции).</span><span class="sxs-lookup"><span data-stu-id="a72b6-299">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="a72b6-300">Открывающий тег `Action` элемента должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="a72b6-300">The opening tag for the `Action` element should look like this:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="a72b6-p140">У исходного элемента `Action` есть дочерние элементы, определяющие идентификатор области задач и URL-адрес страницы, которая должна быть открыта в области задач. Но у элемента `Action` типа `ExecuteFunction` есть один дочерний элемент, который именует функцию, выполняемую элементом управления. На более позднем этапе вы создадите функцию `toggleProtection`. Поэтому замените `TODO6` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="a72b6-p140">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="a72b6-305">Теперь весь код `Control` должен выглядеть вот так:</span><span class="sxs-lookup"><span data-stu-id="a72b6-305">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="a72b6-306">Прокрутите страницу вниз до раздела `Resources` манифеста.</span><span class="sxs-lookup"><span data-stu-id="a72b6-306">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="a72b6-307">Добавьте приведенный ниже код в качестве дочернего элемента `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-307">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="a72b6-308">Добавьте приведенный ниже код в качестве дочернего элемента `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-308">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="a72b6-309">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="a72b6-309">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="a72b6-310">Создание функции защиты листа</span><span class="sxs-lookup"><span data-stu-id="a72b6-310">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="a72b6-311">Откройте файл **.\коммандс\коммандс.ЖС**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-311">Open the file **.\commands\commands.js**.</span></span>

2. <span data-ttu-id="a72b6-312">Добавьте указанную ниже функцию сразу после `action` функции.</span><span class="sxs-lookup"><span data-stu-id="a72b6-312">Add the following function immediately after the `action` function.</span></span> <span data-ttu-id="a72b6-313">Обратите внимание, что `args` мы указываем параметр для функции и самой последней строкой вызовов `args.completed`функции.</span><span class="sxs-lookup"><span data-stu-id="a72b6-313">Note that we specify an `args` parameter to the function and the very last line of the function calls `args.completed`.</span></span> <span data-ttu-id="a72b6-314">Это требование для всех команд надстройки типа **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-314">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="a72b6-315">Это сигнализирует ведущему приложению Office о том, что работа функции завершена и пользовательский интерфейс снова может реагировать.</span><span class="sxs-lookup"><span data-stu-id="a72b6-315">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="a72b6-316">Добавьте следующую строку в конец файла:</span><span class="sxs-lookup"><span data-stu-id="a72b6-316">Add the following line to the end of the file:</span></span>

    ```js
    g.toggleProtection = toggleProtection;
    ```

4. <span data-ttu-id="a72b6-317">Замените `TODO1` в `toggleProtection` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-317">Within the `toggleProtection` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="a72b6-318">В этом коде используется свойство защиты объекта листа в стандартном шаблоне переключателя.</span><span class="sxs-lookup"><span data-stu-id="a72b6-318">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="a72b6-319">Объяснение `TODO2` будет приведено в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="a72b6-319">The `TODO2` will be explained in the next section.</span></span>

    ```js
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

    if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="a72b6-320">Добавление кода для получения свойств документа в объекты скрипта области задач</span><span class="sxs-lookup"><span data-stu-id="a72b6-320">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="a72b6-321">В каждой функции, созданной в этом руководстве до настоящего момента, вы заставяте в очередь команды для *записи* в документ Office.</span><span class="sxs-lookup"><span data-stu-id="a72b6-321">In each function that you've created in this tutorial until now, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="a72b6-322">Каждая функция заканчивалась вызовом метода `context.sync()`, который отправляет поставленные в очередь команды документу для выполнения.</span><span class="sxs-lookup"><span data-stu-id="a72b6-322">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="a72b6-323">Но код, который вы добавили на последнем этапе, вызывает свойство `sheet.protection.protected`, и в этом заключается существенное отличие от ранее написанных функций, так как `sheet` является лишь объектом прокси, существующим в скрипте вашей области задач.</span><span class="sxs-lookup"><span data-stu-id="a72b6-323">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="a72b6-324">В нем нет сведений о фактическом состоянии защиты документа, поэтому его свойство `protection.protected` не может иметь реального значения.</span><span class="sxs-lookup"><span data-stu-id="a72b6-324">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="a72b6-325">Сначала нужно получить сведения о состоянии защиты от документа и задать значение `sheet.protection.protected`, используя их.</span><span class="sxs-lookup"><span data-stu-id="a72b6-325">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="a72b6-326">Только после этого станет возможным вызов `sheet.protection.protected` без исключения.</span><span class="sxs-lookup"><span data-stu-id="a72b6-326">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="a72b6-327">Процесс получения делится на три этапа:</span><span class="sxs-lookup"><span data-stu-id="a72b6-327">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="a72b6-328">Добавление в очередь команды для загрузки (т. е. получения) свойств, которые должен прочесть ваш код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-328">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="a72b6-329">Вызов метода `sync` объекта контекста, чтобы можно было отправить документу находящуюся в очереди команду для выполнения, а также для возврата запрошенных данных.</span><span class="sxs-lookup"><span data-stu-id="a72b6-329">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="a72b6-330">Метод `sync` асинхронный, поэтому его выполнение должно быть завершено до того, как код вызовет полученные свойства.</span><span class="sxs-lookup"><span data-stu-id="a72b6-330">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="a72b6-331">Эти три действия должны выполняться каждый раз, когда коду нужно *считывать* данные из документа Office.</span><span class="sxs-lookup"><span data-stu-id="a72b6-331">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="a72b6-332">Замените `TODO2` в `toggleProtection` функции приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="a72b6-332">Within the `toggleProtection` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="a72b6-333">Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="a72b6-333">Note:</span></span>
   
   - <span data-ttu-id="a72b6-p145">У каждого объекта Excel есть метод `load`. Вы указываете свойства объекта, которые нужно прочесть в параметре как строку имен, разделенных запятыми. В этом случае нужно прочесть подсвойство свойства `protection`. На подсвойство нужно ссылаться почти так же, как и в остальных частях кода. Отличие заключается в том, что вместо символа "." нужно указать косую черту ("/").</span><span class="sxs-lookup"><span data-stu-id="a72b6-p145">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="a72b6-338">Чтобы логика переключения, которая считывает `sheet.protection.protected`, не срабатывала до выполнения `sync` и присвоения `sheet.protection.protected` правильного значения, полученного из документа, она будет перемещена (на следующем этапе) в функцию `then`, которая не выполняется до завершения `sync`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-338">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```js
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="a72b6-p146">Для двух операторов `return` не может использоваться один путь кода, который не разветвляется, поэтому удалите последнюю строку `return context.sync();` в конце `Excel.run`. Вы добавите новую последнюю строку `context.sync` позже.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p146">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="a72b6-341">Вырежьте структуру `if ... else` в функции `toggleProtection` и вставьте вместо `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-341">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="a72b6-p147">Замените `TODO4` приведенным ниже кодом. Примечание:</span><span class="sxs-lookup"><span data-stu-id="a72b6-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="a72b6-344">Благодаря тому, что метод `sync` передается функции `then`, он не будет запускаться до добавления `sheet.protection.unprotect()` или `sheet.protection.protect()` в очередь.</span><span class="sxs-lookup"><span data-stu-id="a72b6-344">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="a72b6-345">Метод `then` вызывает любую функцию, которая ему передана. Не нужно вызывать `sync` дважды, поэтому уберите "()" после `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-345">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="a72b6-346">Когда все будет готово, функция должна выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="a72b6-346">When you are done, the entire function should look like the following:</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {            
          var sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

5. <span data-ttu-id="a72b6-347">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="a72b6-347">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="a72b6-348">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="a72b6-348">Test the add-in</span></span>

1. <span data-ttu-id="a72b6-349">Закройте все приложения Office, в том числе Excel.</span><span class="sxs-lookup"><span data-stu-id="a72b6-349">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="a72b6-p148">Очистите кэш Office, удалив содержимое папки кэша. Это необходимо, чтобы можно было полностью удалить старую версию надстройки из ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p148">Delete the Office cache by deleting the contents of the cache folder. This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="a72b6-352">Для Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-352">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="a72b6-353">Для Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-353">For Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span> 
    
        > [!NOTE]
        > <span data-ttu-id="a72b6-354">Если эта папка не существует, проверьте наличие следующих папок и, если она найдена, удалите содержимое папки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-354">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
        >    - <span data-ttu-id="a72b6-355">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`где `{host}` находится ведущее приложение Office (например, `Excel`);</span><span class="sxs-lookup"><span data-stu-id="a72b6-355">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
        >    - <span data-ttu-id="a72b6-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`где `{host}` находится ведущее приложение Office (например, `Excel`);</span><span class="sxs-lookup"><span data-stu-id="a72b6-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
        >    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`

3. <span data-ttu-id="a72b6-357">Если локальный веб-сервер уже запущен, остановите его, закрыв окно команд узла.</span><span class="sxs-lookup"><span data-stu-id="a72b6-357">If the local web server is already running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="a72b6-358">Так как файл манифеста был обновлен, необходимо повторно Загрузка неопубликованных надстройку, используя обновленный файл манифеста.</span><span class="sxs-lookup"><span data-stu-id="a72b6-358">Because your manifest file has been updated, you must sideload your add-in again, using the updated manifest file.</span></span> <span data-ttu-id="a72b6-359">Запустите локальный веб-сервер и Загрузка неопубликованных надстройку:</span><span class="sxs-lookup"><span data-stu-id="a72b6-359">Start the local web server and sideload your add-in:</span></span> 

    - <span data-ttu-id="a72b6-360">Чтобы протестировать надстройку в Excel, выполните следующую команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="a72b6-360">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="a72b6-361">При этом запустится локальный веб-сервер (если он еще не запущен) и откроется Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="a72b6-361">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="a72b6-362">Чтобы протестировать надстройку в Excel в Интернете, выполните следующую команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="a72b6-362">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="a72b6-363">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="a72b6-363">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="a72b6-364">Чтобы использовать надстройку, откройте новый документ в Excel в Интернете и затем Загрузка неопубликованных свою надстройку, следуя инструкциям в статье [Загрузка неопубликованных Office Add-ins in Office in Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="a72b6-364">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

5. <span data-ttu-id="a72b6-365">На вкладке **Главная** в Excel нажмите кнопку **включить защиту листа** .</span><span class="sxs-lookup"><span data-stu-id="a72b6-365">On the **Home** tab in Excel, choose the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="a72b6-366">Обратите внимание, что большинство элементов управления на ленте отключены (и отображаются визуально серым цветом), как показано на следующем снимке экрана.</span><span class="sxs-lookup"><span data-stu-id="a72b6-366">Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in the following screenshot.</span></span> 

    ![Руководство по Excel: лента с включенной защитой](../images/excel-tutorial-ribbon-with-protection-on-2.png)

6. <span data-ttu-id="a72b6-368">Выберите ячейку, как если бы вы хотели изменить ее содержимое.</span><span class="sxs-lookup"><span data-stu-id="a72b6-368">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="a72b6-369">Excel отображает сообщение об ошибке, указывающее на то, что лист защищен.</span><span class="sxs-lookup"><span data-stu-id="a72b6-369">Excel displays an error message indicating that the worksheet is protected.</span></span>

7. <span data-ttu-id="a72b6-370">Снова нажмите кнопку " **переключать защиту листа** ", а элементы управления повторно включаются, а значения ячеек можно изменить.</span><span class="sxs-lookup"><span data-stu-id="a72b6-370">Choose the **Toggle Worksheet Protection** button again, and the controls are reenabled, and you can change cell values again.</span></span>

## <a name="open-a-dialog"></a><span data-ttu-id="a72b6-371">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="a72b6-371">Open a dialog</span></span>

<span data-ttu-id="a72b6-p154">На данном заключительном этапе, указанном в руководстве, вы откроете диалоговое окно в своей надстройке, передадите сообщение из процесса диалогового окна в процесс области задач и закроете диалоговое окно. Диалоговые окна надстройки Office *не модальные*: пользователь может продолжать работать и с документом в ведущем приложении Office, и с главной страницей в области задач.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p154">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="a72b6-374">Создание страницы диалогового окна</span><span class="sxs-lookup"><span data-stu-id="a72b6-374">Create the dialog page</span></span>

1. <span data-ttu-id="a72b6-375">В папке **./СРК** , расположенной в корне проекта, создайте папку с именем **Dialogs**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-375">In the **./src** folder that's located at the root of the project, create a new folder named **dialogs**.</span></span>

2. <span data-ttu-id="a72b6-376">В папке **./СРК/диалогс** создайте файл с именем **Popup. HTML**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-376">In the **./src/dialogs** folder, create new file named **popup.html**.</span></span>

3. <span data-ttu-id="a72b6-377">Добавьте указанную ниже разметку в **Popup. HTML**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-377">Add the following markup to **popup.html**.</span></span> <span data-ttu-id="a72b6-378">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-378">Note:</span></span>

   - <span data-ttu-id="a72b6-379">На странице находится `<input>`, где пользователь будет вводить свое имя, и кнопка, при нажатии которой имя будет отправлено на страницу области задач, где оно отобразится.</span><span class="sxs-lookup"><span data-stu-id="a72b6-379">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="a72b6-380">Разметка загружает скрипт с именем **Popup. js** , который будет создан на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="a72b6-380">The markup loads a script named **popup.js** that you will create in a later step.</span></span>

   - <span data-ttu-id="a72b6-381">Кроме того, загружается библиотека Office. js, так как она будет использоваться в **Popup. js**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-381">It also loads the Office.js library because it will be used in **popup.js**.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
            <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <p class="ms-font-xl">ENTER YOUR NAME</p>
            <input id="name-box" type="text"/><br/><br/>
            <button id="ok-button" class="ms-Button">OK</button>
        </body>
    </html>
    ```

4. <span data-ttu-id="a72b6-382">В папке **./СРК/диалогс** создайте файл с именем **Popup. js**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-382">In the **./src/dialogs** folder, create new file named **popup.js**.</span></span>

5. <span data-ttu-id="a72b6-383">Добавьте следующий код в **Popup. js**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-383">Add the following code to **popup.js**.</span></span> <span data-ttu-id="a72b6-384">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="a72b6-384">Note the following about this code:</span></span>

   - <span data-ttu-id="a72b6-385">*Каждая страница, вызывающая API в библиотеке Office. js, должна сначала убедиться, что библиотека полностью инициализирована.*</span><span class="sxs-lookup"><span data-stu-id="a72b6-385">*Every page that calls APIs in the Office.js library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="a72b6-386">Лучший способ сделать это — вызвать метод `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-386">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="a72b6-387">Если у вашей надстройки есть собственные задачи инициализации, код должен перейти к методу `then()`, связанному с вызовом `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-387">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="a72b6-388">Вызов `Office.onReady()` должен выполняться перед вызовами Office. js; Таким образом, назначение находится в файле скрипта, который загружается страницей, как в данном случае.</span><span class="sxs-lookup"><span data-stu-id="a72b6-388">The call of `Office.onReady()` must run before any calls to Office.js; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {

                // TODO1: Assign handler to the OK button.

            });

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="a72b6-p158">Замените `TODO1` приведенным ниже кодом. Вы создадите функцию `sendStringToParentPage` на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p158">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

7. <span data-ttu-id="a72b6-p159">Замените `TODO2` приведенным ниже кодом. Метод `messageParent` передает свой параметр родительской странице (в данном случае это страница на панели задач). Параметр может быть логическим или строковым. Во втором случае подразумевается все, что можно сериализовать, представив в виде строки (например, XML или JSON).</span><span class="sxs-lookup"><span data-stu-id="a72b6-p159">Replace `TODO2` with the following code. The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane. The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> <span data-ttu-id="a72b6-394">Файл **Popup. HTML** и загружаемый файл **Popup. js** выполняются в полностью отдельном процессе Microsoft EDGE или Internet Explorer 11 из области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="a72b6-394">The **popup.html** file, and the **popup.js** file that it loads, run in an entirely separate Microsoft Edge or Internet Explorer 11 process from the add-in's task pane.</span></span> <span data-ttu-id="a72b6-395">Если **файл App. js был** перечислен в один файл **пакета** . js, то \*\*\*\* надстройке потребуется загрузить две копии файла **пучок. js** , что противоречит назначению объединения.</span><span class="sxs-lookup"><span data-stu-id="a72b6-395">If **popup.js** was transpiled into the same **bundle.js** file as the **app.js** file, then the add-in would have to load two copies of the **bundle.js** file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="a72b6-396">Таким образом, эта надстройка не переопределяет файл **Popup. js** вообще.</span><span class="sxs-lookup"><span data-stu-id="a72b6-396">Therefore, this add-in does not transpile the **popup.js** file at all.</span></span>

### <a name="update-webpack-config-settings"></a><span data-ttu-id="a72b6-397">Обновление настроек конфигурации webpack</span><span class="sxs-lookup"><span data-stu-id="a72b6-397">Update webpack config settings</span></span>

<span data-ttu-id="a72b6-398">Откройте файл **Pack. config. js** в корневом каталоге проекта и выполните следующие действия.</span><span class="sxs-lookup"><span data-stu-id="a72b6-398">Open the file **webpack.config.js** in the root directory of the project and complete the following steps.</span></span>

1. <span data-ttu-id="a72b6-399">Найдите объект `entry` в объекте `config` и добавьте новую запись для `popup`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-399">Locate the `entry` object within the `config` object and add a new entry for `popup`.</span></span>

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    <span data-ttu-id="a72b6-400">После этого новый объект `entry` будет выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="a72b6-400">After you've done this, the new `entry` object will look like this:</span></span>

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
2. <span data-ttu-id="a72b6-401">Нахождение `plugins` массива в `config` объекте и добавление приведенного ниже объекта в конец этого массива.</span><span class="sxs-lookup"><span data-stu-id="a72b6-401">Locate the `plugins` array within the `config` object and add the following object to the end of that array.</span></span>

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    <span data-ttu-id="a72b6-402">После этого новый массив `plugins` будет выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="a72b6-402">After you've done this, the new `plugins` array will look like this:</span></span>

    ```js
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ['polyfill', 'taskpane']
      }),
      new CopyWebpackPlugin([
      {
        to: "taskpane.css",
        from: "./src/taskpane/taskpane.css"
      }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "popup.html",
        template: "./src/dialogs/popup.html",
        chunks: ["polyfill", "popup"]
      })
    ],
    ```

3. <span data-ttu-id="a72b6-403">Если локальный веб-сервер запущен, остановите его, закрыв окно команд узла.</span><span class="sxs-lookup"><span data-stu-id="a72b6-403">If the local web server is running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="a72b6-404">Выполните указанную ниже команду, чтобы повторно собрать проект.</span><span class="sxs-lookup"><span data-stu-id="a72b6-404">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="a72b6-405">Открытие диалогового окна из области задач</span><span class="sxs-lookup"><span data-stu-id="a72b6-405">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="a72b6-406">Откройте файл **./src/TaskPane/TaskPane.HTML**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-406">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="a72b6-407">Нахождение `<button>` элемента для `freeze-header` кнопки и добавление приведенной ниже разметки после этой строки:</span><span class="sxs-lookup"><span data-stu-id="a72b6-407">Locate the `<button>` element for the `freeze-header` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

3. <span data-ttu-id="a72b6-408">В диалоговом окне пользователю будет предложено ввести имя и передать имя пользователя в область задач.</span><span class="sxs-lookup"><span data-stu-id="a72b6-408">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="a72b6-409">Область задач отобразит его в подписи.</span><span class="sxs-lookup"><span data-stu-id="a72b6-409">The task pane will display it in a label.</span></span> <span data-ttu-id="a72b6-410">Сразу после того `button` , как вы только что добавили, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="a72b6-410">Immediately after the `button` that you just added, add the following markup:</span></span>

    ```html
    <label id="user-name"></label><br/><br/>
    ```

4. <span data-ttu-id="a72b6-411">Откройте файл **./СРК/таскпане/таскпане.ЖС**.</span><span class="sxs-lookup"><span data-stu-id="a72b6-411">Open the file **./src/taskpane/taskpane.js**.</span></span>

5. <span data-ttu-id="a72b6-412">В вызове `Office.onReady` метода Расположите строку, которая назначает обработчик нажатия `freeze-header` кнопки, и добавьте следующий код после этой строки.</span><span class="sxs-lookup"><span data-stu-id="a72b6-412">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `freeze-header` button, and add the following code after that line.</span></span> <span data-ttu-id="a72b6-413">Вы создадите метод `openDialog` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="a72b6-413">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

6. <span data-ttu-id="a72b6-414">Добавьте следующее объявление в конец файла.</span><span class="sxs-lookup"><span data-stu-id="a72b6-414">Add the following declaration to the end of the file.</span></span> <span data-ttu-id="a72b6-415">Эта переменная удерживает объект в контексте выполнения родительской страницы, который служит посредником для контекста выполнения страницы диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="a72b6-415">This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="a72b6-416">Добавьте указанную ниже функцию в конец файла (после объявления `dialog`).</span><span class="sxs-lookup"><span data-stu-id="a72b6-416">Add the following function to the end of the file (after the declaration of `dialog`).</span></span> <span data-ttu-id="a72b6-417">Важно отметить, что в этом коде *отсутствует* вызов `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-417">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="a72b6-418">Это связано с тем, что API, открывающий диалоговое окно, совместно используется всеми ведущими приложениями Office, поэтому относится к общему API JavaScript для Office, а не API для Excel.</span><span class="sxs-lookup"><span data-stu-id="a72b6-418">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="a72b6-419">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a72b6-419">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="a72b6-420">Примечание:</span><span class="sxs-lookup"><span data-stu-id="a72b6-420">Note:</span></span>

   - <span data-ttu-id="a72b6-421">Метод `displayDialogAsync` открывает диалоговое окно в центре экрана.</span><span class="sxs-lookup"><span data-stu-id="a72b6-421">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="a72b6-422">Первый параметр — это URL-адрес открываемой страницы.</span><span class="sxs-lookup"><span data-stu-id="a72b6-422">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="a72b6-p166">Второй параметр передает параметры. `height` и `width` — процентные значения размера окна для приложения Office.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p166">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="a72b6-425">Обработка сообщения из диалогового окна и закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="a72b6-425">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="a72b6-426">В `openDialog` функции в файле file **./СРК/таскпане/таскпане.ЖС**замените `TODO2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a72b6-426">Within the `openDialog` function in the file **./src/taskpane/taskpane.js**, replace `TODO2` with the following code.</span></span> <span data-ttu-id="a72b6-427">Примечание.</span><span class="sxs-lookup"><span data-stu-id="a72b6-427">Note:</span></span>

   - <span data-ttu-id="a72b6-428">Обратный вызов выполняется сразу же после успешного открытия диалогового окна и до того, как пользователь предпримет какие-либо действия в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="a72b6-428">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="a72b6-429">`result.value` — это объект, который выступает в качестве посредника между контекстами выполнения родительских страниц и страниц диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="a72b6-429">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="a72b6-p168">Функция `processMessage` будет создана на более позднем этапе. Этот обработчик будет обрабатывать любые значения, которые отправляются со страницы диалогового окна с вызовами функции `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p168">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="a72b6-432">Добавьте указанную ниже функцию после функции `openDialog`.</span><span class="sxs-lookup"><span data-stu-id="a72b6-432">Add the following function after the `openDialog` function.</span></span>

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

3. <span data-ttu-id="a72b6-433">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="a72b6-433">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="a72b6-434">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="a72b6-434">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="a72b6-435">Если область задач надстройки еще не открыта в Excel, перейдите на вкладку **Главная** и нажмите кнопку **Показать** область задач на ленте, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="a72b6-435">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="a72b6-436">Нажмите кнопку **Open Dialog** (Открыть диалоговое окно) в области задач.</span><span class="sxs-lookup"><span data-stu-id="a72b6-436">Choose the **Open Dialog** button in the task pane.</span></span>

4. <span data-ttu-id="a72b6-437">Когда диалоговое окно открыто, перетащите его и измените его размер.</span><span class="sxs-lookup"><span data-stu-id="a72b6-437">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="a72b6-438">Обратите внимание, что вы можете взаимодействовать с листом и нажимать другие кнопки в области задач, но вы не можете запустить второе диалоговое окно на одной и той же странице панели задач.</span><span class="sxs-lookup"><span data-stu-id="a72b6-438">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

5. <span data-ttu-id="a72b6-439">В диалоговом окне введите имя и нажмите кнопку **ОК** .</span><span class="sxs-lookup"><span data-stu-id="a72b6-439">In the dialog, enter a name and choose the **OK** button.</span></span> <span data-ttu-id="a72b6-440">В области задач отобразится имя, и диалоговое окно закроется.</span><span class="sxs-lookup"><span data-stu-id="a72b6-440">The name appears on the task pane and the dialog closes.</span></span>

6. <span data-ttu-id="a72b6-p171">При желании можно закомментировать строку `dialog.close();` в функции `processMessage`. Повторите шаги этого раздела. Диалоговое окно остается открытым, и вы можете изменить имя. Можно закрыть его вручную, нажав кнопку **X** в правом верхнему углу.</span><span class="sxs-lookup"><span data-stu-id="a72b6-p171">Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Руководство по Excel - Диалоговое окно](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a><span data-ttu-id="a72b6-446">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="a72b6-446">Next steps</span></span>

<span data-ttu-id="a72b6-447">В этом руководстве показано создание надстройки Excel для области задач, которая взаимодействует с таблицами, диаграммами, листами, диалоговыми окнами в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="a72b6-447">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="a72b6-448">Чтобы узнать больше о создании надстроек Excel, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="a72b6-448">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="a72b6-449">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="a72b6-449">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a><span data-ttu-id="a72b6-450">См. также</span><span class="sxs-lookup"><span data-stu-id="a72b6-450">See also</span></span>

* [<span data-ttu-id="a72b6-451">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a72b6-451">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="a72b6-452">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a72b6-452">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="a72b6-453">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a72b6-453">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="a72b6-454">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a72b6-454">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
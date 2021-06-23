---
title: Руководство по надстройкам Excel
description: В этом руководстве вы создадите надстройку Excel, которая создает, заполняет, фильтрует и сортирует таблицу, создает диаграмму, замораживает заголовок таблицы, защищает лист и открывает диалоговое окно.
ms.date: 05/12/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: f169499e343d2fc7fac89f407b78717536add4fc
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077241"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="5bdda-103">Учебник: Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="5bdda-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="5bdda-104">С помощью данного учебника вы сможете создать надстройку области задач Excel, которая выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="5bdda-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
>
> - <span data-ttu-id="5bdda-105">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="5bdda-105">Creates a table</span></span>
> - <span data-ttu-id="5bdda-106">Фильтрация и сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="5bdda-106">Filters and sorts a table</span></span>
> - <span data-ttu-id="5bdda-107">Создание графика</span><span class="sxs-lookup"><span data-stu-id="5bdda-107">Creates a chart</span></span>
> - <span data-ttu-id="5bdda-108">Закрепление заголовка таблицы</span><span class="sxs-lookup"><span data-stu-id="5bdda-108">Freezes a table header</span></span>
> - <span data-ttu-id="5bdda-109">Защита листа</span><span class="sxs-lookup"><span data-stu-id="5bdda-109">Protects a worksheet</span></span>
> - <span data-ttu-id="5bdda-110">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="5bdda-110">Opens a dialog</span></span>

> [!TIP]
> <span data-ttu-id="5bdda-111">Если вы уже завершили[создание надстройки панели задач Excel](../quickstarts/excel-quickstart-jquery.md)с помощью генератора Yeoman, и хотите использовать этот проект в качестве отправной точки для данного руководства, перейдите непосредственно в раздел[Создание таблицы](#create-a-table), чтобы начать работу с этим руководством.</span><span class="sxs-lookup"><span data-stu-id="5bdda-111">If you've already completed the [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md) quick start using the Yeoman generator, and want to use that project as a starting point for this tutorial, go directly to the [Create a table](#create-a-table) section to start this tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5bdda-112">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="5bdda-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="5bdda-113">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="5bdda-113">Create your add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="5bdda-114">**Выберите тип проекта:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="5bdda-114">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="5bdda-115">**Выберите тип сценария:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="5bdda-115">**Choose a script type:** `JavaScript`</span></span>
- <span data-ttu-id="5bdda-116">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="5bdda-116">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="5bdda-117">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="5bdda-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office.](../images/yo-office-excel.png)

<span data-ttu-id="5bdda-119">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="5bdda-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a><span data-ttu-id="5bdda-120">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="5bdda-120">Create a table</span></span>

<span data-ttu-id="5bdda-121">На этом этапе руководства мы проверим программным способом, поддерживает ли надстройка текущую версию Excel, установленную у пользователя, а также добавим таблицу на лист, заполним ее данными и отформатируем.</span><span class="sxs-lookup"><span data-stu-id="5bdda-121">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="5bdda-122">Написание кода надстройки</span><span class="sxs-lookup"><span data-stu-id="5bdda-122">Code the add-in</span></span>

1. <span data-ttu-id="5bdda-123">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="5bdda-123">Open the project in your code editor.</span></span>

2. <span data-ttu-id="5bdda-124">Откройте файл **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-124">Open the file **./src/taskpane/taskpane.html**.</span></span>  <span data-ttu-id="5bdda-125">Этот файл содержит HTML-разметку для панели задач.</span><span class="sxs-lookup"><span data-stu-id="5bdda-125">This file contains the HTML markup for the task pane.</span></span>

3. <span data-ttu-id="5bdda-126">Найдите элемент `<main>` и удалите все строки, которые появляются после открывающего тега `<main>` и перед закрывающим тегом `</main>`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-126">Locate the `<main>` element and delete all lines that appear after the opening `<main>` tag and before the closing `</main>` tag.</span></span>

4. <span data-ttu-id="5bdda-127">Добавляйте указанные ниже исправления сразу после открывающего тега `<main>`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-127">Add the following markup immediately after the opening `<main>` tag:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

5. <span data-ttu-id="5bdda-128">Откройте файл **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-128">Open the file **./src/taskpane/taskpane.js**.</span></span> <span data-ttu-id="5bdda-129">Этот файл содержит код API JavaScript для Office, облегчающий взаимодействие между областью задач и клиентским приложением Office.</span><span class="sxs-lookup"><span data-stu-id="5bdda-129">This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.</span></span>

6. <span data-ttu-id="5bdda-130">Удалите все ссылки на кнопку`run` и функцию`run()`, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="5bdda-130">Remove all references to the `run` button and the `run()` function by doing the following:</span></span>

    - <span data-ttu-id="5bdda-131">Найдите и удалите строку `document.getElementById("run").onclick = run;`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-131">Locate and delete the line `document.getElementById("run").onclick = run;`.</span></span>

    - <span data-ttu-id="5bdda-132">Найдите и удалите всю функцию `run()`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-132">Locate and delete the entire `run()` function.</span></span>

7. <span data-ttu-id="5bdda-133">В вызове метода `Office.onReady` найдите строку `if (info.host === Office.HostType.Excel) {` и добавьте следующий код непосредственно после этой строки.</span><span class="sxs-lookup"><span data-stu-id="5bdda-133">Within the `Office.onReady` method call, locate the line `if (info.host === Office.HostType.Excel) {` and add the following code immediately after that line.</span></span> <span data-ttu-id="5bdda-134">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-134">Note:</span></span>

    - <span data-ttu-id="5bdda-p104">Первая часть этого кода определяет, поддерживает ли установленная у пользователя версия Excel ту версию файла Excel.js, которая включает все API, используемые в этой серии руководств. В рабочей надстройке можно использовать текст условного блока, чтобы скрыть или отключить пользовательский интерфейс, где вызываются неподдерживаемые API. При этом пользователь по-прежнему сможет использовать те части надстройки, которые поддерживаются в его версии Excel.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p104">The first part of this code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    - <span data-ttu-id="5bdda-138">Вторая часть этого кода добавляет обработчик событий для кнопки `create-table`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-138">The second part of this code adds an event handler for the `create-table` button.</span></span>

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

8. <span data-ttu-id="5bdda-139">Добавьте следующую функцию в конец файла.</span><span class="sxs-lookup"><span data-stu-id="5bdda-139">Add the following function to the end of the file.</span></span> <span data-ttu-id="5bdda-140">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-140">Note:</span></span>

    - <span data-ttu-id="5bdda-p106">Бизнес-логика Excel.js будет добавлена в функцию, передаваемую методу `Excel.run`. Эта логика выполняется не сразу. Вместо этого она добавляется в очередь ожидания команд.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p106">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

    - <span data-ttu-id="5bdda-144">Метод `context.sync` отправляет все команды из очереди в Excel для выполнения.</span><span class="sxs-lookup"><span data-stu-id="5bdda-144">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

    - <span data-ttu-id="5bdda-p107">За методом `Excel.run` следует блок `catch`. Рекомендуется всегда следовать этой методике.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p107">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

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

9. <span data-ttu-id="5bdda-147">В функции `createTable()` замените `TODO1` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-147">Within the `createTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5bdda-148">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-148">Note:</span></span>

    - <span data-ttu-id="5bdda-p109">Код создает таблицу с помощью метода `add` коллекции таблиц листов, которая существует всегда, даже если она пуста. Это стандартный способ создания объектов Excel.js. API конструкторов классов не существуют, а для создания объекта Excel никогда не следует использовать оператор `new`. Вместо этого следует добавить его к объекту родительской коллекции.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p109">The code creates a table by using the `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

    - <span data-ttu-id="5bdda-p110">Первый параметр метода `add`— это диапазон, содержащий только первую строку, а не весь диапазон таблицы, который мы в конечном итоге будем использовать. Это связано с тем, что при заполнении строк данных (на следующем этапе) надстройка добавляет к таблице новые строки, а не записывает их в ячейки имеющихся строк. Это обычный шаблон, потому что количество строк в таблице часто неизвестно на момент ее создания.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p110">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a common pattern, because the number of rows a table will have is often unknown when the table is created.</span></span>

    - <span data-ttu-id="5bdda-156">Имена таблиц должны быть уникальными в рамках всей книги, а не только одного листа.</span><span class="sxs-lookup"><span data-stu-id="5bdda-156">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

10. <span data-ttu-id="5bdda-157">В функции `createTable()` замените `TODO2` следующим кодом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-157">Within the `createTable()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="5bdda-158">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-158">Note:</span></span>

    - <span data-ttu-id="5bdda-159">значения ячеек диапазона задаются с помощью массива массивов.</span><span class="sxs-lookup"><span data-stu-id="5bdda-159">The cell values of a range are set with an array of arrays.</span></span>

    - <span data-ttu-id="5bdda-p112">Новые строки создаются в таблице путем вызова метода `add` коллекции ее строк. Вы можете добавить несколько строк в одном вызове метода `add`, включив несколько массивов значений ячеек в родительский массив, передаваемый в качестве второго параметра.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p112">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

11. <span data-ttu-id="5bdda-162">В функции `createTable()` замените `TODO3` следующим кодом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-162">Within the `createTable()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="5bdda-163">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-163">Note:</span></span>

    - <span data-ttu-id="5bdda-164">код получает ссылку на столбец **Сумма**, передавая его индекс (с отсчетом от нуля) в метод `getItemAt` коллекции столбцов таблицы.</span><span class="sxs-lookup"><span data-stu-id="5bdda-164">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

        > [!NOTE]
        > <span data-ttu-id="5bdda-165">У объектов коллекций Excel.js (например, `TableCollection`, `WorksheetCollection` и `TableColumnCollection`) есть свойство `items`, представляющее собой массив дочерних типов объектов (например, `Table`, `Worksheet` или `TableColumn`). Однако сам объект `*Collection` не является массивом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-165">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

    - <span data-ttu-id="5bdda-166">Затем код форматирует диапазон столбца **Сумма** как денежные суммы в евро с точностью до второго знака после запятой.</span><span class="sxs-lookup"><span data-stu-id="5bdda-166">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span>

    - <span data-ttu-id="5bdda-p114">Напоследок он обеспечивает достаточные ширину столбцов и высоту строк для размещения самого длинного (или самого высокого) элемента данных. Обратите внимание, что код должен привести объекты `Range` к нужному формату. У объектов `TableColumn` и `TableRow` нет свойств формата.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p114">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

12. <span data-ttu-id="5bdda-170">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="5bdda-170">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5bdda-171">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="5bdda-171">Test the add-in</span></span>

1. <span data-ttu-id="5bdda-172">Выполните указанные ниже действия, чтобы запустить локальный веб-сервер и загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="5bdda-172">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5bdda-173">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="5bdda-173">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="5bdda-174">Если вам будет предложено установить сертификат после того, как вы запустите одну из указанных ниже команд, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="5bdda-174">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="5bdda-175">Если вы тестируете свою надстройку на Mac, перед продолжением выполните следующую команду в корневом каталоге вашего проекта.</span><span class="sxs-lookup"><span data-stu-id="5bdda-175">If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding.</span></span> <span data-ttu-id="5bdda-176">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="5bdda-176">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="5bdda-177">Чтобы проверить надстройку в Excel, выполните приведенную ниже команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="5bdda-177">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="5bdda-178">При этом запускается локальный веб-сервер (если он еще не запущен) и открывается приложение Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="5bdda-178">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="5bdda-179">Чтобы проверить надстройку в Excel в Интернете, выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="5bdda-179">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="5bdda-180">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="5bdda-180">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="5bdda-181">Чтобы использовать надстройку, откройте новый документ в Excel в Интернете, а затем загрузите неопубликованную надстройку, следуя инструкциям из статьи [Загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="5bdda-181">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

2. <span data-ttu-id="5bdda-182">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="5bdda-182">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-3b.png)

3. <span data-ttu-id="5bdda-184">В области задач нажмите кнопку **Создать таблицу**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-184">In the task pane, choose the **Create Table** button.</span></span>

    ![Снимок экрана с приложением Excel, демонстрирующий область задач надстройки с кнопкой "Создать таблицу", а также таблицу на листе, заполненную данными даты, продавца, категории и суммы.](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table&quot;></a><span data-ttu-id=&quot;5bdda-186&quot;>Фильтрация и сортировка таблицы</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-186&quot;>Filter and sort a table</span></span>

<span data-ttu-id=&quot;5bdda-187&quot;>Из этого раздела руководства вы узнаете, как отфильтровать и отсортировать созданную ранее таблицу.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-187&quot;>In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name=&quot;filter-the-table&quot;></a><span data-ttu-id=&quot;5bdda-188&quot;>Фильтрация таблицы</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-188&quot;>Filter the table</span></span>

1. <span data-ttu-id=&quot;5bdda-189&quot;>Откройте файл **./src/taskpane/taskpane.html**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-189&quot;>Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id=&quot;5bdda-190&quot;>Найдите элемент `<button>` для кнопки `create-table` и после нее добавьте следующий текст:</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-190&quot;>Locate the `<button>` element for the `create-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class=&quot;ms-Button&quot; id=&quot;filter-table&quot;>Filter Table</button><br/><br/>
    ```

3. <span data-ttu-id=&quot;5bdda-191&quot;>Откройте файл **./src/taskpane/taskpane.js**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-191&quot;>Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id=&quot;5bdda-192&quot;>В вызове метода `Office.onReady` найдите строку, которая назначает обработчик щелчка для кнопки `create-table`, и добавьте следующий код после этой строки:</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-192&quot;>Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById(&quot;filter-table").onclick = filterTable;
    ```

5. <span data-ttu-id="5bdda-193">Добавьте следующую функцию в конец файла:</span><span class="sxs-lookup"><span data-stu-id="5bdda-193">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="5bdda-194">В функции `filterTable()` замените `TODO1` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-194">Within the `filterTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5bdda-195">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-195">Note:</span></span>

   - <span data-ttu-id="5bdda-p120">Код получает ссылку на столбец, который нужно отфильтровать, передавая имя столбца методу `getItem`, а не передавая его индекс методу `getItemAt`, как это делает метод `createTable`. Так как пользователи могут перемещать столбцы, по заданному индексу может располагаться уже другой столбец. Следовательно, для получения ссылки безопаснее использовать имя столбца. Мы спокойно использовали метод `getItemAt` в предыдущем разделе, потому что мы использовали его в методе, который создает таблицу, и пользователь никак не мог переместить столбец.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p120">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="5bdda-200">Метод `applyValuesFilter` является одним из нескольких методов фильтрации объекта `Filter`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-200">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ```

### <a name="sort-the-table"></a><span data-ttu-id="5bdda-201">Сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="5bdda-201">Sort the table</span></span>

1. <span data-ttu-id="5bdda-202">Откройте файл **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-202">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="5bdda-203">Найдите элемент `<button>` для кнопки `filter-table` и после нее добавьте следующий текст:</span><span class="sxs-lookup"><span data-stu-id="5bdda-203">Locate the `<button>` element for the `filter-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

3. <span data-ttu-id="5bdda-204">Откройте файл **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-204">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="5bdda-205">В вызове метода `Office.onReady` найдите строку, которая назначает обработчик щелчка для кнопки `filter-table`, и добавьте следующий код после этой строки:</span><span class="sxs-lookup"><span data-stu-id="5bdda-205">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `filter-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

5. <span data-ttu-id="5bdda-206">Добавьте следующую функцию в конец файла:</span><span class="sxs-lookup"><span data-stu-id="5bdda-206">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="5bdda-207">В функции `sortTable()` замените `TODO1` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-207">Within the `sortTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5bdda-208">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-208">Note:</span></span>

   - <span data-ttu-id="5bdda-209">Код создает массив объектов `SortField`, состоящий из одного элемента, так как надстройка сортирует таблицу только по столбцу Merchant.</span><span class="sxs-lookup"><span data-stu-id="5bdda-209">The code creates an array of `SortField` objects, which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="5bdda-210">Свойство `key` объекта `SortField` — это нулевой индекс столбца, который используется для сортировки таблицы.</span><span class="sxs-lookup"><span data-stu-id="5bdda-210">The `key` property of a `SortField` object is the zero-based index of the column used for sorting.</span></span> <span data-ttu-id="5bdda-211">Строки в таблице сортируются на основе значений в столбце в соотетствующем столбце.</span><span class="sxs-lookup"><span data-stu-id="5bdda-211">The rows of the table are sorted based on the values in the referenced column.</span></span>

   - <span data-ttu-id="5bdda-212">Элемент `sort` объекта `Table` — это объект `TableSort`, а не метод.</span><span class="sxs-lookup"><span data-stu-id="5bdda-212">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="5bdda-213">Объекты `SortField` передаются методу `apply` объекта `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-213">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

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

7. <span data-ttu-id="5bdda-214">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="5bdda-214">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5bdda-215">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="5bdda-215">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="5bdda-216">Если область задач надстройки еще не открыта в Excel, перейдите на вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="5bdda-216">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="5bdda-217">Если таблица, ранее добавленная в этом руководстве, отсутствует на открытом листе, нажмите кнопку **Создать таблицу** в области задач.</span><span class="sxs-lookup"><span data-stu-id="5bdda-217">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button in the task pane.</span></span>

4. <span data-ttu-id="5bdda-218">Нажмите кнопки **Фильтровать таблицу** и **Сортировать таблицу** в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="5bdda-218">Choose the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

    ![Снимок экрана: приложение Excel с кнопками "Фильтровать таблицу" и "Сортировать таблицу", отображаемыми в области задач надстройки.](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart&quot;></a><span data-ttu-id=&quot;5bdda-220&quot;>Создание диаграммы</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-220&quot;>Create a chart</span></span>

<span data-ttu-id=&quot;5bdda-221&quot;>На этом этапе руководства мы создадим диаграмму, используя данные из ранее созданной таблицы, а затем отформатируем эту диаграмму.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-221&quot;>In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name=&quot;chart-a-chart-using-table-data&quot;></a><span data-ttu-id=&quot;5bdda-222&quot;>Создание диаграммы с помощью таблицы данных</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-222&quot;>Chart a chart using table data</span></span>

1. <span data-ttu-id=&quot;5bdda-223&quot;>Откройте файл **./src/taskpane/taskpane.html**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-223&quot;>Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id=&quot;5bdda-224&quot;>Найдите элемент `<button>` для кнопки `sort-table` и после нее добавьте следующий текст:</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-224&quot;>Locate the `<button>` element for the `sort-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class=&quot;ms-Button&quot; id=&quot;create-chart&quot;>Create Chart</button><br/><br/>
    ```

3. <span data-ttu-id=&quot;5bdda-225&quot;>Откройте файл **./src/taskpane/taskpane.js**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-225&quot;>Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id=&quot;5bdda-226&quot;>В вызове метода `Office.onReady` найдите строку, которая назначает обработчик щелчка для кнопки `sort-table`, и добавьте следующий код после этой строки:</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-226&quot;>Within the `Office.onReady` method call, locate the line that assigns a click handler to the `sort-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById(&quot;create-chart").onclick = createChart;
    ```

5. <span data-ttu-id="5bdda-227">Добавьте следующую функцию в конец файла:</span><span class="sxs-lookup"><span data-stu-id="5bdda-227">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="5bdda-228">В функции `createChart()` замените `TODO1` следующим кодом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-228">Within the `createChart()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5bdda-229">Обратите внимание, что для исключения строки заголовков в коде вместо метода `getRange` используется метод `Table.getDataBodyRange`, чтобы получить нужный диапазон данных для диаграммы.</span><span class="sxs-lookup"><span data-stu-id="5bdda-229">Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

7. <span data-ttu-id="5bdda-230">В функции `createChart()` замените `TODO2` следующим кодом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-230">Within the `createChart()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="5bdda-231">Обратите внимание на следующие параметры:</span><span class="sxs-lookup"><span data-stu-id="5bdda-231">Note the following parameters:</span></span>

   - <span data-ttu-id="5bdda-p126">Первый параметр метода `add` задает тип диаграммы. Существует несколько десятков типов.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p126">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="5bdda-234">Второй параметр задает диапазон данных, включаемых в диаграмму.</span><span class="sxs-lookup"><span data-stu-id="5bdda-234">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="5bdda-p127">Третий параметр определяет, как следует отображать на диаграмме ряд точек данных из таблицы: по строкам или по столбцам. Значение `auto` сообщает Excel, что следует выбрать оптимальный способ.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p127">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise. The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'Auto');
    ```

8. <span data-ttu-id="5bdda-237">В функции `createChart()` замените `TODO3` следующим кодом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-237">Within the `createChart()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="5bdda-238">Большая часть этого кода не требует объяснений.</span><span class="sxs-lookup"><span data-stu-id="5bdda-238">Most of this code is self-explanatory.</span></span> <span data-ttu-id="5bdda-239">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-239">Note:</span></span>

   - <span data-ttu-id="5bdda-p129">Параметры метода `setPosition` задают левую верхнюю и правую нижнюю ячейки области листа, которые должны содержать диаграмму. Excel может настраивать такие параметры, как ширина линий, чтобы диаграмма хорошо выглядела в выделенном для нее пространстве.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p129">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>

   - <span data-ttu-id="5bdda-p130">"Ряд" — это набор точек данных из столбца таблицы. Так как в таблице есть только один нестроковый столбец, Excel делает вывод, что это единственный столбец точек данных для диаграммы. Он рассматривает другие столбцы как метки диаграммы. Следовательно, в диаграмме будет только один ряд, обозначенный индексом 0. К нему следует добавить метку "Значение в &euro;".</span><span class="sxs-lookup"><span data-stu-id="5bdda-p130">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in &euro;".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "Right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in \u20AC';
    ```

9. <span data-ttu-id="5bdda-247">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="5bdda-247">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5bdda-248">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="5bdda-248">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="5bdda-249">Если область задач надстройки еще не открыта в Excel, перейдите на вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="5bdda-249">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="5bdda-250">Если таблица, ранее добавленная в этом руководстве, отсутствует на открытом листе, нажмите кнопку **Создать таблицу**, а затем кнопки **Фильтровать таблицу** и **Сортировать таблицу** в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="5bdda-250">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button, and then the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

4. <span data-ttu-id="5bdda-p131">Нажмите кнопку **Create Chart** (Создать диаграмму). Будет создана диаграмма, включающая только данные из отфильтрованных строк. Метки точек данных в нижней части диаграммы отсортированы согласно заданному для нее порядку, то есть по именам продавцов в обратном алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p131">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Снимок экрана: Excel с кнопкой "Создать диаграмму" в области задач надстройки и диаграммой на листе с данными расходов на продукты и образование.](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header&quot;></a><span data-ttu-id=&quot;5bdda-255&quot;>Закрепление заголовка таблицы</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-255&quot;>Freeze a table header</span></span>

<span data-ttu-id=&quot;5bdda-p132&quot;>Когда таблица достаточно длинная, при прокрутке строка заголовков может исчезать с экрана. В этом разделе учебника мы расскажем, как закрепить строку заголовков созданной ранее таблицы, чтобы она была видна, даже когда пользователь прокручивает лист.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-p132&quot;>When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name=&quot;freeze-the-tables-header-row&quot;></a><span data-ttu-id=&quot;5bdda-258&quot;>Закрепление строки заголовков таблицы</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-258&quot;>Freeze the table's header row</span></span>

1. <span data-ttu-id=&quot;5bdda-259&quot;>Откройте файл **./src/taskpane/taskpane.html**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-259&quot;>Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id=&quot;5bdda-260&quot;>Найдите элемент `<button>` для кнопки `create-chart` и после нее добавьте следующий текст:</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-260&quot;>Locate the `<button>` element for the `create-chart` button, and add the following markup after that line:</span></span>

    ```html
    <button class=&quot;ms-Button&quot; id=&quot;freeze-header&quot;>Freeze Header</button><br/><br/>
    ```

3. <span data-ttu-id=&quot;5bdda-261&quot;>Откройте файл **./src/taskpane/taskpane.js**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-261&quot;>Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id=&quot;5bdda-262&quot;>В вызове метода `Office.onReady` найдите строку, которая назначает обработчик щелчка для кнопки `create-chart`, и добавьте следующий код после этой строки:</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;5bdda-262&quot;>Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-chart` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById(&quot;freeze-header").onclick = freezeHeader;
    ```

5. <span data-ttu-id="5bdda-263">Добавьте следующую функцию в конец файла:</span><span class="sxs-lookup"><span data-stu-id="5bdda-263">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="5bdda-264">В функции `freezeHeader()` замените `TODO1` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-264">Within the `freezeHeader()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5bdda-265">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-265">Note:</span></span>

   - <span data-ttu-id="5bdda-266">Коллекция `Worksheet.freezePanes` — это набор закрепленных строк, которые не исчезают с экрана при прокрутке листа.</span><span class="sxs-lookup"><span data-stu-id="5bdda-266">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="5bdda-p134">Метод `freezeRows` принимает в качестве параметра количество строк сверху, которые необходимо закрепить. Мы передаем значение `1`, чтобы закрепить первую строку.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p134">The `freezeRows` method takes as a parameter the number of rows, from the top, that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

7. <span data-ttu-id="5bdda-269">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="5bdda-269">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5bdda-270">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="5bdda-270">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="5bdda-271">Если область задач надстройки еще не открыта в Excel, перейдите на вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="5bdda-271">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="5bdda-272">Если таблица, ранее добавленная в этом руководстве, присутствует на листе, удалите ее.</span><span class="sxs-lookup"><span data-stu-id="5bdda-272">If the table you added previously in this tutorial is present in the worksheet, delete it.</span></span>

4. <span data-ttu-id="5bdda-273">В области задач нажмите кнопку **Создать таблицу**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-273">In the task pane, choose the **Create Table** button.</span></span>

5. <span data-ttu-id="5bdda-274">Нажмите кнопку **Закрепить заголовок**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-274">In the task pane, choose the **Freeze Header** button.</span></span>

6. <span data-ttu-id="5bdda-275">Прокрутите лист вниз, чтобы убедиться, что заголовок таблицы по-прежнему остается на экране, даже когда верхние строки исчезают.</span><span class="sxs-lookup"><span data-stu-id="5bdda-275">Scroll down the worksheet far enough to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Снимок экрана: лист Excel с закрепленным заголовком таблицы.](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="5bdda-277">Защита листа</span><span class="sxs-lookup"><span data-stu-id="5bdda-277">Protect a worksheet</span></span>

<span data-ttu-id="5bdda-278">На этом этапе обучения, вы добавите на ленту кнопку, с помощью которой можно включить и выключить защиту листа.</span><span class="sxs-lookup"><span data-stu-id="5bdda-278">In this step of the tutorial, you'll add a button to the ribbon that toggles worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="5bdda-279">Настройка манифеста для добавления второй кнопки на ленту</span><span class="sxs-lookup"><span data-stu-id="5bdda-279">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="5bdda-280">Откройте файл манифеста **./manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-280">Open the manifest file **./manifest.xml**.</span></span>

2. <span data-ttu-id="5bdda-281">Найдите элемент `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-281">Locate the `<Control>` element.</span></span> <span data-ttu-id="5bdda-282">Этот элемент определяет кнопку **Show Taskpane** (Показать область задач) на вкладке **Главная**, которую вы используете для запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="5bdda-282">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="5bdda-283">Мы добавим вторую кнопку в эту же группу на ленте **Главная**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-283">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="5bdda-284">Добавьте следующий код между закрывающим тегом`</Control>` и закрывающим тегом`</Group>`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-284">In between the closing `</Control>` tag and the closing `</Group>` tag, add the following markup.</span></span>

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

3. <span data-ttu-id="5bdda-285">В XML-коде, добавленном в файл манифеста, замените `TODO1` строкой, которая присваивает кнопке идентификатор, уникальный в пределах этого файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="5bdda-285">Within the XML you just added to the manifest file, replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="5bdda-286">Так как кнопка будет включать и выключать защиту листа, укажите "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="5bdda-286">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="5bdda-287">После завершения открывающий тег элемента `Control` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-287">When you are done, the opening tag for the `Control` element should look like this:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="5bdda-288">Следующие три `TODO`s устанавливают идентификаторы ресурсов или `resid`s.</span><span class="sxs-lookup"><span data-stu-id="5bdda-288">The next three `TODO`s set resource IDs, or `resid`s.</span></span> <span data-ttu-id="5bdda-289">Ресурс должен быть строкой (максимальная длина — 32 символа), и вы создадите эти три строки на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="5bdda-289">A resource is a string (with a maximum length of 32 characters), and you'll create these three strings in a later step.</span></span> <span data-ttu-id="5bdda-290">Сейчас вам нужно присвоить идентификаторы ресурсам.</span><span class="sxs-lookup"><span data-stu-id="5bdda-290">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="5bdda-291">Кнопка должна называться "Переключение защиты", но у строки должен быть *идентификатор* "ProtectionButtonLabel", поэтому элемент `Label` выглядит следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-291">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the `Label` element should look like this:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="5bdda-292">Элемент `SuperTip` определяет подсказку для кнопки.</span><span class="sxs-lookup"><span data-stu-id="5bdda-292">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="5bdda-293">Заголовок этой подсказки должен совпадать с названием кнопки, поэтому мы используем тот же ИД ресурса — "ProtectionButtonLabel".</span><span class="sxs-lookup"><span data-stu-id="5bdda-293">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="5bdda-294">Описание подсказки будет следующим: "Click to turn protection of the worksheet on and off" (Нажмите для включения или выключения защиты листа).</span><span class="sxs-lookup"><span data-stu-id="5bdda-294">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="5bdda-295">У `resid` должно быть значение "ProtectionButtonToolTip".</span><span class="sxs-lookup"><span data-stu-id="5bdda-295">But the `resid` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="5bdda-296">Поэтому после завершения элемент `SuperTip` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-296">So, when you are done, the `SuperTip` element should look like this:</span></span>

    ```xml
    <Supertip>
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE]
   > <span data-ttu-id="5bdda-p139">В рабочей надстройке не нужно использовать один и тот же значок для двух разных кнопок, но сейчас мы предлагаем сделать это для простоты. Поэтому код `Icon` в новом теге `Control` представляет собой лишь копию элемента `Icon` из существующего тега `Control`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p139">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span>

6. <span data-ttu-id="5bdda-299">Для элемента `Action` в исходном элементе `Control`, задан тип `ShowTaskpane`, но новая кнопка будет не открывать область задач, а выполнять специальную функцию, которую вы создадите позже.</span><span class="sxs-lookup"><span data-stu-id="5bdda-299">The `Action` element inside the original `Control` element has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="5bdda-300">Поэтому замените `TODO5` на `ExecuteFunction` (тип действия для кнопок, запускающих специальные функции).</span><span class="sxs-lookup"><span data-stu-id="5bdda-300">So, replace `TODO5` with `ExecuteFunction`, which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="5bdda-301">Открывающий тег элемента `Action` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-301">The opening tag for the `Action` element should look like this:</span></span>

    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="5bdda-p141">У исходного элемента `Action` есть дочерние элементы, определяющие идентификатор области задач и URL-адрес страницы, которая должна быть открыта в области задач. Но у элемента `Action` типа `ExecuteFunction` есть один дочерний элемент, который именует функцию, выполняемую элементом управления. На более позднем этапе вы создадите функцию `toggleProtection`. Поэтому замените `TODO6` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-p141">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>

    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="5bdda-306">Теперь весь код `Control` должен выглядеть вот так:</span><span class="sxs-lookup"><span data-stu-id="5bdda-306">The entire `Control` markup should now look like the following:</span></span>

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

8. <span data-ttu-id="5bdda-307">Прокрутите страницу вниз до раздела `Resources` манифеста.</span><span class="sxs-lookup"><span data-stu-id="5bdda-307">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="5bdda-308">Добавьте приведенный ниже код в качестве дочернего элемента `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-308">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="5bdda-309">Добавьте приведенный ниже код в качестве дочернего элемента `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-309">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="5bdda-310">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="5bdda-310">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="5bdda-311">Создание функции защиты листа</span><span class="sxs-lookup"><span data-stu-id="5bdda-311">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="5bdda-312">Откройте файл **.\commands\commands.js**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-312">Open the file **.\commands\commands.js**.</span></span>

2. <span data-ttu-id="5bdda-313">Добавьте указанную ниже функцию сразу после функции `action`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-313">Add the following function immediately after the `action` function.</span></span> <span data-ttu-id="5bdda-314">Обратите внимание, что мы указываем параметр `args` для функции, а самая последняя строка функции вызывает `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-314">Note that we specify an `args` parameter to the function and the very last line of the function calls `args.completed`.</span></span> <span data-ttu-id="5bdda-315">Это требование для всех команд надстройки типа **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-315">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="5bdda-316">Это сигнализирует клиентскому приложению Office, что действие функции завершено и пользовательский интерфейс снова может отвечать на запросы.</span><span class="sxs-lookup"><span data-stu-id="5bdda-316">It signals the Office client application that the function has finished and the UI can become responsive again.</span></span>

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

3. <span data-ttu-id="5bdda-317">Добавьте следующую строку в конец файла:</span><span class="sxs-lookup"><span data-stu-id="5bdda-317">Add the following line to the end of the file:</span></span>

    ```js
    g.toggleProtection = toggleProtection;
    ```

4. <span data-ttu-id="5bdda-318">В функции `toggleProtection` замените `TODO1` следующим кодом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-318">Within the `toggleProtection` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5bdda-319">В этом коде используется свойство защиты объекта листа в стандартном шаблоне переключателя.</span><span class="sxs-lookup"><span data-stu-id="5bdda-319">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="5bdda-320">Объяснение `TODO2` будет приведено в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="5bdda-320">The `TODO2` will be explained in the next section.</span></span>

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="5bdda-321">Добавление кода для получения свойств документа в объекты скрипта области задач</span><span class="sxs-lookup"><span data-stu-id="5bdda-321">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="5bdda-322">В каждой функции, созданной в этом руководстве до настоящего момента, вы помещаете в очередь команды на *запись* в документе Office.</span><span class="sxs-lookup"><span data-stu-id="5bdda-322">In each function that you've created in this tutorial until now, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="5bdda-323">Каждая функция заканчивалась вызовом метода `context.sync()`, который отправляет выставленные в очередь команды документу для выполнения.</span><span class="sxs-lookup"><span data-stu-id="5bdda-323">Each function ended with a call to the `context.sync()` method, which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="5bdda-324">При этом код, который вы добавили на последнем этапе, вызывает свойство`sheet.protection.protected property`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-324">However, the code you added in the last step calls the `sheet.protection.protected property`.</span></span> <span data-ttu-id="5bdda-325">В этом заключается существенное отличие от ранее написанных функций, так как `sheet` является лишь объектом прокси, существующим в скрипте вашей области задач.</span><span class="sxs-lookup"><span data-stu-id="5bdda-325">This is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="5bdda-326">Объект-прокси не знает о фактическом состоянии защиты документа, поэтому его свойство `protection.protected` не может иметь реального значения.</span><span class="sxs-lookup"><span data-stu-id="5bdda-326">The proxy object doesn't know the actual protection state of the document, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="5bdda-327">Чтобы избежать возникновения ошибки, сначала нужно получить сведения о состоянии защиты от документа и задать значение `sheet.protection.protected`, используя их.</span><span class="sxs-lookup"><span data-stu-id="5bdda-327">To avoid an exception error, you must first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="5bdda-328">Процесс получения делится на три этапа:</span><span class="sxs-lookup"><span data-stu-id="5bdda-328">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="5bdda-329">Добавление в очередь команды для загрузки (т. е. получения) свойств, которые должен прочесть ваш код.</span><span class="sxs-lookup"><span data-stu-id="5bdda-329">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="5bdda-330">Вызов метода `sync` объекта контекста, чтобы можно было отправить документу находящуюся в очереди команду для выполнения, а также для возврата запрошенных данных.</span><span class="sxs-lookup"><span data-stu-id="5bdda-330">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="5bdda-331">Метод `sync` асинхронный, поэтому его выполнение должно быть завершено до того, как код вызовет полученные свойства.</span><span class="sxs-lookup"><span data-stu-id="5bdda-331">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="5bdda-332">Эти три действия должны выполняться каждый раз, когда коду нужно *прочесть* данные из документа Office.</span><span class="sxs-lookup"><span data-stu-id="5bdda-332">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="5bdda-333">В функции `toggleProtection` замените `TODO2` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-333">Within the `toggleProtection` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="5bdda-334">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-334">Note:</span></span>

   - <span data-ttu-id="5bdda-p146">У каждого объекта Excel есть метод `load`. Вы указываете свойства объекта, которые нужно прочесть в параметре как строку имен, разделенных запятыми. В этом случае нужно прочесть подсвойство свойства `protection`. На подсвойство нужно ссылаться почти так же, как и в остальных частях кода. Отличие заключается в том, что вместо символа "." нужно указать косую черту ("/").</span><span class="sxs-lookup"><span data-stu-id="5bdda-p146">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="5bdda-339">Чтобы логика переключения, которая считывает `sheet.protection.protected`, не срабатывала до выполнения `sync` и присвоения `sheet.protection.protected` правильного значения, полученного из документа, она будет перемещена (на следующем этапе) в функцию `then`, которая не выполняется до завершения `sync`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-339">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span>

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

2. <span data-ttu-id="5bdda-p147">Для двух операторов `return` не может использоваться один путь кода, который не разветвляется, поэтому удалите последнюю строку `return context.sync();` в конце `Excel.run`. Вы добавите новую последнюю строку `context.sync` позже.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p147">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="5bdda-342">Вырежьте структуру `if ... else` в функции `toggleProtection` и вставьте вместо `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-342">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="5bdda-p148">Замените `TODO4` приведенным ниже кодом. Примечание:</span><span class="sxs-lookup"><span data-stu-id="5bdda-p148">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="5bdda-345">Благодаря тому, что метод `sync` передается функции `then`, он не будет запускаться до добавления `sheet.protection.unprotect()` или `sheet.protection.protect()` в очередь.</span><span class="sxs-lookup"><span data-stu-id="5bdda-345">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="5bdda-346">Метод `then` вызывает любую функцию, которая ему передана. Не нужно вызывать `sync` дважды, поэтому уберите "()" после `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-346">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="5bdda-347">Когда все будет готово, функция должна выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="5bdda-347">When you are done, the entire function should look like the following:</span></span>

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

5. <span data-ttu-id="5bdda-348">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="5bdda-348">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5bdda-349">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="5bdda-349">Test the add-in</span></span>

1. <span data-ttu-id="5bdda-350">Закройте все приложения Office, в том числе Excel.</span><span class="sxs-lookup"><span data-stu-id="5bdda-350">Close all Office applications, including Excel.</span></span>

2. <span data-ttu-id="5bdda-p149">Очистите кэш Office, удалив содержимое (все файлы и вложенные папки) папки кэша. Это необходимо для полного удаления старой версии надстройки из клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p149">Delete the Office cache by deleting the contents (all the files and subfolders) of the cache folder. This is necessary to completely clear the old version of the add-in from the client application.</span></span>

    - <span data-ttu-id="5bdda-353">Для Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-353">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="5bdda-354">Для Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-354">For Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

      > [!NOTE]
      > <span data-ttu-id="5bdda-355">Если эта папка не существует, проверьте наличие следующих папок и в случае их присутствия удалите содержимое папки:</span><span class="sxs-lookup"><span data-stu-id="5bdda-355">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
      >  - <span data-ttu-id="5bdda-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`, где `{host}` — это приложение Office (например, `Excel`)</span><span class="sxs-lookup"><span data-stu-id="5bdda-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
      >  - <span data-ttu-id="5bdda-357">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`, где `{host}` — это приложение Office (например, `Excel`)</span><span class="sxs-lookup"><span data-stu-id="5bdda-357">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`

3. <span data-ttu-id="5bdda-358">Если локальный веб-сервер уже запущен, остановите его, закрыв окно команды узла.</span><span class="sxs-lookup"><span data-stu-id="5bdda-358">If the local web server is already running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="5bdda-359">Так как файл манифеста был обновлен, требуется повторно загрузить неопубликованную надстройку, используя обновленный файл манифеста.</span><span class="sxs-lookup"><span data-stu-id="5bdda-359">Because your manifest file has been updated, you must sideload your add-in again, using the updated manifest file.</span></span> <span data-ttu-id="5bdda-360">Запустите локальный веб-сервер и загрузите неопубликованную надстройку:</span><span class="sxs-lookup"><span data-stu-id="5bdda-360">Start the local web server and sideload your add-in:</span></span>

    - <span data-ttu-id="5bdda-361">Чтобы проверить надстройку в Excel, выполните приведенную ниже команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="5bdda-361">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="5bdda-362">При этом запускается локальный веб-сервер (если он еще не запущен) и открывается приложение Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="5bdda-362">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="5bdda-363">Чтобы проверить надстройку в Excel в Интернете, выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="5bdda-363">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="5bdda-364">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="5bdda-364">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="5bdda-365">Чтобы использовать надстройку, откройте новый документ в Excel в Интернете, а затем загрузите неопубликованную надстройку, следуя инструкциям из статьи [Загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="5bdda-365">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

5. <span data-ttu-id="5bdda-366">На вкладке **Главная** в Excel нажмите кнопку **Переключение защиты листа**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-366">On the **Home** tab in Excel, choose the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="5bdda-367">Обратите внимание, что большинство элементов управления на ленте отключены (серые), как показано на следующем снимке экрана.</span><span class="sxs-lookup"><span data-stu-id="5bdda-367">Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in the following screenshot.</span></span>

    ![Снимок экрана: лента Excel с выделенной и нажатой кнопкой "Включить защиту листа".](../images/excel-tutorial-ribbon-with-protection-on-2.png)

6. <span data-ttu-id="5bdda-p155">Выберите ячейку, как если бы вы хотели изменить ее содержимое. В Excel отобразится сообщение об ошибке, указывающее, что лист защищен.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p155">Choose a cell as you would if you wanted to change its content. Excel displays an error message indicating that the worksheet is protected.</span></span>

7. <span data-ttu-id="5bdda-372">Нажмите кнопку **Переключение защиты листа** еще раз, и элементы управления включатся, после чего вы сможете изменить значения ячеек.</span><span class="sxs-lookup"><span data-stu-id="5bdda-372">Choose the **Toggle Worksheet Protection** button again, and the controls are reenabled, and you can change cell values again.</span></span>

## <a name="open-a-dialog"></a><span data-ttu-id="5bdda-373">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="5bdda-373">Open a dialog</span></span>

<span data-ttu-id="5bdda-p156">На данном заключительном этапе, указанном в руководстве, вы откроете диалоговое окно в своей надстройке, передадите сообщение из процесса диалогового окна в процесс области задач и закроете диалоговое окно. Диалоговые окна надстройки Office *не модальные*: пользователь может продолжать работать и с документом в приложении Office, и с главной страницей в области задач.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p156">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="5bdda-376">Создание страницы диалогового окна</span><span class="sxs-lookup"><span data-stu-id="5bdda-376">Create the dialog page</span></span>

1. <span data-ttu-id="5bdda-377">В папке **./src**, расположенной в корне проекта, создайте папку с именем **dialogs**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-377">In the **./src** folder that's located at the root of the project, create a new folder named **dialogs**.</span></span>

2. <span data-ttu-id="5bdda-378">В папке **./src/dialogs** создайте файл с именем **popup.html**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-378">In the **./src/dialogs** folder, create new file named **popup.html**.</span></span>

3. <span data-ttu-id="5bdda-p157">Добавьте в файл **popup.html** следующий код. Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p157">Add the following markup to **popup.html**. Note:</span></span>

   - <span data-ttu-id="5bdda-381">На странице есть поле`<input>`, где пользователь будет вводить свое имя, и кнопка, при нажатии которой это имя будет отправлено в панель задач, где оно будет отображаться.</span><span class="sxs-lookup"><span data-stu-id="5bdda-381">The page has an `<input>` field where the user will enter their name, and a button that will send this name to the task pane where it will display.</span></span>

   - <span data-ttu-id="5bdda-382">Код загружает скрипт под названием **popup.js**, который будет создан на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="5bdda-382">The markup loads a script named **popup.js** that you will create in a later step.</span></span>

   - <span data-ttu-id="5bdda-383">Он также загружает библиотеку Office.js, так как она будет использоваться в **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-383">It also loads the Office.js library because it will be used in **popup.js**.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
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

4. <span data-ttu-id="5bdda-384">В папке **./src/dialogs** создайте файл с именем **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-384">In the **./src/dialogs** folder, create new file named **popup.js**.</span></span>

5. <span data-ttu-id="5bdda-385">Добавьте указанный ниже код в файл **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-385">Add the following code to **popup.js**.</span></span> <span data-ttu-id="5bdda-386">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="5bdda-386">Note the following about this code:</span></span>

   - <span data-ttu-id="5bdda-387">*Каждая страница, вызывающая API в библиотеке Office.js, должна сначала убедиться, что библиотека полностью инициализирована.*</span><span class="sxs-lookup"><span data-stu-id="5bdda-387">*Every page that calls APIs in the Office.js library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="5bdda-388">Лучший способ сделать это — вызвать метод `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-388">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="5bdda-389">Если у вашей надстройки есть собственные задачи инициализации, код должен перейти к методу `then()`, связанному с вызовом `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-389">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="5bdda-390">Вызов метода `Office.onReady()` должен выполняться до каких-либо вызовов Office.js, поэтому назначение указано в файле скрипта, загружаемом страницей, как в этом случае.</span><span class="sxs-lookup"><span data-stu-id="5bdda-390">The call of `Office.onReady()` must run before any calls to Office.js; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>

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

6. <span data-ttu-id="5bdda-p160">Замените `TODO1` приведенным ниже кодом. Вы создадите функцию `sendStringToParentPage` на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p160">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

7. <span data-ttu-id="5bdda-393">Замените `TODO2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-393">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="5bdda-394">Метод `messageParent` передает свой параметр родительской странице (в данном случае это страница на панели задач).</span><span class="sxs-lookup"><span data-stu-id="5bdda-394">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="5bdda-395">Параметр должен быть строковым. Это подразумевает все, что можно сериализовать в виде строки (например, XML или JSON), или любой тип, который можно представить в виде строки.</span><span class="sxs-lookup"><span data-stu-id="5bdda-395">The parameter must be a string, which includes anything that can be serialized as a string, such as XML or JSON, or any type that can be cast to a string.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> <span data-ttu-id="5bdda-396">Файл **popup.html** и загружаемый им файл **popup.js** выполняются в полностью отдельном процессе Microsoft Edge или Internet Explorer 11 из области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="5bdda-396">The **popup.html** file, and the **popup.js** file that it loads, run in an entirely separate Microsoft Edge or Internet Explorer 11 process from the add-in's task pane.</span></span> <span data-ttu-id="5bdda-397">Если файл **popup.js** был передан в тот же файл **bundle.js**, что и файл **app.js**, надстройка загрузит два экземпляра файла **bundle.js**, что противоречит цели объединения.</span><span class="sxs-lookup"><span data-stu-id="5bdda-397">If **popup.js** was transpiled into the same **bundle.js** file as the **app.js** file, then the add-in would have to load two copies of the **bundle.js** file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="5bdda-398">Поэтому эта надстройка вообще не передает файл **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-398">Therefore, this add-in does not transpile the **popup.js** file at all.</span></span>

### <a name="update-webpack-config-settings"></a><span data-ttu-id="5bdda-399">Обновление настроек конфигурации webpack</span><span class="sxs-lookup"><span data-stu-id="5bdda-399">Update webpack config settings</span></span>

<span data-ttu-id="5bdda-400">Откройте файл **webpack.config.js** в корневом каталоге проекта и выполните описанные ниже шаги.</span><span class="sxs-lookup"><span data-stu-id="5bdda-400">Open the file **webpack.config.js** in the root directory of the project and complete the following steps.</span></span>

1. <span data-ttu-id="5bdda-401">Найдите объект `entry` в объекте `config` и добавьте новую запись для `popup`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-401">Locate the `entry` object within the `config` object and add a new entry for `popup`.</span></span>

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    <span data-ttu-id="5bdda-402">После этого новый объект `entry` будет выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-402">After you've done this, the new `entry` object will look like this:</span></span>

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
2. <span data-ttu-id="5bdda-403">Найдите массив `plugins` в объекте `config` и добавьте следующий объект в конец массива.</span><span class="sxs-lookup"><span data-stu-id="5bdda-403">Locate the `plugins` array within the `config` object and add the following object to the end of that array.</span></span>

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    <span data-ttu-id="5bdda-404">После этого новый массив `plugins` будет выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5bdda-404">After you've done this, the new `plugins` array will look like this:</span></span>

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

3. <span data-ttu-id="5bdda-405">Если локальный веб-сервер запущен, остановите его, закрыв окно команды узла.</span><span class="sxs-lookup"><span data-stu-id="5bdda-405">If the local web server is running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="5bdda-406">Выполните указанную ниже команду, чтобы повторно собрать проект.</span><span class="sxs-lookup"><span data-stu-id="5bdda-406">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="5bdda-407">Открытие диалогового окна из области задач</span><span class="sxs-lookup"><span data-stu-id="5bdda-407">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="5bdda-408">Откройте файл **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-408">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="5bdda-409">Найдите элемент `<button>` для кнопки `freeze-header` и после нее добавьте следующий текст:</span><span class="sxs-lookup"><span data-stu-id="5bdda-409">Locate the `<button>` element for the `freeze-header` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

3. <span data-ttu-id="5bdda-410">В диалоговом окне пользователю будет предложено ввести имя и передать имя пользователя в область задач.</span><span class="sxs-lookup"><span data-stu-id="5bdda-410">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="5bdda-411">Область задач отобразит его в подписи.</span><span class="sxs-lookup"><span data-stu-id="5bdda-411">The task pane will display it in a label.</span></span> <span data-ttu-id="5bdda-412">Непосредственно после только что добавленного тега `button` добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="5bdda-412">Immediately after the `button` that you just added, add the following markup:</span></span>

    ```html
    <label id="user-name"></label><br/><br/>
    ```

4. <span data-ttu-id="5bdda-413">Откройте файл **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-413">Open the file **./src/taskpane/taskpane.js**.</span></span>

5. <span data-ttu-id="5bdda-414">В вызове метода `Office.onReady` найдите строку, назначающую обработчик щелчка для кнопки `freeze-header`, и добавьте следующий код после этой строки.</span><span class="sxs-lookup"><span data-stu-id="5bdda-414">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `freeze-header` button, and add the following code after that line.</span></span> <span data-ttu-id="5bdda-415">Вы создадите метод `openDialog` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="5bdda-415">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

6. <span data-ttu-id="5bdda-p165">Добавьте следующее объявление в конец файла. Эта переменная удерживает объект в контексте выполнения родительской страницы, который служит посредником для контекста выполнения страницы диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p165">Add the following declaration to the end of the file. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="5bdda-418">Добавьте следующую функцию в конец файла (после объявления `dialog`).</span><span class="sxs-lookup"><span data-stu-id="5bdda-418">Add the following function to the end of the file (after the declaration of `dialog`).</span></span> <span data-ttu-id="5bdda-419">Важно отметить, что в этом коде *отсутствует* вызов `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-419">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="5bdda-420">Это связано с тем, что API, открывающий диалоговое окно, совместно используется всеми приложениями Office, поэтому относится к общему API JavaScript для Office, а не API для Excel.</span><span class="sxs-lookup"><span data-stu-id="5bdda-420">This is because the API to open a dialog is shared among all Office applications, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="5bdda-p167">Замените `TODO1` приведенным ниже кодом. Примечание:</span><span class="sxs-lookup"><span data-stu-id="5bdda-p167">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="5bdda-423">Метод `displayDialogAsync` открывает диалоговое окно в центре экрана.</span><span class="sxs-lookup"><span data-stu-id="5bdda-423">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="5bdda-424">Первый параметр — это URL-адрес открываемой страницы.</span><span class="sxs-lookup"><span data-stu-id="5bdda-424">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="5bdda-p168">Второй параметр передает параметры. `height` и `width` — процентные значения размера окна для приложения Office.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="5bdda-427">Обработка сообщения из диалогового окна и закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="5bdda-427">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="5bdda-428">В функции `openDialog` в файле **./src/taskpane/taskpane.js** замените `TODO2` следующим кодом.</span><span class="sxs-lookup"><span data-stu-id="5bdda-428">Within the `openDialog` function in the file **./src/taskpane/taskpane.js**, replace `TODO2` with the following code.</span></span> <span data-ttu-id="5bdda-429">Примечание.</span><span class="sxs-lookup"><span data-stu-id="5bdda-429">Note:</span></span>

   - <span data-ttu-id="5bdda-430">Обратный вызов выполняется сразу же после успешного открытия диалогового окна и до того, как пользователь предпримет какие-либо действия в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="5bdda-430">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="5bdda-431">`result.value` — это объект, который выступает в качестве посредника между контекстами выполнения родительских страниц и страниц диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="5bdda-431">The `result.value` is the object that acts as an intermediary between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="5bdda-p170">Функция `processMessage` будет создана на более позднем этапе. Этот обработчик будет обрабатывать любые значения, которые отправляются со страницы диалогового окна с вызовами функции `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p170">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="5bdda-434">Добавьте указанную ниже функцию после функции `openDialog`.</span><span class="sxs-lookup"><span data-stu-id="5bdda-434">Add the following function after the `openDialog` function.</span></span>

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

3. <span data-ttu-id="5bdda-435">Убедитесь, что вы сохранили все изменения, внесенные в проект.</span><span class="sxs-lookup"><span data-stu-id="5bdda-435">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5bdda-436">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="5bdda-436">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="5bdda-437">Если область задач надстройки еще не открыта в Excel, перейдите на вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="5bdda-437">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="5bdda-438">Нажмите кнопку **Open Dialog** (Открыть диалоговое окно) в области задач.</span><span class="sxs-lookup"><span data-stu-id="5bdda-438">Choose the **Open Dialog** button in the task pane.</span></span>

4. <span data-ttu-id="5bdda-p171">Когда диалоговое окно открыто, перетащите его и измените его размер. Обратите внимание, что вы можете взаимодействовать с листом и нажимать другие кнопки в области задач, но невозможно запустить второе диалоговое окно на одной и той же странице панели задач.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p171">While the dialog is open, drag it and resize it. Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

5. <span data-ttu-id="5bdda-441">В диалоговом окне введите имя и нажмите кнопку **OK**.</span><span class="sxs-lookup"><span data-stu-id="5bdda-441">In the dialog, enter a name and choose the **OK** button.</span></span> <span data-ttu-id="5bdda-442">В области задач отобразится имя, и диалоговое окно закроется.</span><span class="sxs-lookup"><span data-stu-id="5bdda-442">The name appears on the task pane and the dialog closes.</span></span>

6. <span data-ttu-id="5bdda-p173">При желании можно закомментировать строку `dialog.close();` в функции `processMessage`. Повторите шаги этого раздела. Диалоговое окно остается открытым, и вы можете изменить имя. Можно закрыть его вручную, нажав кнопку **X** в правом верхнему углу.</span><span class="sxs-lookup"><span data-stu-id="5bdda-p173">Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Снимок экрана: Excel с кнопкой "Открыть диалоговое окно", отображаемой в области задач надстройки, и диалоговым окном, отображаемым поверх листа.](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a><span data-ttu-id="5bdda-448">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="5bdda-448">Next steps</span></span>

<span data-ttu-id="5bdda-449">В этом руководстве показано создание надстройки Excel для области задач, которая взаимодействует с таблицами, диаграммами, листами, диалоговыми окнами в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="5bdda-449">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="5bdda-450">Чтобы узнать больше о создании надстроек Excel, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="5bdda-450">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5bdda-451">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="5bdda-451">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a><span data-ttu-id="5bdda-452">См. также</span><span class="sxs-lookup"><span data-stu-id="5bdda-452">See also</span></span>

- [<span data-ttu-id="5bdda-453">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5bdda-453">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="5bdda-454">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5bdda-454">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="5bdda-455">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="5bdda-455">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)

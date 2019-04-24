---
title: Руководство по надстройкам Excel
description: В этом руководстве показана разработка надстройки Excel, которая создает, заполняет, фильтрует и сортирует данные таблиц, создает диаграммы, закрепляет заголовки таблиц, защищает листы и открывает диалоговые окна.
ms.date: 01/28/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 410b2391d207f7c83f9accb349448dbc0c92a0e2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451314"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="ff647-103">Учебник: Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="ff647-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="ff647-104">С помощью данного учебника вы сможете создать надстройку области задач Excel, которая выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="ff647-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="ff647-105">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-105">Creates a table</span></span>
> * <span data-ttu-id="ff647-106">Фильтрация и сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="ff647-107">Создание графика</span><span class="sxs-lookup"><span data-stu-id="ff647-107">Creates a chart</span></span>
> * <span data-ttu-id="ff647-108">Закрепление заголовка таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-108">Freezes a table header</span></span>
> * <span data-ttu-id="ff647-109">Защита листа</span><span class="sxs-lookup"><span data-stu-id="ff647-109">Protects a worksheet</span></span>
> * <span data-ttu-id="ff647-110">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="ff647-110">Opens a dialog</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ff647-111">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="ff647-111">Prerequisites</span></span>

<span data-ttu-id="ff647-112">Для работы с этим учебником необходимо установить указанные ниже компоненты.</span><span class="sxs-lookup"><span data-stu-id="ff647-112">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="ff647-p101">Excel 2016, версия 1711 (сборка 8730.1000 "нажми и работай") или более поздняя. Чтобы установить эту версию, необходимо быть участником программы предварительной оценки Office. [Дополнительные сведения](https://products.office.com/office-insider?tab=tab-1)</span><span class="sxs-lookup"><span data-stu-id="ff647-p101">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="ff647-116">Node</span><span class="sxs-lookup"><span data-stu-id="ff647-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="ff647-117">[Git Bash](https://git-scm.com/downloads) (или другой клиент Git)</span><span class="sxs-lookup"><span data-stu-id="ff647-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

- <span data-ttu-id="ff647-118">Чтобы протестировать надстройку в этом руководстве, необходимо подключиться к Интернету.</span><span class="sxs-lookup"><span data-stu-id="ff647-118">You need to have an Internet connection to test the add-in in this tutorial.</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="ff647-119">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="ff647-119">Create your add-in project</span></span>

<span data-ttu-id="ff647-120">Выполните указанные ниже действия для создания проекта надстройки Excel, который будет использоваться в качестве основы для этого учебника.</span><span class="sxs-lookup"><span data-stu-id="ff647-120">Complete the following steps to create the Excel add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="ff647-121">Клонируйте репозиторий GitHub [Excel Add-in Tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span><span class="sxs-lookup"><span data-stu-id="ff647-121">Clone the GitHub repository [Excel add-in tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="ff647-122">Откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="ff647-122">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="ff647-123">Выполните команду `npm install`, чтобы установить инструменты и библиотеки, указанные в файле package.json.</span><span class="sxs-lookup"><span data-stu-id="ff647-123">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="ff647-124">Сделайте так, чтобы операционная система компьютера разработки доверяла сертификату. Для этого выполните действия, описанные в [этой статье](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="ff647-124">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="create-a-table"></a><span data-ttu-id="ff647-125">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-125">Create a table</span></span>

<span data-ttu-id="ff647-126">На этом этапе руководства мы проверим программным способом, поддерживает ли надстройка текущую версию Excel, установленную у пользователя, а также добавим таблицу на лист, заполним ее данными и отформатируем.</span><span class="sxs-lookup"><span data-stu-id="ff647-126">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="ff647-127">Написание кода надстройки</span><span class="sxs-lookup"><span data-stu-id="ff647-127">Code the add-in</span></span>

1. <span data-ttu-id="ff647-128">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="ff647-128">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff647-129">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="ff647-129">Open the file index.html.</span></span>

3. <span data-ttu-id="ff647-130">Замените `TODO1` на следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="ff647-130">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="ff647-131">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-131">Open the app.js file.</span></span>

5. <span data-ttu-id="ff647-p102">Замените `TODO1` на приведенный ниже код. Этот код определяет, поддерживает ли установленная у пользователя версия Excel ту версию файла Excel.js, которая включает все API, используемые в этой серии руководств. В рабочей надстройке можно использовать текст условного блока, чтобы скрыть или отключить пользовательский интерфейс, где вызываются неподдерживаемые API. При этом пользователь по-прежнему сможет использовать те части надстройки, которые поддерживаются в его версии Excel.</span><span class="sxs-lookup"><span data-stu-id="ff647-p102">Replace the `TODO1` with the following code. This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="ff647-136">Замените `TODO2` на следующий код:</span><span class="sxs-lookup"><span data-stu-id="ff647-136">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="ff647-137">Замените `TODO3` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="ff647-137">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="ff647-138">Примечание.</span><span class="sxs-lookup"><span data-stu-id="ff647-138">Note:</span></span>

   - <span data-ttu-id="ff647-p104">Бизнес-логика Excel.js будет добавлена в функцию, передаваемую методу `Excel.run`. Эта логика выполняется не сразу. Вместо этого она добавляется в очередь ожидания команд.</span><span class="sxs-lookup"><span data-stu-id="ff647-p104">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="ff647-142">Метод `context.sync` отправляет все команды из очереди в Excel для выполнения.</span><span class="sxs-lookup"><span data-stu-id="ff647-142">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

   - <span data-ttu-id="ff647-p105">За методом `Excel.run` следует блок `catch`. Рекомендуется всегда следовать этой методике.</span><span class="sxs-lookup"><span data-stu-id="ff647-p105">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="ff647-p106">Замените `TODO4` на приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ff647-p106">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="ff647-p107">код создает таблицу с помощью метода `add` коллекции таблиц на листе, которая всегда существует, даже если она пуста. Это стандартный способ создания объектов Excel.js. API конструкторов классов не существуют, а для создания объекта Excel никогда не следует использовать оператор `new`. Вместо этого следует добавить его к объекту родительской коллекции.</span><span class="sxs-lookup"><span data-stu-id="ff647-p107">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

   - <span data-ttu-id="ff647-p108">Первый параметр метода `add`— это диапазон, содержащий только первую строку, а не весь диапазон таблицы, который мы в конечном итоге будем использовать. Это связано с тем, что при заполнении строк данных (на следующем этапе) надстройка добавляет к таблице новые строки, а не записывает их в ячейки имеющихся строк. Такой шаблон более распространен, так как количество строк в таблице часто неизвестно на момент ее создания.</span><span class="sxs-lookup"><span data-stu-id="ff647-p108">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

   - <span data-ttu-id="ff647-154">Имена таблиц должны быть уникальными в рамках всей книги, а не только одного листа.</span><span class="sxs-lookup"><span data-stu-id="ff647-154">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="ff647-155">Замените `TODO5` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="ff647-155">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="ff647-156">Примечание:</span><span class="sxs-lookup"><span data-stu-id="ff647-156">Note:</span></span>

   - <span data-ttu-id="ff647-157">значения ячеек диапазона задаются с помощью массива массивов.</span><span class="sxs-lookup"><span data-stu-id="ff647-157">The cell values of a range are set with an array of arrays.</span></span>

   - <span data-ttu-id="ff647-p110">Новые строки создаются в таблице путем вызова метода `add` коллекции ее строк. Вы можете добавить несколько строк в одном вызове метода `add`, включив несколько массивов значений ячеек в родительский массив, передаваемый в качестве второго параметра.</span><span class="sxs-lookup"><span data-stu-id="ff647-p110">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="ff647-p111">Замените `TODO6` на приведенный ниже код. Примечание:</span><span class="sxs-lookup"><span data-stu-id="ff647-p111">Replace `TODO6` with the following code. Note:</span></span>

   - <span data-ttu-id="ff647-162">код получает ссылку на столбец **Сумма**, передавая его индекс (с отсчетом от нуля) в метод `getItemAt` коллекции столбцов таблицы.</span><span class="sxs-lookup"><span data-stu-id="ff647-162">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff647-163">У объектов коллекций Excel.js (например, `TableCollection`, `WorksheetCollection` и `TableColumnCollection`) есть свойство `items`, представляющее собой массив дочерних типов объектов (например, `Table`, `Worksheet` или `TableColumn`). Однако сам объект `*Collection` не является массивом.</span><span class="sxs-lookup"><span data-stu-id="ff647-163">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="ff647-164">Затем код форматирует диапазон столбца **Сумма** как денежные суммы в евро с точностью до второго знака после запятой.</span><span class="sxs-lookup"><span data-stu-id="ff647-164">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

   - <span data-ttu-id="ff647-p112">Напоследок он обеспечивает достаточные ширину столбцов и высоту строк для размещения самого длинного (или самого высокого) элемента данных. Обратите внимание, что код должен привести объекты `Range` к нужному формату. У объектов `TableColumn` и `TableRow` нет свойств формата.</span><span class="sxs-lookup"><span data-stu-id="ff647-p112">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff647-168">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="ff647-168">Test the add-in</span></span>

1. <span data-ttu-id="ff647-169">Откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="ff647-169">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="ff647-170">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="ff647-170">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff647-171">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="ff647-171">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff647-172">Загрузите неопубликованную надстройку одним из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="ff647-172">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="ff647-173">[Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="ff647-173">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="ff647-174">[Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="ff647-174">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="ff647-175">[iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="ff647-175">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="ff647-176">В меню **Главная** выберите пункт **Показать область задач**.</span><span class="sxs-lookup"><span data-stu-id="ff647-176">On the **Home** menu, choose **Show Taskpane**.</span></span>

6. <span data-ttu-id="ff647-177">В области задач нажмите кнопку **Create Table** (Создать таблицу).</span><span class="sxs-lookup"><span data-stu-id="ff647-177">In the task pane, choose **Create Table**.</span></span>

    ![Руководство по Excel: создание таблицы](../images/excel-tutorial-create-table.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="ff647-179">Фильтрация и сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-179">Filter and sort a table</span></span>

<span data-ttu-id="ff647-180">Из этого раздела руководства вы узнаете, как отфильтровать и отсортировать созданную ранее таблицу.</span><span class="sxs-lookup"><span data-stu-id="ff647-180">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="ff647-181">Фильтрация таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-181">Filter the table</span></span>

1. <span data-ttu-id="ff647-182">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="ff647-182">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff647-183">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="ff647-183">Open the file index.html.</span></span>

3. <span data-ttu-id="ff647-184">Под элементом `div`, содержащим кнопку `create-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="ff647-184">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. <span data-ttu-id="ff647-185">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-185">Open the app.js file.</span></span>

5. <span data-ttu-id="ff647-186">Под строкой, назначающей обработчик нажатия кнопки `create-table`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="ff647-186">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="ff647-187">Под функцией `createTable` добавьте следующую функцию:</span><span class="sxs-lookup"><span data-stu-id="ff647-187">Just below the `createTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="ff647-188">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="ff647-188">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="ff647-189">Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ff647-189">Note:</span></span>

   - <span data-ttu-id="ff647-p114">Код получает ссылку на столбец, который нужно отфильтровать, передавая имя столбца методу `getItem`, а не передавая его индекс методу `getItemAt`, как это делает метод `createTable`. Так как пользователи могут перемещать столбцы, по заданному индексу может располагаться уже другой столбец. Следовательно, для получения ссылки безопаснее использовать имя столбца. Мы спокойно использовали метод `getItemAt` в предыдущем разделе, потому что мы использовали его в методе, который создает таблицу, и пользователь никак не мог переместить столбец.</span><span class="sxs-lookup"><span data-stu-id="ff647-p114">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="ff647-194">Метод `applyValuesFilter` является одним из нескольких методов фильтрации объекта `Filter`.</span><span class="sxs-lookup"><span data-stu-id="ff647-194">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="ff647-195">Сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-195">Sort the table</span></span>

1. <span data-ttu-id="ff647-196">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="ff647-196">Open the file index.html.</span></span>

2. <span data-ttu-id="ff647-197">Под элементом `div`, содержащим кнопку `filter-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="ff647-197">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. <span data-ttu-id="ff647-198">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-198">Open the app.js file.</span></span>

4. <span data-ttu-id="ff647-199">Под строкой, назначающей обработчик нажатия кнопки `filter-table`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="ff647-199">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="ff647-200">Под функцией `filterTable` добавьте приведенную ниже функцию.</span><span class="sxs-lookup"><span data-stu-id="ff647-200">Below the `filterTable` function add the following function.</span></span>

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

6. <span data-ttu-id="ff647-201">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="ff647-201">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="ff647-202">Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ff647-202">Note:</span></span>

   - <span data-ttu-id="ff647-203">Код создает массив объектов `SortField`, состоящий из одного элемента, так как надстройка сортирует таблицу только по столбцу Merchant.</span><span class="sxs-lookup"><span data-stu-id="ff647-203">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="ff647-204">Свойство `key` объекта `SortField` — это отсчитываемый от нуля индекс столбца, по которому необходимо сортировать таблицу.</span><span class="sxs-lookup"><span data-stu-id="ff647-204">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="ff647-205">Элемент `sort` объекта `Table` — это объект `TableSort`, а не метод.</span><span class="sxs-lookup"><span data-stu-id="ff647-205">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="ff647-206">Объекты `SortField` передаются методу `apply` объекта `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="ff647-206">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

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

### <a name="test-the-add-in"></a><span data-ttu-id="ff647-207">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="ff647-207">Test the add-in</span></span>

1. <span data-ttu-id="ff647-208">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши **Ctrl+C**, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="ff647-208">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="ff647-209">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="ff647-209">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff647-210">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="ff647-210">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="ff647-211">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="ff647-211">In order to do this, you need to kill the server process so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="ff647-212">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="ff647-212">After the build, you restart the server.</span></span> <span data-ttu-id="ff647-213">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="ff647-213">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ff647-214">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="ff647-214">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff647-215">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="ff647-215">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff647-216">Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="ff647-216">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="ff647-217">Если по той или иной причине на открытом листе нет таблицы, нажмите в области задач кнопку **Create Table** (Создать таблицу).</span><span class="sxs-lookup"><span data-stu-id="ff647-217">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table**.</span></span>

6. <span data-ttu-id="ff647-218">Нажмите кнопки **Filter Table** (Фильтровать таблицу) и **Sort Table** (Сортировать таблицу) в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="ff647-218">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Учебник Excel - Фильтрация и сортировка таблицы](../images/excel-tutorial-filter-and-sort-table.png)

## <a name="create-a-chart"></a><span data-ttu-id="ff647-220">Создание диаграммы</span><span class="sxs-lookup"><span data-stu-id="ff647-220">Create a chart</span></span>

<span data-ttu-id="ff647-221">На этом этапе руководства мы создадим диаграмму, используя данные из ранее созданной таблицы, а затем отформатируем эту диаграмму.</span><span class="sxs-lookup"><span data-stu-id="ff647-221">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="ff647-222">Создание диаграммы с помощью таблицы данных</span><span class="sxs-lookup"><span data-stu-id="ff647-222">Chart a chart using table data</span></span>

1. <span data-ttu-id="ff647-223">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="ff647-223">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff647-224">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="ff647-224">Open the file index.html.</span></span>

3. <span data-ttu-id="ff647-225">Под элементом `div`, содержащим кнопку `sort-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="ff647-225">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="ff647-226">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-226">Open the app.js file.</span></span>

5. <span data-ttu-id="ff647-227">Под строкой, назначающей обработчик нажатия кнопки `sort-chart`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="ff647-227">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="ff647-228">Под функцией `sortTable` добавьте приведенную ниже функцию.</span><span class="sxs-lookup"><span data-stu-id="ff647-228">Below the `sortTable` function add the following function.</span></span>

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

7. <span data-ttu-id="ff647-p119">Замените `TODO1` приведенным ниже кодом. Обратите внимание на то, что для исключения строки заголовков в коде вместо метода `Table.getDataBodyRange` используется метод `getRange`, чтобы получить нужный диапазон данных для диаграммы.</span><span class="sxs-lookup"><span data-stu-id="ff647-p119">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="ff647-p120">Замените `TODO2` приведенным ниже кодом. Обратите внимание на следующие параметры:</span><span class="sxs-lookup"><span data-stu-id="ff647-p120">Replace `TODO2` with the following code. Note the following parameters:</span></span>

   - <span data-ttu-id="ff647-p121">Первый параметр метода `add` задает тип диаграммы. Существует несколько десятков типов.</span><span class="sxs-lookup"><span data-stu-id="ff647-p121">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="ff647-235">Второй параметр задает диапазон данных, включаемых в диаграмму.</span><span class="sxs-lookup"><span data-stu-id="ff647-235">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="ff647-236">Третий параметр определяет, как следует отображать на диаграмме ряд точек данных из таблицы: по строкам или по столбцам.</span><span class="sxs-lookup"><span data-stu-id="ff647-236">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="ff647-237">Значение `auto` сообщает Excel, что следует выбрать оптимальный способ.</span><span class="sxs-lookup"><span data-stu-id="ff647-237">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="ff647-p123">Замените `TODO3` на приведенный ниже код. Большая часть этого кода не требует объяснений. Примечание.</span><span class="sxs-lookup"><span data-stu-id="ff647-p123">Replace `TODO3` with the following code. Most of this code is self-explanatory. Note:</span></span>
   
   - <span data-ttu-id="ff647-p124">Параметры метода `setPosition` задают левую верхнюю и правую нижнюю ячейки области листа, которые должны содержать диаграмму. Excel может настраивать такие параметры, как ширина линий, чтобы диаграмма хорошо выглядела в выделенном для нее пространстве.</span><span class="sxs-lookup"><span data-stu-id="ff647-p124">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="ff647-p125">"Ряд" — это набор точек данных из столбца таблицы. Так как в таблице есть только один нестроковый столбец, Excel делает вывод, что это единственный столбец точек данных для диаграммы. Он рассматривает другие столбцы как метки диаграммы. Следовательно, в диаграмме будет только один ряд, обозначенный индексом 0. К нему следует добавить метку "Значение в €".</span><span class="sxs-lookup"><span data-stu-id="ff647-p125">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff647-248">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="ff647-248">Test the add-in</span></span>

1. <span data-ttu-id="ff647-249">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши **Ctrl+C**, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="ff647-249">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="ff647-250">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="ff647-250">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff647-p127">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу. Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки. После сборки необходимо перезапустить сервер. Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="ff647-p127">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ff647-255">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="ff647-255">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff647-256">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="ff647-256">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff647-257">Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="ff647-257">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="ff647-258">Если по той или иной причине на открытом листе нет таблицы, нажмите в области задач кнопку **Create Table** (Создать таблицу), а затем — кнопки **Filter Table** (Фильтровать таблицу) и **Sort Table** (Сортировать таблицу) в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="ff647-258">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>

6. <span data-ttu-id="ff647-p128">Нажмите кнопку **Create Chart** (Создать диаграмму). Будет создана диаграмма, включающая только данные из отфильтрованных строк. Метки точек данных в нижней части диаграммы отсортированы согласно заданному для нее порядку, то есть по именам продавцов в обратном алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="ff647-p128">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Руководство по Excel - Создание диаграммы](../images/excel-tutorial-create-chart.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="ff647-263">Закрепление заголовка таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-263">Freeze a table header</span></span>

<span data-ttu-id="ff647-p129">Когда таблица достаточно длинная, при прокрутке строка заголовков может исчезать с экрана. В этом разделе учебника мы расскажем, как закрепить строку заголовков созданной ранее таблицы, чтобы она была видна, даже когда пользователь прокручивает лист.</span><span class="sxs-lookup"><span data-stu-id="ff647-p129">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="ff647-266">Закрепление строки заголовков таблицы</span><span class="sxs-lookup"><span data-stu-id="ff647-266">Freeze the table's header row</span></span>

1. <span data-ttu-id="ff647-267">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="ff647-267">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff647-268">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="ff647-268">Open the file index.html.</span></span>

3. <span data-ttu-id="ff647-269">Под элементом `div`, содержащим кнопку `create-chart`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="ff647-269">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. <span data-ttu-id="ff647-270">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-270">Open the app.js file.</span></span>

5. <span data-ttu-id="ff647-271">Под строкой, назначающей обработчик нажатия кнопки `create-chart`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="ff647-271">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="ff647-272">Под функцией `createChart` добавьте следующую функцию:</span><span class="sxs-lookup"><span data-stu-id="ff647-272">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="ff647-p130">Замените `TODO1` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ff647-p130">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="ff647-275">Коллекция `Worksheet.freezePanes` — это набор закрепленных строк, которые не исчезают с экрана при прокрутке листа.</span><span class="sxs-lookup"><span data-stu-id="ff647-275">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="ff647-p131">Метод `freezeRows` принимает в качестве параметра количество строк сверху, которые необходимо закрепить. Мы передаем значение `1`, чтобы закрепить первую строку.</span><span class="sxs-lookup"><span data-stu-id="ff647-p131">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff647-278">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="ff647-278">Test the add-in</span></span>

1. <span data-ttu-id="ff647-279">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши **Ctrl+C**, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="ff647-279">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="ff647-280">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="ff647-280">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff647-p133">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу. Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки. После сборки необходимо перезапустить сервер. Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="ff647-p133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ff647-285">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="ff647-285">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff647-286">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="ff647-286">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff647-287">Повторно загрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**.</span><span class="sxs-lookup"><span data-stu-id="ff647-287">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="ff647-288">Если таблица на листе, удалите ее.</span><span class="sxs-lookup"><span data-stu-id="ff647-288">If the table is in the worksheet, delete it.</span></span>

6. <span data-ttu-id="ff647-289">В области задач нажмите кнопку **Create Table** (Создать таблицу).</span><span class="sxs-lookup"><span data-stu-id="ff647-289">In the task pane, choose **Create Table**.</span></span>

7. <span data-ttu-id="ff647-290">Нажмите кнопку **Freeze Header** (Закрепить заголовок).</span><span class="sxs-lookup"><span data-stu-id="ff647-290">Choose the **Freeze Header** button.</span></span>

8. <span data-ttu-id="ff647-291">Прокрутите лист вниз, чтобы убедиться, что заголовок таблицы по-прежнему остается на экране, даже когда более высокие строки исчезают.</span><span class="sxs-lookup"><span data-stu-id="ff647-291">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Учебник Excel - Закрепление заголовка](../images/excel-tutorial-freeze-header.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="ff647-293">Защита листа</span><span class="sxs-lookup"><span data-stu-id="ff647-293">Protect a worksheet</span></span>

<span data-ttu-id="ff647-294">На данном этапе, описанном в руководстве, вы добавите на ленту еще одну кнопку, при нажатии которой будет выполнена определенная вами функция включения или выключения защиты листа.</span><span class="sxs-lookup"><span data-stu-id="ff647-294">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="ff647-295">Настройка манифеста для добавления второй кнопки на ленту</span><span class="sxs-lookup"><span data-stu-id="ff647-295">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="ff647-296">Откройте файл манифеста my-office-add-in-manifest.xml.</span><span class="sxs-lookup"><span data-stu-id="ff647-296">Open the manifest file my-office-add-in-manifest.xml.</span></span>

2. <span data-ttu-id="ff647-p134">Найдите элемент `<Control>`. Этот элемент определяет кнопку **Show Taskpane** (Показать область задач) на вкладке **Главная**, которую вы используете для запуска надстройки. Мы добавим вторую кнопку в эту же группу на ленте **Главная**. Добавьте приведенный ниже код между закрывающим тегом элемента управления (`</Control>`) и закрывающим тегом группы (`</Group>`).</span><span class="sxs-lookup"><span data-stu-id="ff647-p134">Find the `<Control>` element. This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in. We're going to add a second button to the same group on the **Home** ribbon. In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="ff647-301">Замените `TODO1` строкой, которая присваивает кнопке идентификатор, уникальный в пределах этого файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="ff647-301">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="ff647-302">Так как кнопка будет включать и выключать защиту листа, укажите "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="ff647-302">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="ff647-303">Когда сделаете это, весь открывающий тег элемента управления должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="ff647-303">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="ff647-p136">Следующие три элемента `TODO` устанавливают "resid", или идентификаторы ресурса. Ресурс должен быть строкой, и вы создадите эти три строки на следующем этапе. Сейчас вам нужно присвоить идентификаторы ресурсам. Кнопка должна называться "Toggle Protection" (Переключение защиты), но у строки должен быть *идентификатор* "ProtectionButtonLabel", поэтому готовый элемент `Label` выглядит следующим образом:</span><span class="sxs-lookup"><span data-stu-id="ff647-p136">The next three `TODO`s set "resid"s, which is short for resource ID. A resource is a string, and you'll create these three strings in a later step. For now, you need to give IDs to the resources. The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="ff647-p137">Элемент `SuperTip` определяет подсказку для кнопки. Заголовок этой подсказки должен совпадать с названием кнопки, поэтому мы используем тот же ИД ресурса — "ProtectionButtonLabel". Описание подсказки будет следующим: "Click to turn protection of the worksheet on and off" (Нажмите для включения или выключения защиты листа). У `ID` должно быть значение "ProtectionButtonToolTip". После выполнения весь код `SuperTip` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="ff647-p137">The `SuperTip` element defines the tool tip for the button. The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel". The tool tip description will be "Click to turn protection of the worksheet on and off". But the `ID` should be "ProtectionButtonToolTip". So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="ff647-p138">В рабочей надстройке не нужно использовать один и тот же значок для двух разных кнопок, но сейчас мы предлагаем сделать это для простоты. Поэтому код `Icon` в новом теге `Control` представляет собой лишь копию элемента `Icon` из существующего тега `Control`.</span><span class="sxs-lookup"><span data-stu-id="ff647-p138">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="ff647-p139">Для элемента `Action` в исходном элементе `Control`, уже присутствующем в манифесте, задан тип `ShowTaskpane`, но новая кнопка будет не открывать область задач, а выполнять специальную функцию, которую вы создадите позже. Поэтому замените `TODO5` на `ExecuteFunction`(тип действия для кнопок, запускающих специальные функции). Открывающий тег `Action` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="ff647-p139">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step. So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions. The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="ff647-p140">У исходного элемента `Action` есть дочерние элементы, определяющие идентификатор области задач и URL-адрес страницы, которая должна быть открыта в области задач. Но у элемента `Action` типа `ExecuteFunction` есть один дочерний элемент, который именует функцию, выполняемую элементом управления. На более позднем этапе вы создадите функцию `toggleProtection`. Поэтому замените `TODO6` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="ff647-p140">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="ff647-322">Теперь весь код `Control` должен выглядеть вот так:</span><span class="sxs-lookup"><span data-stu-id="ff647-322">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="ff647-323">Прокрутите страницу вниз до раздела `Resources` манифеста.</span><span class="sxs-lookup"><span data-stu-id="ff647-323">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="ff647-324">Добавьте приведенный ниже код в качестве дочернего элемента `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="ff647-324">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="ff647-325">Добавьте приведенный ниже код в качестве дочернего элемента `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="ff647-325">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="ff647-326">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="ff647-326">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="ff647-327">Создание функции защиты листа</span><span class="sxs-lookup"><span data-stu-id="ff647-327">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="ff647-328">Откройте файл \function-file\function-file.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-328">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="ff647-329">В файле уже есть функция-выражение, вызываемая сразу после создания (IIFE).</span><span class="sxs-lookup"><span data-stu-id="ff647-329">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="ff647-330">Добавьте в *иифе*следующий код.</span><span class="sxs-lookup"><span data-stu-id="ff647-330">*Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="ff647-331">Обратите внимание на то, что мы указываем параметр `args` для метода, а самая последняя строка метода вызывает `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="ff647-331">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="ff647-332">Это требование для всех команд надстройки типа **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="ff647-332">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="ff647-333">Это сигнализирует ведущему приложению Office о том, что работа функции завершена и пользовательский интерфейс снова может реагировать.</span><span class="sxs-lookup"><span data-stu-id="ff647-333">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

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

3. <span data-ttu-id="ff647-p142">Замените `TODO1` приведенным ниже кодом. В этом коде используется свойство защиты объекта листа в стандартном шаблоне переключателя. Объяснение `TODO2` будет приведено в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="ff647-p142">Replace `TODO1` with the following code. This code uses the worksheet object's protection property in a standard toggle pattern. The `TODO2` will be explained in the next section.</span></span>

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="ff647-337">Добавление кода для получения свойств документа в объекты скрипта области задач</span><span class="sxs-lookup"><span data-stu-id="ff647-337">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="ff647-p143">В случае всех описанных ранее функций из этой серии руководств вы ставили в очередь команды для *записи* данных в документ Office. Каждая функция заканчивалась вызовом метода `context.sync()`, который отправляет выставленные в очередь команды документу для выполнения. Но код, который вы добавили на последнем этапе, вызывает свойство `sheet.protection.protected`, и в этом заключается существенное отличие от ранее написанных функций, так как `sheet` является лишь объектом прокси, существующим в скрипте вашей области задач. В нем нет сведений о фактическом состоянии защиты документа, поэтому его свойство `protection.protected` не может иметь реального значения. Сначала нужно получить сведения о состоянии защиты от документа и задать значение `sheet.protection.protected`, используя их. Только после этого станет возможным вызов `sheet.protection.protected` без исключения. Процесс получения делится на три этапа:</span><span class="sxs-lookup"><span data-stu-id="ff647-p143">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed. But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script. It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value. It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`. Only then can `sheet.protection.protected` be called without causing an exception to be thrown. This fetching process has three steps:</span></span>

   1. <span data-ttu-id="ff647-345">Добавление в очередь команды для загрузки (т. е. получения) свойств, которые должен прочесть ваш код.</span><span class="sxs-lookup"><span data-stu-id="ff647-345">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="ff647-346">Вызов метода `sync` объекта контекста, чтобы можно было отправить документу находящуюся в очереди команду для выполнения, а также для возврата запрошенных данных.</span><span class="sxs-lookup"><span data-stu-id="ff647-346">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="ff647-347">Метод `sync` асинхронный, поэтому его выполнение должно быть завершено до того, как код вызовет полученные свойства.</span><span class="sxs-lookup"><span data-stu-id="ff647-347">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="ff647-348">Эти три действия должны выполняться каждый раз, когда коду нужно *прочесть* данные из документа Office.</span><span class="sxs-lookup"><span data-stu-id="ff647-348">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="ff647-p144">В функции `toggleProtection` замените `TODO2` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ff647-p144">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   
   - <span data-ttu-id="ff647-p145">У каждого объекта Excel есть метод `load`. Вы указываете свойства объекта, которые нужно прочесть в параметре как строку имен, разделенных запятыми. В этом случае нужно прочесть подсвойство свойства `protection`. На подсвойство нужно ссылаться почти так же, как и в остальных частях кода. Отличие заключается в том, что вместо символа "." нужно указать косую черту ("/").</span><span class="sxs-lookup"><span data-stu-id="ff647-p145">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="ff647-355">Чтобы логика переключения, которая считывает `sheet.protection.protected`, не срабатывала до выполнения `sync` и присвоения `sheet.protection.protected` правильного значения, полученного из документа, она будет перемещена (на следующем этапе) в функцию `then`, которая не выполняется до завершения `sync`.</span><span class="sxs-lookup"><span data-stu-id="ff647-355">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

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

2. <span data-ttu-id="ff647-p146">Для двух операторов `return` не может использоваться один путь кода, который не разветвляется, поэтому удалите последнюю строку `return context.sync();` в конце `Excel.run`. Вы добавите новую последнюю строку `context.sync` позже.</span><span class="sxs-lookup"><span data-stu-id="ff647-p146">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="ff647-358">Вырежьте структуру `if ... else` в функции `toggleProtection` и вставьте вместо `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="ff647-358">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="ff647-p147">Замените `TODO4` приведенным ниже кодом. Примечание:</span><span class="sxs-lookup"><span data-stu-id="ff647-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="ff647-361">Благодаря тому, что метод `sync` передается функции `then`, он не будет запускаться до добавления `sheet.protection.unprotect()` или `sheet.protection.protect()` в очередь.</span><span class="sxs-lookup"><span data-stu-id="ff647-361">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="ff647-362">Метод `then` вызывает любую функцию, которая ему передана. Не нужно вызывать `sync` дважды, поэтому уберите "()" после `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="ff647-362">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="ff647-363">Когда все будет готово, функция должна выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="ff647-363">When you are done, the entire function should look like the following:</span></span>

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

### <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="ff647-364">Настройка HTML-файла для загрузки скрипта</span><span class="sxs-lookup"><span data-stu-id="ff647-364">Configure the script-loading HTML file</span></span>

<span data-ttu-id="ff647-p148">Откройте файл /function-file/function-file.html. Это HTML-файл без пользовательского интерфейса, вызываемый, когда пользователь нажимает кнопку **Toggle Worksheet Protection** (Переключение защиты листа). Он предназначен для загрузки метода JavaScript, который должен выполняться при нажатии кнопки. Вы не будете изменять этот файл. Просто обратите внимание на то, что второй тег `<script>` загружает functionfile.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-p148">Open the /function-file/function-file.html file. This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button. Its purpose is to load the JavaScript method that should run when the button is pushed. You are not going to change this file. Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="ff647-p149">Файл function-file.html и загружаемый им файл function-file.js выполняются в полностью отдельном процессе IE из области задач надстройки. Если файл function-file.js был передан в тот же файл bundle.js, что и файл app.js, надстройка загрузит два экземпляра файла bundle.js, и это отменяет цель объединения. Кроме того, файл function-file.js не содержит код JavaScript, который не поддерживается в IE. По этим двум причинам такая надстройка не передает файл function-file.js вообще.</span><span class="sxs-lookup"><span data-stu-id="ff647-p149">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane. If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling. In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE. For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

### <a name="test-the-add-in"></a><span data-ttu-id="ff647-374">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="ff647-374">Test the add-in</span></span>

1. <span data-ttu-id="ff647-375">Закройте все приложения Office, в том числе Excel.</span><span class="sxs-lookup"><span data-stu-id="ff647-375">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="ff647-p150">Очистите кэш Office, удалив содержимое папки кэша. Это необходимо, чтобы можно было полностью удалить старую версию надстройки из ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="ff647-p150">Delete the Office cache by deleting the contents of the cache folder. This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="ff647-378">Для Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="ff647-378">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="ff647-379">Для Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="ff647-379">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

3. <span data-ttu-id="ff647-p151">Если по той или иной причине ваш сервер не работает, в окне Git Bash или системной командной строке с поддержкой Node.JS перейдите к папке **Start** проекта и выполните команду `npm start`. Повторную сборку проекта выполнять не нужно, так как единственный файл JavaScript, который вы изменили, не относится к сборке bundle.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-p151">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`. You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>

4. <span data-ttu-id="ff647-p152">Используя новую версию измененного файла манифеста, повторите процесс загрузки неопубликованного приложения с помощью одного из указанных далее методов. *Нужно перезаписать предыдущий экземпляр файла манифеста.*</span><span class="sxs-lookup"><span data-stu-id="ff647-p152">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods. *You should overwrite the previous copy of the manifest file.*</span></span>

    - <span data-ttu-id="ff647-384">[Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="ff647-384">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="ff647-385">[Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="ff647-385">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="ff647-386">[iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="ff647-386">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="ff647-387">Откройте любой лист в Excel.</span><span class="sxs-lookup"><span data-stu-id="ff647-387">Open any worksheet in Excel.</span></span>

6. <span data-ttu-id="ff647-p153">На ленте **Главная** нажмите кнопку **Toggle Worksheet Protection** (Переключение защиты листа). Обратите внимание на то, что большинство элементов управления на ленте отключены (серые), как показано на приведенном ниже снимке экрана.</span><span class="sxs-lookup"><span data-stu-id="ff647-p153">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 

7. <span data-ttu-id="ff647-p154">Выберите ячейку, как если бы вы хотели изменить ее содержимое. Появится сообщение об ошибке и защите листа.</span><span class="sxs-lookup"><span data-stu-id="ff647-p154">Choose a cell as you would if you wanted to change its content. You get an error telling you that the worksheet is protected.</span></span>

8. <span data-ttu-id="ff647-392">Нажмите кнопку **Toggle Worksheet Protection** (Переключение защиты листа) еще раз, и элементы управления включатся, после чего вы сможете изменить значения ячеек.</span><span class="sxs-lookup"><span data-stu-id="ff647-392">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Руководство по Excel: лента с включенной защитой](../images/excel-tutorial-ribbon-with-protection-on.png)

## <a name="open-a-dialog"></a><span data-ttu-id="ff647-394">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="ff647-394">Open a dialog</span></span>

<span data-ttu-id="ff647-p155">На данном заключительном этапе, указанном в руководстве, вы откроете диалоговое окно в своей надстройке, передадите сообщение из процесса диалогового окна в процесс области задач и закроете диалоговое окно. Диалоговые окна надстройки Office *не модальные*: пользователь может продолжать работать и с документом в ведущем приложении Office, и с главной страницей в области задач.</span><span class="sxs-lookup"><span data-stu-id="ff647-p155">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="ff647-397">Создание страницы диалогового окна</span><span class="sxs-lookup"><span data-stu-id="ff647-397">Create the dialog page</span></span>

1. <span data-ttu-id="ff647-398">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="ff647-398">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff647-399">Создайте в корневой папке проекта (где находится index.html) файл popup.html.</span><span class="sxs-lookup"><span data-stu-id="ff647-399">Create a file in the root of the project (where index.html is) called popup.html.</span></span>

3. <span data-ttu-id="ff647-p156">Добавьте в файл popup.html приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ff647-p156">Add the following markup to popup.html. Note:</span></span>

   - <span data-ttu-id="ff647-402">На странице находится `<input>`, где пользователь будет вводить свое имя, и кнопка, при нажатии которой имя будет отправлено на страницу области задач, где оно отобразится.</span><span class="sxs-lookup"><span data-stu-id="ff647-402">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="ff647-403">Код загружает скрипт под названием popup.js, который будет создан на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="ff647-403">The markup loads a script called popup.js that you will create in a later step.</span></span>

   - <span data-ttu-id="ff647-404">Он загружает также библиотеку Office.JS и jQuery, так как они будут использоваться в popup.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-404">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css" />

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <div class="padding">
                <p class="ms-font-xl">ENTER YOUR NAME</p>
            </div>
            <div class="padding">
                <input id="name-box" type="text"/>
            </div>
            <div class="padding">
                <button id="ok-button" class="ms-Button">OK</button>
            </div>
        </body>
    </html>
    ```

4. <span data-ttu-id="ff647-405">Создайте в корневой папке проекта файл popup.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-405">Create a file in the root of the project called popup.js.</span></span>

5. <span data-ttu-id="ff647-406">Добавьте указанный ниже код в файл popup.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-406">Add the following code to popup.js.</span></span> <span data-ttu-id="ff647-407">Обратите внимание на указанные ниже особенности этого кода.</span><span class="sxs-lookup"><span data-stu-id="ff647-407">Note the following about this code:</span></span>

   - <span data-ttu-id="ff647-408">*Каждая страница, вызывающая API в библиотеке Office.JS, должна сначала убедиться, что библиотека полностью инициализирована.*</span><span class="sxs-lookup"><span data-stu-id="ff647-408">*Every page that calls APIs in the Office.JS library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="ff647-409">Лучший способ сделать это — вызвать метод `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="ff647-409">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="ff647-410">Если у вашей надстройки есть собственные задачи инициализации, код должен перейти к методу `then()`, связанному с вызовом `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="ff647-410">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="ff647-411">Файл app.js в корневом каталоге проекта можно рассматривать как пример.</span><span class="sxs-lookup"><span data-stu-id="ff647-411">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="ff647-412">Вызов метода `Office.onReady()` должен выполняться до каких-либо вызовов Office.JS, поэтому назначение указано в файле скрипта, загружаемом страницей, как в этом случае.</span><span class="sxs-lookup"><span data-stu-id="ff647-412">The call of `Office.onReady()` must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   - <span data-ttu-id="ff647-413">Функция jQuery `ready` вызывается в методе `then()`.</span><span class="sxs-lookup"><span data-stu-id="ff647-413">The jQuery `ready` function is called inside the `then()` method.</span></span> <span data-ttu-id="ff647-414">В большинстве случаев код загрузки (в том числе начальной) или инициализации из других библиотек JavaScript должен находиться в методе `then()`, связанном с вызовом `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="ff647-414">In most cases, the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `then()` method that is chained to the call of `Office.onReady()`.</span></span>

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {
                $(document).ready(function () {  

                    // TODO1: Assign handler to the OK button.

                });
            });

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="ff647-p160">Замените `TODO1` приведенным ниже кодом. Вы создадите функцию `sendStringToParentPage` на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="ff647-p160">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="ff647-p161">Замените `TODO2` приведенным ниже кодом. Метод `messageParent` передает свой параметр родительской странице (в данном случае это страница на панели задач). Параметр может быть логическим или строковым. Во втором случае подразумевается все, что можно сериализовать, представив в виде строки (например, XML или JSON).</span><span class="sxs-lookup"><span data-stu-id="ff647-p161">Replace `TODO2` with the following code. The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane. The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="ff647-420">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="ff647-420">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="ff647-p162">Файл popup.html и загружаемый им файл popup.js выполняются в полностью отдельном процессе Internet Explorer из области задач надстройки. Если файл popup.js был передан в тот же файл bundle.js, что и файл app.js, надстройка загрузит два экземпляра файла bundle.js, и это отменяет цель объединения. Кроме того, файл popup.js не содержит код JavaScript, который не поддерживается в IE. По этим двум причинам эта надстройка не передает файл popup.js вообще.</span><span class="sxs-lookup"><span data-stu-id="ff647-p162">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane. If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling. In addition, the popup.js file does not contain any JavaScript that is unsupported by IE. For these two reasons, this add-in does not transpile the popup.js file at all.</span></span>

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="ff647-425">Открытие диалогового окна из области задач</span><span class="sxs-lookup"><span data-stu-id="ff647-425">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="ff647-426">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="ff647-426">Open the file index.html.</span></span>

2. <span data-ttu-id="ff647-427">Под `div` с кнопкой `freeze-header` добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="ff647-427">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. <span data-ttu-id="ff647-p163">В диалоговом окне пользователю будет предложено ввести имя и передать имя пользователя в область задач. Область задач отобразит его в подписи. Непосредственно под только что добавленным тегом `div` добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="ff647-p163">The dialog will prompt the user to enter a name and pass the user's name to the task pane. The task pane will display it in a label. Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. <span data-ttu-id="ff647-431">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="ff647-431">Open the app.js file.</span></span>

5. <span data-ttu-id="ff647-p164">Под строкой, назначающей обработчик щелчков для кнопки `freeze-header`, добавьте приведенный ниже код. Вы создадите метод `openDialog` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="ff647-p164">Below the line that assigns a click handler to the `freeze-header` button, add the following code. You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="ff647-p165">Под функцией `freezeHeader` добавьте указанное ниже объявление. Эта переменная удерживает объект в контексте выполнения родительской страницы, который служит посредником для контекста выполнения страницы диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="ff647-p165">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="ff647-p166">Добавьте приведенную ниже функцию под объявлением `dialog`. Важно отметить, что в этом коде *отсутствует* вызов `Excel.run`. Это связано с тем, что API, открывающий диалоговое окно, совместно используется всеми ведущими приложениями Office, поэтому относится к общему API JavaScript для Office, а не API для Excel.</span><span class="sxs-lookup"><span data-stu-id="ff647-p166">Below the declaration of `dialog`, add the following function. The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`. This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="ff647-439">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="ff647-439">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="ff647-440">Примечание:</span><span class="sxs-lookup"><span data-stu-id="ff647-440">Note:</span></span>

   - <span data-ttu-id="ff647-441">Метод `displayDialogAsync` открывает диалоговое окно в центре экрана.</span><span class="sxs-lookup"><span data-stu-id="ff647-441">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="ff647-442">Первый параметр — это URL-адрес открываемой страницы.</span><span class="sxs-lookup"><span data-stu-id="ff647-442">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="ff647-p168">Второй параметр передает параметры. `height` и `width` — процентные значения размера окна для приложения Office.</span><span class="sxs-lookup"><span data-stu-id="ff647-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="ff647-445">Обработка сообщения из диалогового окна и закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="ff647-445">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="ff647-p169">Продолжайте работать в файле app.js. Замените `TODO2` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="ff647-p169">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="ff647-448">Обратный вызов выполняется сразу же после успешного открытия диалогового окна и до того, как пользователь предпримет какие-либо действия в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="ff647-448">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="ff647-449">`result.value` — это объект, который выступает в качестве посредника между контекстами выполнения родительских страниц и страниц диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="ff647-449">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="ff647-p170">Функция `processMessage` будет создана на более позднем этапе. Этот обработчик будет обрабатывать любые значения, которые отправляются со страницы диалогового окна с вызовами функции `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="ff647-p170">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="ff647-452">Добавьте указанную ниже функцию под функцией `openDialog`.</span><span class="sxs-lookup"><span data-stu-id="ff647-452">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff647-453">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="ff647-453">Test the add-in</span></span>

1. <span data-ttu-id="ff647-454">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши **Ctrl+C**, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="ff647-454">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="ff647-455">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="ff647-455">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff647-p172">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу. Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки. После сборки необходимо перезапустить сервер. Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="ff647-p172">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ff647-460">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="ff647-460">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff647-461">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="ff647-461">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff647-462">Повторно загрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Show Taskpane** (Показать область задач) для повторного открытия надстройки.</span><span class="sxs-lookup"><span data-stu-id="ff647-462">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="ff647-463">Нажмите кнопку **Open Dialog** (Открыть диалоговое окно) в области задач.</span><span class="sxs-lookup"><span data-stu-id="ff647-463">Choose the **Open Dialog** button in the task pane.</span></span>

6. <span data-ttu-id="ff647-464">Когда диалоговое окно открыто, перетащите его и измените его размер.</span><span class="sxs-lookup"><span data-stu-id="ff647-464">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="ff647-465">Обратите внимание, что вы можете взаимодействовать с листом и нажимать другие кнопки в области задач, но вы не можете запустить второе диалоговое окно на одной и той же странице панели задач.</span><span class="sxs-lookup"><span data-stu-id="ff647-465">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

7. <span data-ttu-id="ff647-p174">В диалоговом окне введите имя и нажмите кнопку **OK**. В области задач отобразится имя, и диалоговое окно закроется.</span><span class="sxs-lookup"><span data-stu-id="ff647-p174">In the dialog, enter a name and choose **OK**. The name appears on the task pane and the dialog closes.</span></span>

8. <span data-ttu-id="ff647-p175">При желании можно закомментировать строку `dialog.close();` в функции `processMessage`. Повторите шаги этого раздела. Диалоговое окно остается открытым, и вы можете изменить имя. Можно закрыть его вручную, нажав кнопку **X** в правом верхнему углу.</span><span class="sxs-lookup"><span data-stu-id="ff647-p175">Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Руководство по Excel - Диалоговое окно](../images/excel-tutorial-dialog-open.png)

## <a name="next-steps"></a><span data-ttu-id="ff647-473">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="ff647-473">Next steps</span></span>

<span data-ttu-id="ff647-474">В этом руководстве показано создание надстройки Excel для области задач, которая взаимодействует с таблицами, диаграммами, листами, диалоговыми окнами в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="ff647-474">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="ff647-475">Чтобы узнать больше о создании надстроек Excel, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="ff647-475">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="ff647-476">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="ff647-476">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)

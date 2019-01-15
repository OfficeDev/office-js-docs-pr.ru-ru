---
title: Руководство по надстройкам Excel
description: В этом руководстве показана разработка надстройки Excel, которая создает, заполняет, фильтрует и сортирует данные таблиц, создает диаграммы, закрепляет заголовки таблиц, защищает листы и открывает диалоговые окна.
ms.date: 01/09/2019
ms.topic: tutorial
ms.openlocfilehash: de5a08be53d7a6c2f4df4d9419e3713266800f7e
ms.sourcegitcommit: 384e217fd51d73d13ccfa013bfc6e049b66bd98c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/11/2019
ms.locfileid: "27896359"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="2d7b3-103">Учебник: Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="2d7b3-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="2d7b3-104">С помощью данного учебника вы сможете создать надстройку области задач Excel, которая выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="2d7b3-105">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-105">Creates a table</span></span>
> * <span data-ttu-id="2d7b3-106">Фильтрация и сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="2d7b3-107">Создание графика</span><span class="sxs-lookup"><span data-stu-id="2d7b3-107">Creates a chart</span></span>
> * <span data-ttu-id="2d7b3-108">Закрепление заголовка таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-108">Freezes a table header</span></span>
> * <span data-ttu-id="2d7b3-109">Защита листа</span><span class="sxs-lookup"><span data-stu-id="2d7b3-109">Protects a worksheet</span></span>
> * <span data-ttu-id="2d7b3-110">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="2d7b3-110">Opens a dialog</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2d7b3-111">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="2d7b3-111">Prerequisites</span></span>

<span data-ttu-id="2d7b3-112">Для работы с этим учебником необходимо установить указанные ниже компоненты.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-112">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="2d7b3-113">Excel 2016, версия 1711 (сборка 8730.1000 "нажми и работай") или более поздняя.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-113">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="2d7b3-114">Чтобы установить эту версию, необходимо быть участником программы предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-114">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="2d7b3-115">[Дополнительные сведения](https://products.office.com/office-insider?tab=tab-1)</span><span class="sxs-lookup"><span data-stu-id="2d7b3-115">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="2d7b3-116">Node</span><span class="sxs-lookup"><span data-stu-id="2d7b3-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="2d7b3-117">[Git Bash](https://git-scm.com/downloads) (или другой клиент Git)</span><span class="sxs-lookup"><span data-stu-id="2d7b3-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="2d7b3-118">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="2d7b3-118">Create your add-in project</span></span>

<span data-ttu-id="2d7b3-119">Выполните указанные ниже действия для создания проекта надстройки Excel, который будет использоваться в качестве основы для этого учебника.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-119">Complete the following steps to create the Excel add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="2d7b3-120">Клонируйте репозиторий GitHub [Excel Add-in Tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-120">Clone the GitHub repository [Excel add-in tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="2d7b3-121">Откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-121">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="2d7b3-122">Выполните команду `npm install`, чтобы установить инструменты и библиотеки, указанные в файле package.json.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-122">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="2d7b3-123">Сделайте так, чтобы операционная система компьютера разработки доверяла сертификату. Для этого выполните действия, описанные в [этой статье](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-123">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="create-a-table"></a><span data-ttu-id="2d7b3-124">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-124">Create a table</span></span>

<span data-ttu-id="2d7b3-125">На этом этапе руководства мы проверим программным способом, поддерживает ли надстройка текущую версию Excel, установленную у пользователя, а также добавим таблицу на лист, заполним ее данными и отформатируем.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-125">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="2d7b3-126">Написание кода надстройки</span><span class="sxs-lookup"><span data-stu-id="2d7b3-126">Code the add-in</span></span>

1. <span data-ttu-id="2d7b3-127">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-127">Open the project in your code editor.</span></span>

2. <span data-ttu-id="2d7b3-128">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-128">Open the file index.html.</span></span>

3. <span data-ttu-id="2d7b3-129">Замените `TODO1` на следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-129">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="2d7b3-130">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-130">Open the app.js file.</span></span>

5. <span data-ttu-id="2d7b3-131">Замените `TODO1` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-131">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="2d7b3-132">Этот код определяет, поддерживает ли установленная у пользователя версия Excel ту версию файла Excel.js, которая включает все API, используемые в этой серии руководств.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-132">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="2d7b3-133">В рабочей надстройке можно использовать текст условного блока, чтобы скрыть или отключить пользовательский интерфейс, где вызываются неподдерживаемые API.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-133">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="2d7b3-134">При этом пользователь по-прежнему сможет использовать те части надстройки, которые поддерживаются в его версии Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-134">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="2d7b3-135">Замените `TODO2` на следующий код:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-135">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="2d7b3-136">Замените `TODO3` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-136">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="2d7b3-137">Примечание.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-137">Note:</span></span>

   - <span data-ttu-id="2d7b3-138">Бизнес-логика Excel.js будет добавлена в функцию, передаваемую методу `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-138">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="2d7b3-139">Эта логика выполняется не сразу.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-139">This logic does not execute immediately.</span></span> <span data-ttu-id="2d7b3-140">Вместо этого она добавляется в очередь ожидания команд.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-140">Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="2d7b3-141">Метод `context.sync` отправляет все команды из очереди в Excel для выполнения.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-141">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

   - <span data-ttu-id="2d7b3-142">За методом `Excel.run` следует блок `catch`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-142">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="2d7b3-143">Рекомендуется всегда следовать этой методике.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-143">This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="2d7b3-p106">Замените `TODO4` на приведенный ниже код. Примечание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p106">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-146">код создает таблицу с помощью метода `add` коллекции таблиц на листе, которая всегда существует, даже если она пуста.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-146">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="2d7b3-147">Это стандартный способ создания объектов Excel.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-147">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="2d7b3-148">API конструкторов классов не существуют, а для создания объекта Excel никогда не следует использовать оператор `new`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-148">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="2d7b3-149">Вместо этого следует добавить его к объекту родительской коллекции.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-149">Instead, you add to a parent collection object.</span></span>

   - <span data-ttu-id="2d7b3-150">Первый параметр метода `add` — это диапазон, содержащий только первую строку, а не весь диапазон таблицы, который мы в конечном итоге будем использовать.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-150">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="2d7b3-151">Это связано с тем, что при заполнении строк данных (на следующем этапе) надстройка добавляет к таблице новые строки, а не записывает их в ячейки имеющихся строк.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-151">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="2d7b3-152">Такой шаблон более распространен, так как количество строк в таблице часто неизвестно на момент ее создания.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-152">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

   - <span data-ttu-id="2d7b3-153">Имена таблиц должны быть уникальными в рамках всей книги, а не только одного листа.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-153">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="2d7b3-p109">Замените `TODO5` на приведенный ниже код. Примечание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p109">Replace `TODO5` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-156">значения ячеек диапазона задаются с помощью массива массивов.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-156">The cell values of a range are set with an array of arrays.</span></span>

   - <span data-ttu-id="2d7b3-157">Новые строки создаются в таблице путем вызова метода `add` коллекции ее строк.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-157">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="2d7b3-158">Вы можете добавить несколько строк в одном вызове метода `add`, включив несколько массивов значений ячеек в родительский массив, передаваемый в качестве второго параметра.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-158">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="2d7b3-p111">Замените `TODO6` на приведенный ниже код. Примечание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p111">Replace `TODO6` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-161">код получает ссылку на столбец **Сумма**, передавая его индекс (с отсчетом от нуля) в метод `getItemAt` коллекции столбцов таблицы.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-161">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="2d7b3-162">У объектов коллекций Excel.js (например, `TableCollection`, `WorksheetCollection` и `TableColumnCollection`) есть свойство `items`, представляющее собой массив дочерних типов объектов (например, `Table`, `Worksheet` или `TableColumn`). Однако сам объект `*Collection` не является массивом.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-162">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="2d7b3-163">Затем код форматирует диапазон столбца **Сумма** как денежные суммы в евро с точностью до второго знака после запятой.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-163">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

   - <span data-ttu-id="2d7b3-164">Напоследок он обеспечивает достаточные ширину столбцов и высоту строк для размещения самого длинного (или самого высокого) элемента данных.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-164">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="2d7b3-165">Обратите внимание, что код должен привести объекты `Range` к нужному формату.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-165">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="2d7b3-166">У объектов `TableColumn` и `TableRow` нет свойств формата.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-166">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

### <a name="test-the-add-in"></a><span data-ttu-id="2d7b3-167">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="2d7b3-167">Test the add-in</span></span>

1. <span data-ttu-id="2d7b3-168">Откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-168">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="2d7b3-169">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-169">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="2d7b3-170">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-170">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="2d7b3-171">Загрузите неопубликованную надстройку одним из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-171">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="2d7b3-172">Windows: [загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="2d7b3-172">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="2d7b3-173">[Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="2d7b3-173">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="2d7b3-174">iPad и Mac: [загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="2d7b3-174">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="2d7b3-175">В меню **Главная** выберите пункт **Показать область задач**.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-175">On the **Home** menu, choose **Show Taskpane**.</span></span>

6. <span data-ttu-id="2d7b3-176">В области задач нажмите кнопку **Create Table** (Создать таблицу).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-176">In the task pane, choose **Create Table**.</span></span>

    ![Руководство по Excel: создание таблицы](../images/excel-tutorial-create-table.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="2d7b3-178">Фильтрация и сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-178">Filter and sort a table</span></span>

<span data-ttu-id="2d7b3-179">Из этого раздела руководства вы узнаете, как отфильтровать и отсортировать созданную ранее таблицу.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-179">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="2d7b3-180">Фильтрация таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-180">Filter the table</span></span>

1. <span data-ttu-id="2d7b3-181">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-181">Open the project in your code editor.</span></span>

2. <span data-ttu-id="2d7b3-182">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-182">Open the file index.html.</span></span>

3. <span data-ttu-id="2d7b3-183">Под элементом `div`, содержащим кнопку `create-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-183">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. <span data-ttu-id="2d7b3-184">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-184">Open the app.js file.</span></span>

5. <span data-ttu-id="2d7b3-185">Под строкой, назначающей обработчик нажатия кнопки `create-table`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-185">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="2d7b3-186">Под функцией `createTable` добавьте следующую функцию:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-186">Just below the `createTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="2d7b3-p113">Замените `TODO1` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p113">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-189">Код получает ссылку на столбец, который нужно отфильтровать, передавая имя столбца методу `getItem`, а не передавая его индекс методу `getItemAt`, как это делает метод `createTable`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-189">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does.</span></span> <span data-ttu-id="2d7b3-190">Так как пользователи могут перемещать столбцы, по заданному индексу может располагаться уже другой столбец.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-190">Since users can move table columns, the column at a given index might change after the table is created.</span></span> <span data-ttu-id="2d7b3-191">Следовательно, для получения ссылки безопаснее использовать имя столбца.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-191">Hence, it is safer to use the column name to get a reference to the column.</span></span> <span data-ttu-id="2d7b3-192">Мы спокойно использовали метод `getItemAt` в предыдущем разделе, потому что мы использовали его в методе, который создает таблицу, и пользователь никак не мог переместить столбец.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-192">We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="2d7b3-193">Метод `applyValuesFilter` является одним из нескольких методов фильтрации объекта `Filter`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-193">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="2d7b3-194">Сортировка таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-194">Sort the table</span></span>

1. <span data-ttu-id="2d7b3-195">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-195">Open the file index.html.</span></span>

2. <span data-ttu-id="2d7b3-196">Под элементом `div`, содержащим кнопку `filter-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-196">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. <span data-ttu-id="2d7b3-197">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-197">Open the app.js file.</span></span>

4. <span data-ttu-id="2d7b3-198">Под строкой, назначающей обработчик нажатия кнопки `filter-table`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-198">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="2d7b3-199">Под функцией `filterTable` добавьте приведенную ниже функцию.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-199">Below the `filterTable` function add the following function.</span></span>

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

6. <span data-ttu-id="2d7b3-p115">Замените `TODO1` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p115">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-202">Код создает массив объектов `SortField`, состоящий из одного элемента, так как надстройка сортирует таблицу только по столбцу Merchant.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-202">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="2d7b3-203">Свойство `key` объекта `SortField` — это отсчитываемый от нуля индекс столбца, по которому необходимо сортировать таблицу.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-203">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="2d7b3-204">Элемент `sort` объекта `Table` — это объект `TableSort`, а не метод.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-204">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="2d7b3-205">Объекты `SortField` передаются методу `apply` объекта `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-205">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

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

### <a name="test-the-add-in"></a><span data-ttu-id="2d7b3-206">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="2d7b3-206">Test the add-in</span></span>

1. <span data-ttu-id="2d7b3-207">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши **Ctrl+C**, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-207">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="2d7b3-208">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-208">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="2d7b3-209">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-209">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="2d7b3-210">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-210">In order to do this, you need to kill the server process so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="2d7b3-211">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-211">After the build, you restart the server.</span></span> <span data-ttu-id="2d7b3-212">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-212">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="2d7b3-213">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-213">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="2d7b3-214">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-214">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="2d7b3-215">Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-215">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="2d7b3-216">Если по той или иной причине на открытом листе нет таблицы, нажмите в области задач кнопку **Create Table** (Создать таблицу).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-216">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table**.</span></span>

6. <span data-ttu-id="2d7b3-217">Нажмите кнопки **Filter Table** (Фильтровать таблицу) и **Sort Table** (Сортировать таблицу) в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-217">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Учебник Excel - Фильтрация и сортировка таблицы](../images/excel-tutorial-filter-and-sort-table.png)

## <a name="create-a-chart"></a><span data-ttu-id="2d7b3-219">Создание диаграммы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-219">Create a chart</span></span>

<span data-ttu-id="2d7b3-220">На этом этапе руководства мы создадим диаграмму, используя данные из ранее созданной таблицы, а затем отформатируем эту диаграмму.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-220">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="2d7b3-221">Создание диаграммы с помощью таблицы данных</span><span class="sxs-lookup"><span data-stu-id="2d7b3-221">Chart a chart using table data</span></span>

1. <span data-ttu-id="2d7b3-222">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-222">Open the project in your code editor.</span></span>

2. <span data-ttu-id="2d7b3-223">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-223">Open the file index.html.</span></span>

3. <span data-ttu-id="2d7b3-224">Под элементом `div`, содержащим кнопку `sort-table`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-224">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="2d7b3-225">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-225">Open the app.js file.</span></span>

5. <span data-ttu-id="2d7b3-226">Под строкой, назначающей обработчик нажатия кнопки `sort-chart`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-226">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="2d7b3-227">Под функцией `sortTable` добавьте приведенную ниже функцию.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-227">Below the `sortTable` function add the following function.</span></span>

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

7. <span data-ttu-id="2d7b3-p119">Замените `TODO1` приведенным ниже кодом. Обратите внимание на то, что для исключения строки заголовков в коде вместо метода `getRange` используется метод `Table.getDataBodyRange`, чтобы получить нужный диапазон данных для диаграммы.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p119">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="2d7b3-230">Замените `TODO2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-230">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="2d7b3-231">Обратите внимание на следующие параметры:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-231">Note the following parameters:</span></span>

   - <span data-ttu-id="2d7b3-p121">Первый параметр метода `add` задает тип диаграммы. Существует несколько десятков типов.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p121">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="2d7b3-234">Второй параметр задает диапазон данных, включаемых в диаграмму.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-234">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="2d7b3-235">Третий параметр определяет, как следует отображать на диаграмме ряд точек данных из таблицы: по строкам или по столбцам.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-235">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="2d7b3-236">Значение `auto` сообщает Excel, что следует выбрать оптимальный способ.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-236">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="2d7b3-237">Замените `TODO3` на приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-237">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="2d7b3-238">Большая часть этого кода не требует объяснений.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-238">Most of this code is self-explanatory.</span></span> <span data-ttu-id="2d7b3-239">Примечание.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-239">Note:</span></span>
   
   - <span data-ttu-id="2d7b3-240">Параметры метода `setPosition` задают левую верхнюю и правую нижнюю ячейки области листа, которые должны содержать диаграмму.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-240">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="2d7b3-241">Excel может настраивать такие параметры, как ширина линий, чтобы диаграмма хорошо выглядела в выделенном для нее пространстве.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-241">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="2d7b3-242">"Ряд" — это набор точек данных из столбца таблицы.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-242">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="2d7b3-243">Так как в таблице есть только один нестроковый столбец, Excel делает вывод, что это единственный столбец точек данных для диаграммы.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-243">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="2d7b3-244">Он рассматривает другие столбцы как метки диаграммы.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-244">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="2d7b3-245">Следовательно, в диаграмме будет только один ряд, обозначенный индексом 0.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-245">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="2d7b3-246">К нему следует добавить метку "Значение в €".</span><span class="sxs-lookup"><span data-stu-id="2d7b3-246">This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="2d7b3-247">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="2d7b3-247">Test the add-in</span></span>

1. <span data-ttu-id="2d7b3-248">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши **Ctrl+C**, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-248">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="2d7b3-249">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-249">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="2d7b3-250">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-250">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="2d7b3-251">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-251">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="2d7b3-252">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-252">After the build, you restart the server.</span></span> <span data-ttu-id="2d7b3-253">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-253">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="2d7b3-254">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-254">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="2d7b3-255">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-255">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="2d7b3-256">Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-256">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="2d7b3-257">Если по той или иной причине на открытом листе нет таблицы, нажмите в области задач кнопку **Create Table** (Создать таблицу), а затем — кнопки **Filter Table** (Фильтровать таблицу) и **Sort Table** (Сортировать таблицу) в любом порядке.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-257">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>

6. <span data-ttu-id="2d7b3-258">Нажмите кнопку **Create Chart** (Создать диаграмму).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-258">Choose the **Create Chart** button.</span></span> <span data-ttu-id="2d7b3-259">Будет создана диаграмма, включающая только данные из отфильтрованных строк.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-259">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="2d7b3-260">Метки точек данных в нижней части диаграммы отсортированы согласно заданному для нее порядку, то есть по именам продавцов в обратном алфавитном порядке.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-260">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Руководство по Excel - Создание диаграммы](../images/excel-tutorial-create-chart.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="2d7b3-262">Закрепление заголовка таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-262">Freeze a table header</span></span>

<span data-ttu-id="2d7b3-263">Когда таблица достаточно длинная, при прокрутке строка заголовков может исчезать с экрана.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-263">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="2d7b3-264">В этом разделе учебника мы расскажем, как закрепить строку заголовков созданной ранее таблицы, чтобы она была видна, даже когда пользователь прокручивает лист.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-264">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="2d7b3-265">Закрепление строки заголовков таблицы</span><span class="sxs-lookup"><span data-stu-id="2d7b3-265">Freeze the table's header row</span></span>

1. <span data-ttu-id="2d7b3-266">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-266">Open the project in your code editor.</span></span>

2. <span data-ttu-id="2d7b3-267">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-267">Open the file index.html.</span></span>

3. <span data-ttu-id="2d7b3-268">Под элементом `div`, содержащим кнопку `create-chart`, добавьте следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-268">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. <span data-ttu-id="2d7b3-269">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-269">Open the app.js file.</span></span>

5. <span data-ttu-id="2d7b3-270">Под строкой, назначающей обработчик нажатия кнопки `create-chart`, добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-270">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="2d7b3-271">Под функцией `createChart` добавьте следующую функцию:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-271">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="2d7b3-p130">Замените `TODO1` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p130">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-274">Коллекция `Worksheet.freezePanes` — это набор закрепленных строк, которые не исчезают с экрана при прокрутке листа.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-274">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="2d7b3-p131">Метод `freezeRows` принимает в качестве параметра количество строк сверху, которые необходимо закрепить. Мы передаем значение `1`, чтобы закрепить первую строку.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p131">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="2d7b3-277">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="2d7b3-277">Test the add-in</span></span>

1. <span data-ttu-id="2d7b3-278">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши **Ctrl+C**, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-278">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="2d7b3-279">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-279">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="2d7b3-280">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-280">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="2d7b3-281">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-281">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="2d7b3-282">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-282">After the build, you restart the server.</span></span> <span data-ttu-id="2d7b3-283">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-283">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="2d7b3-284">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-284">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="2d7b3-285">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-285">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="2d7b3-286">Повторно загрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-286">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="2d7b3-287">Если таблица на листе, удалите ее.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-287">If the table is in the worksheet, delete it.</span></span>

6. <span data-ttu-id="2d7b3-288">В области задач нажмите кнопку **Create Table** (Создать таблицу).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-288">In the task pane, choose **Create Table**.</span></span>

7. <span data-ttu-id="2d7b3-289">Нажмите кнопку **Freeze Header** (Закрепить заголовок).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-289">Choose the **Freeze Header** button.</span></span>

8. <span data-ttu-id="2d7b3-290">Прокрутите лист вниз, чтобы убедиться, что заголовок таблицы по-прежнему остается на экране, даже когда более высокие строки исчезают.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-290">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Учебник Excel - Закрепление заголовка](../images/excel-tutorial-freeze-header.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="2d7b3-292">Защита листа</span><span class="sxs-lookup"><span data-stu-id="2d7b3-292">Protect a worksheet</span></span>

<span data-ttu-id="2d7b3-293">На данном этапе, описанном в руководстве, вы добавите на ленту еще одну кнопку, при нажатии которой будет выполнена определенная вами функция включения или выключения защиты листа.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-293">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="2d7b3-294">Настройка манифеста для добавления второй кнопки на ленту</span><span class="sxs-lookup"><span data-stu-id="2d7b3-294">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="2d7b3-295">Откройте файл манифеста my-office-add-in-manifest.xml.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-295">Open the manifest file my-office-add-in-manifest.xml.</span></span>

2. <span data-ttu-id="2d7b3-296">Найдите элемент `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-296">Find the `<Control>` element.</span></span> <span data-ttu-id="2d7b3-297">Этот элемент определяет кнопку **Show Taskpane** (Показать область задач) на вкладке **Главная**, которую вы используете для запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-297">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="2d7b3-298">Мы добавим вторую кнопку в эту же группу на ленте **Главная**.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-298">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="2d7b3-299">Добавьте приведенный ниже код между закрывающим тегом элемента управления (`</Control>`) и закрывающим тегом группы (`</Group>`).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-299">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

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

3. <span data-ttu-id="2d7b3-300">Замените `TODO1` строкой, которая присваивает кнопке идентификатор, уникальный в пределах этого файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-300">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="2d7b3-301">Так как кнопка будет включать и выключать защиту листа, укажите "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="2d7b3-301">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="2d7b3-302">Когда сделаете это, весь открывающий тег элемента управления должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-302">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="2d7b3-303">Следующие три элемента `TODO` устанавливают "resid", или идентификаторы ресурса.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-303">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="2d7b3-304">Ресурс должен быть строкой, и вы создадите эти три строки на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-304">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="2d7b3-305">Сейчас вам нужно присвоить идентификаторы ресурсам.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-305">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="2d7b3-306">Кнопка должна называться "Toggle Protection" (Переключение защиты), но у строки должен быть *идентификатор* "ProtectionButtonLabel", поэтому готовый элемент `Label` выглядит следующим образом:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-306">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="2d7b3-307">Элемент `SuperTip` определяет подсказку для кнопки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-307">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="2d7b3-308">Заголовок этой подсказки должен совпадать с названием кнопки, поэтому мы используем тот же ИД ресурса — "ProtectionButtonLabel".</span><span class="sxs-lookup"><span data-stu-id="2d7b3-308">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="2d7b3-309">Описание подсказки будет следующим: "Click to turn protection of the worksheet on and off" (Нажмите для включения или выключения защиты листа).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-309">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="2d7b3-310">У `ID` должно быть значение "ProtectionButtonToolTip".</span><span class="sxs-lookup"><span data-stu-id="2d7b3-310">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="2d7b3-311">После выполнения весь код `SuperTip` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-311">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="2d7b3-312">В рабочей надстройке не нужно использовать один и тот же значок для двух разных кнопок, но сейчас мы предлагаем сделать это для простоты.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-312">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="2d7b3-313">Поэтому код `Icon` в новом теге `Control` представляет собой лишь копию элемента `Icon` из существующего тега `Control`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-313">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="2d7b3-314">Для элемента `Action` в исходном элементе `Control`, уже присутствующем в манифесте, задан тип `ShowTaskpane`, но новая кнопка будет не открывать область задач, а выполнять специальную функцию, которую вы создадите позже.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-314">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="2d7b3-315">Поэтому замените `TODO5` на `ExecuteFunction` (тип действия для кнопок, запускающих специальные функции).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-315">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="2d7b3-316">Открывающий тег `Action` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-316">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="2d7b3-317">У исходного элемента `Action` есть дочерние элементы, определяющие идентификатор области задач и URL-адрес страницы, которая должна быть открыта в области задач.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-317">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="2d7b3-318">Но у элемента `Action` типа `ExecuteFunction` есть один дочерний элемент, который именует функцию, выполняемую элементом управления.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-318">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="2d7b3-319">На более позднем этапе вы создадите функцию `toggleProtection`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-319">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="2d7b3-320">Поэтому замените `TODO6` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-320">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="2d7b3-321">Теперь весь код `Control` должен выглядеть вот так:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-321">The entire `Control` markup should now look like the following:</span></span>

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

8. <span data-ttu-id="2d7b3-322">Прокрутите страницу вниз до раздела `Resources` манифеста.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-322">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="2d7b3-323">Добавьте приведенный ниже код в качестве дочернего элемента `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-323">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="2d7b3-324">Добавьте приведенный ниже код в качестве дочернего элемента `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-324">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="2d7b3-325">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-325">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="2d7b3-326">Создание функции защиты листа</span><span class="sxs-lookup"><span data-stu-id="2d7b3-326">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="2d7b3-327">Откройте файл \function-file\function-file.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-327">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="2d7b3-328">В файле уже есть функция-выражение, вызываемая сразу после создания (IIFE).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-328">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="2d7b3-329">Пользовательская логика инициализации не требуется, поэтому оставьте тело функции, назначенной `Office.initialize`, пустым.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-329">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="2d7b3-330">(Но не удаляйте его.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-330">(But do not delete it.</span></span> <span data-ttu-id="2d7b3-331">Свойство `Office.initialize` не может быть неопределенным или иметь значение NULL.) *За пределами IIFE* добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-331">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="2d7b3-332">Обратите внимание на то, что мы указываем параметр `args` для метода, а самая последняя строка метода вызывает `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-332">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="2d7b3-333">Это требование для всех команд надстройки типа **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-333">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="2d7b3-334">Это сигнализирует ведущему приложению Office о том, что работа функции завершена и пользовательский интерфейс снова может реагировать.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-334">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

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

3. <span data-ttu-id="2d7b3-335">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-335">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="2d7b3-336">В этом коде используется свойство защиты объекта листа в стандартном шаблоне переключателя.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-336">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="2d7b3-337">Объяснение `TODO2` будет приведено в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-337">The `TODO2` will be explained in the next section.</span></span>

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="2d7b3-338">Добавление кода для получения свойств документа в объекты скрипта области задач</span><span class="sxs-lookup"><span data-stu-id="2d7b3-338">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="2d7b3-339">В случае всех описанных ранее функций из этой серии руководств вы ставили в очередь команды для *записи* данных в документ Office.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-339">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="2d7b3-340">Каждая функция заканчивалась вызовом метода `context.sync()`, который отправляет выставленные в очередь команды документу для выполнения.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-340">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="2d7b3-341">Но код, который вы добавили на последнем этапе, вызывает свойство `sheet.protection.protected`, и в этом заключается существенное отличие от ранее написанных функций, так как `sheet` является лишь объектом прокси, существующим в скрипте вашей области задач.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-341">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="2d7b3-342">В нем нет сведений о фактическом состоянии защиты документа, поэтому его свойство `protection.protected` не может иметь реального значения.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-342">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="2d7b3-343">Сначала нужно получить сведения о состоянии защиты от документа и задать значение `sheet.protection.protected`, используя их.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-343">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="2d7b3-344">Только после этого станет возможным вызов `sheet.protection.protected` без исключения.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-344">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="2d7b3-345">Процесс получения делится на три этапа:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-345">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="2d7b3-346">Добавление в очередь команды для загрузки (т. е. получения) свойств, которые должен прочесть ваш код.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-346">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="2d7b3-347">Вызов метода `sync` объекта контекста, чтобы можно было отправить документу находящуюся в очереди команду для выполнения, а также для возврата запрошенных данных.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-347">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="2d7b3-348">Метод `sync` асинхронный, поэтому его выполнение должно быть завершено до того, как код вызовет полученные свойства.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-348">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="2d7b3-349">Эти три действия должны выполняться каждый раз, когда коду нужно *прочесть* данные из документа Office.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-349">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="2d7b3-p144">В функции `toggleProtection` замените `TODO2` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p144">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   
   - <span data-ttu-id="2d7b3-352">У каждого объекта Excel есть метод `load`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-352">Every Excel object has a `load` method.</span></span> <span data-ttu-id="2d7b3-353">Вы указываете свойства объекта, которые нужно прочесть в параметре как строку имен, разделенных запятыми.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-353">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="2d7b3-354">В этом случае нужно прочесть подсвойство свойства `protection`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-354">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="2d7b3-355">На подсвойство нужно ссылаться почти так же, как и в остальных частях кода. Отличие заключается в том, что вместо символа "." нужно указать косую черту ("/").</span><span class="sxs-lookup"><span data-stu-id="2d7b3-355">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="2d7b3-356">Чтобы логика переключения, которая считывает `sheet.protection.protected`, не срабатывала до выполнения `sync` и присвоения `sheet.protection.protected` правильного значения, полученного из документа, она будет перемещена (на следующем этапе) в функцию `then`, которая не выполняется до завершения `sync`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-356">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

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

2. <span data-ttu-id="2d7b3-357">Для двух операторов `return` не может использоваться один путь кода, который не разветвляется, поэтому удалите последнюю строку `return context.sync();` в конце `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-357">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="2d7b3-358">Вы добавите новую последнюю строку `context.sync` позже.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-358">You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="2d7b3-359">Вырежьте структуру `if ... else` в функции `toggleProtection` и вставьте вместо `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-359">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="2d7b3-p147">Замените `TODO4` приведенным ниже кодом. Примечание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-362">Благодаря тому, что метод `sync` передается функции `then`, он не будет запускаться до добавления `sheet.protection.unprotect()` или `sheet.protection.protect()` в очередь.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-362">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="2d7b3-363">Метод `then` вызывает любую функцию, которая ему передана. Не нужно вызывать `sync` дважды, поэтому уберите "()" после `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-363">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="2d7b3-364">Когда все будет готово, функция должна выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-364">When you are done, the entire function should look like the following:</span></span>

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

### <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="2d7b3-365">Настройка HTML-файла для загрузки скрипта</span><span class="sxs-lookup"><span data-stu-id="2d7b3-365">Configure the script-loading HTML file</span></span>

<span data-ttu-id="2d7b3-366">Откройте файл /function-file/function-file.html.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-366">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="2d7b3-367">Это HTML-файл без пользовательского интерфейса, вызываемый, когда пользователь нажимает кнопку **Toggle Worksheet Protection** (Переключение защиты листа).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-367">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="2d7b3-368">Он предназначен для загрузки метода JavaScript, который должен выполняться при нажатии кнопки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-368">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="2d7b3-369">Вы не будете изменять этот файл.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-369">You are not going to change this file.</span></span> <span data-ttu-id="2d7b3-370">Просто обратите внимание на то, что второй тег `<script>` загружает functionfile.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-370">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="2d7b3-371">Файл function-file.html и загружаемый им файл function-file.js выполняются в полностью отдельном процессе IE из области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-371">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="2d7b3-372">Если файл function-file.js был передан в тот же файл bundle.js, что и файл app.js, надстройка загрузит два экземпляра файла bundle.js, и это отменяет цель объединения.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-372">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="2d7b3-373">Кроме того, файл function-file.js не содержит код JavaScript, который не поддерживается в IE.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-373">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="2d7b3-374">По этим двум причинам такая надстройка не передает файл function-file.js вообще.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-374">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

### <a name="test-the-add-in"></a><span data-ttu-id="2d7b3-375">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="2d7b3-375">Test the add-in</span></span>

1. <span data-ttu-id="2d7b3-376">Закройте все приложения Office, в том числе Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-376">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="2d7b3-377">Очистите кэш Office, удалив содержимое папки кэша.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-377">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="2d7b3-378">Это необходимо, чтобы можно было полностью удалить старую версию надстройки из ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-378">This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="2d7b3-379">Для Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-379">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="2d7b3-380">Для Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-380">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

3. <span data-ttu-id="2d7b3-381">Если по той или иной причине ваш сервер не работает, в окне Git Bash или системной командной строке с поддержкой Node.JS перейдите к папке **Start** проекта и выполните команду `npm start`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-381">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="2d7b3-382">Повторную сборку проекта выполнять не нужно, так как единственный файл JavaScript, который вы изменили, не относится к сборке bundle.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-382">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>

4. <span data-ttu-id="2d7b3-383">Используя новую версию измененного файла манифеста, повторите процесс загрузки неопубликованного приложения с помощью одного из указанных далее методов.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-383">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="2d7b3-384">*Нужно перезаписать предыдущий экземпляр файла манифеста.*</span><span class="sxs-lookup"><span data-stu-id="2d7b3-384">*You should overwrite the previous copy of the manifest file.*</span></span>

    - <span data-ttu-id="2d7b3-385">Windows: [загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="2d7b3-385">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="2d7b3-386">[Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="2d7b3-386">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="2d7b3-387">iPad и Mac: [загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="2d7b3-387">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="2d7b3-388">Откройте любой лист в Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-388">Open any worksheet in Excel.</span></span>

6. <span data-ttu-id="2d7b3-p153">На ленте **Главная** нажмите кнопку **Toggle Worksheet Protection** (Переключение защиты листа). Обратите внимание на то, что большинство элементов управления на ленте отключены (серые), как показано на приведенном ниже снимке экрана.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p153">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 

7. <span data-ttu-id="2d7b3-391">Выберите ячейку, как если бы вы хотели изменить ее содержимое.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-391">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="2d7b3-392">Появится сообщение об ошибке и защите листа.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-392">You get an error telling you that the worksheet is protected.</span></span>

8. <span data-ttu-id="2d7b3-393">Нажмите кнопку **Toggle Worksheet Protection** (Переключение защиты листа) еще раз, и элементы управления включатся, после чего вы сможете изменить значения ячеек.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-393">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Руководство по Excel: лента с включенной защитой](../images/excel-tutorial-ribbon-with-protection-on.png)

## <a name="open-a-dialog"></a><span data-ttu-id="2d7b3-395">Открытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="2d7b3-395">Open a dialog</span></span>

<span data-ttu-id="2d7b3-396">На данном заключительном этапе, указанном в руководстве, вы откроете диалоговое окно в своей надстройке, передадите сообщение из процесса диалогового окна в процесс области задач и закроете диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-396">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="2d7b3-397">Диалоговые окна надстройки Office *не модальные*: пользователь может продолжать работать и с документом в ведущем приложении Office, и с главной страницей в области задач.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-397">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="2d7b3-398">Создание страницы диалогового окна</span><span class="sxs-lookup"><span data-stu-id="2d7b3-398">Create the dialog page</span></span>

1. <span data-ttu-id="2d7b3-399">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-399">Open the project in your code editor.</span></span>

2. <span data-ttu-id="2d7b3-400">Создайте в корневой папке проекта (где находится index.html) файл popup.html.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-400">Create a file in the root of the project (where index.html is) called popup.html.</span></span>

3. <span data-ttu-id="2d7b3-p156">Добавьте в файл popup.html приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p156">Add the following markup to popup.html. Note:</span></span>

   - <span data-ttu-id="2d7b3-403">На странице находится `<input>`, где пользователь будет вводить свое имя, и кнопка, при нажатии которой имя будет отправлено на страницу области задач, где оно отобразится.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-403">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="2d7b3-404">Код загружает скрипт под названием popup.js, который будет создан на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-404">The markup loads a script called popup.js that you will create in a later step.</span></span>

   - <span data-ttu-id="2d7b3-405">Он загружает также библиотеку Office.JS и jQuery, так как они будут использоваться в popup.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-405">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

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

4. <span data-ttu-id="2d7b3-406">Создайте файл в корневой папке проекта с именем popup.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-406">Create a file in the root of the project called popup.js.</span></span>

5. <span data-ttu-id="2d7b3-407">Добавьте указанный ниже код в файл popup.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-407">Add the following code to popup.js.</span></span> <span data-ttu-id="2d7b3-408">Обратите внимание на указанные ниже особенности этого кода.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-408">Note the following about this code:</span></span>

   - <span data-ttu-id="2d7b3-409">*Каждая страница, вызывающая API в библиотеке Office.JS, должна сначала убедиться, что библиотека полностью инициализирована.*</span><span class="sxs-lookup"><span data-stu-id="2d7b3-409">*Every page that calls APIs in the Office.JS library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="2d7b3-410">Лучший способ сделать это — вызвать метод `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-410">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="2d7b3-411">Если у вашей надстройки есть собственные задачи инициализации, код должен перейти к методу `then()`, связанному с вызовом `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-411">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="2d7b3-412">Файл app.js в корневом каталоге проекта можно рассматривать как пример.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-412">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="2d7b3-413">Вызов метода `Office.onReady()` должен выполняться до каких-либо вызовов Office.JS, поэтому назначение указано в файле скрипта, загружаемом страницей, как в этом случае.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-413">The code that makes the assignment must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   - <span data-ttu-id="2d7b3-414">Функция jQuery `ready` вызывается в методе `then()`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-414">The jQuery `ready` function is called inside the `then()` method.</span></span> <span data-ttu-id="2d7b3-415">В большинстве случаев код загрузки (в том числе начальной) или инициализации из других библиотек JavaScript должен находиться в методе `then()`, связанном с вызовом `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-415">In most cases, the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `then()` method that is chained to the call of `Office.onReady()`.</span></span>

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

6. <span data-ttu-id="2d7b3-416">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-416">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="2d7b3-417">Вы создадите функцию `sendStringToParentPage` на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-417">You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="2d7b3-418">Замените `TODO2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-418">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="2d7b3-419">Метод `messageParent` передает свой параметр родительской странице (в данном случае это страница на панели задач).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-419">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="2d7b3-420">Параметр может быть логическим или строковым. Во втором случае подразумевается все, что можно сериализовать, представив в виде строки (например, XML или JSON).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-420">The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="2d7b3-421">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-421">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="2d7b3-422">Файл popup.html и загружаемый им файл popup.js выполняются в полностью отдельном процессе Internet Explorer из области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-422">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane.</span></span> <span data-ttu-id="2d7b3-423">Если файл popup.js был передан в тот же файл bundle.js, что и файл app.js, надстройка загрузит два экземпляра файла bundle.js, и это отменяет цель объединения.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-423">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="2d7b3-424">Кроме того, файл popup.js не содержит код JavaScript, который не поддерживается в IE.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-424">In addition, the popup.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="2d7b3-425">По этим двум причинам эта надстройка не передает файл popup.js вообще.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-425">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span>

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="2d7b3-426">Открытие диалогового окна из области задач</span><span class="sxs-lookup"><span data-stu-id="2d7b3-426">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="2d7b3-427">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-427">Open the file index.html.</span></span>

2. <span data-ttu-id="2d7b3-428">Под `div` с кнопкой `freeze-header` добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-428">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. <span data-ttu-id="2d7b3-429">В диалоговом окне пользователю будет предложено ввести имя и передать имя пользователя в область задач.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-429">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="2d7b3-430">Область задач отобразит его в подписи.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-430">The task pane will display it in a label.</span></span> <span data-ttu-id="2d7b3-431">Непосредственно под только что добавленным тегом `div` добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-431">Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. <span data-ttu-id="2d7b3-432">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-432">Open the app.js file.</span></span>

5. <span data-ttu-id="2d7b3-433">Под строкой, назначающей обработчик щелчков для кнопки `freeze-header`, добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-433">Below the line that assigns a click handler to the `freeze-header` button, add the following code.</span></span> <span data-ttu-id="2d7b3-434">Вы создадите метод `openDialog` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-434">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="2d7b3-p165">Под функцией `freezeHeader` добавьте указанное ниже объявление. Эта переменная удерживает объект в контексте выполнения родительской страницы, который служит посредником для контекста выполнения страницы диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p165">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="2d7b3-437">Добавьте приведенную ниже функцию под объявлением `dialog`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-437">Below the declaration of `dialog`, add the following function.</span></span> <span data-ttu-id="2d7b3-438">Важно отметить, что в этом коде *отсутствует* вызов `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-438">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="2d7b3-439">Это связано с тем, что API, открывающий диалоговое окно, совместно используется всеми ведущими приложениями Office, поэтому относится к общему API JavaScript для Office, а не API для Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-439">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="2d7b3-p167">Замените `TODO1` приведенным ниже кодом. Примечание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p167">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-442">Метод `displayDialogAsync` открывает диалоговое окно в центре экрана.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-442">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="2d7b3-443">Первый параметр — это URL-адрес открываемой страницы.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-443">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="2d7b3-p168">Второй параметр передает параметры. `height` и `width` — процентные значения размера окна для приложения Office.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="2d7b3-446">Обработка сообщения из диалогового окна и закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="2d7b3-446">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="2d7b3-p169">Продолжайте работать в файле app.js. Замените `TODO2` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-p169">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="2d7b3-449">Обратный вызов выполняется сразу же после успешного открытия диалогового окна и до того, как пользователь предпримет какие-либо действия в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-449">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="2d7b3-450">`result.value` — это объект, который выступает в качестве посредника между контекстами выполнения родительских страниц и страниц диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-450">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="2d7b3-451">Функция `processMessage` будет создана на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-451">The `processMessage` function will be created in a later step.</span></span> <span data-ttu-id="2d7b3-452">Этот обработчик будет обрабатывать любые значения, которые отправляются со страницы диалогового окна с вызовами функции `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-452">This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="2d7b3-453">Добавьте указанную ниже функцию под функцией `openDialog`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-453">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="2d7b3-454">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="2d7b3-454">Test the add-in</span></span>

1. <span data-ttu-id="2d7b3-455">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши **Ctrl+C**, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-455">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="2d7b3-456">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-456">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="2d7b3-457">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-457">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="2d7b3-458">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-458">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="2d7b3-459">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-459">After the build, you restart the server.</span></span> <span data-ttu-id="2d7b3-460">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-460">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="2d7b3-461">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="2d7b3-461">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="2d7b3-462">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-462">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="2d7b3-463">Повторно загрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Show Taskpane** (Показать область задач) для повторного открытия надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-463">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="2d7b3-464">Нажмите кнопку **Open Dialog** (Открыть диалоговое окно) в области задач.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-464">Choose the **Open Dialog** button in the task pane.</span></span>

6. <span data-ttu-id="2d7b3-465">Когда диалоговое окно открыто, перетащите его и измените его размер.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-465">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="2d7b3-466">Обратите внимание, что вы можете взаимодействовать с листом и нажимать другие кнопки в области задач, но вы не можете запустить второе диалоговое окно на одной и той же странице панели задач.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-466">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

7. <span data-ttu-id="2d7b3-467">В диалоговом окне введите имя и нажмите кнопку **OK**.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-467">In the dialog, enter a name and choose **OK**.</span></span> <span data-ttu-id="2d7b3-468">В области задач отобразится имя, и диалоговое окно закроется.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-468">The name appears on the task pane and the dialog closes.</span></span>

8. <span data-ttu-id="2d7b3-469">При желании можно закомментировать строку `dialog.close();` в функции `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-469">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="2d7b3-470">Повторите шаги этого раздела.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-470">Then repeat the steps of this section.</span></span> <span data-ttu-id="2d7b3-471">Диалоговое окно остается открытым, и вы можете изменить имя.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-471">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="2d7b3-472">Можно закрыть его вручную, нажав кнопку **X** в правом верхнему углу.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-472">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Руководство по Excel - Диалоговое окно](../images/excel-tutorial-dialog-open.png)

## <a name="next-steps"></a><span data-ttu-id="2d7b3-474">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="2d7b3-474">Next steps</span></span>

<span data-ttu-id="2d7b3-475">В этом руководстве показано создание надстройки Excel для области задач, которая взаимодействует с таблицами, диаграммами, листами, диалоговыми окнами в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7b3-475">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="2d7b3-476">Чтобы узнать больше о создании надстроек Excel, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="2d7b3-476">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="2d7b3-477">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="2d7b3-477">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)

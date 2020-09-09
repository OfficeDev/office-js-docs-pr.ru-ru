---
title: Работа с таблицами с использованием API JavaScript для Excel
description: Примеры кода, демонстрирующие выполнение типовых задач с таблицами с помощью API JavaScript для Excel.
ms.date: 09/09/2019
localization_priority: Normal
ms.openlocfilehash: b358ff33aa3681043f86d650ae2dd9b01a95f962
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408476"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a><span data-ttu-id="fb568-103">Работа с таблицами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="fb568-103">Work with tables using the Excel JavaScript API</span></span>

<span data-ttu-id="fb568-104">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для таблиц с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="fb568-104">This article provides code samples that show how to perform common tasks with tables using the Excel JavaScript API.</span></span> <span data-ttu-id="fb568-105">Полный список свойств и методов, `Table` `TableCollection` поддерживаемых объектами and, представлен в статье [объект Table (API JavaScript для Excel)](/javascript/api/excel/excel.table) и [объект TableCollection (API JavaScript для Excel)](/javascript/api/excel/excel.tablecollection).</span><span class="sxs-lookup"><span data-stu-id="fb568-105">For the complete list of properties and methods that the `Table` and `TableCollection` objects support, see [Table Object (JavaScript API for Excel)](/javascript/api/excel/excel.table) and [TableCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.tablecollection).</span></span>

## <a name="create-a-table"></a><span data-ttu-id="fb568-106">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="fb568-106">Create a table</span></span>

<span data-ttu-id="fb568-107">В примере кода ниже показано, как создать таблицу на листе **Sample** (Пример).</span><span class="sxs-lookup"><span data-stu-id="fb568-107">The following code sample creates a table in the worksheet named **Sample**.</span></span> <span data-ttu-id="fb568-108">В таблице имеются заголовки, а также четыре столбца и семь строк с данными.</span><span class="sxs-lookup"><span data-stu-id="fb568-108">The table has headers and contains four columns and seven rows of data.</span></span> <span data-ttu-id="fb568-109">Если приложение Excel, в котором выполняется код, поддерживает [набор требований](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1,2**, ширина столбцов и высота строк задаются в соответствии с текущими данными в таблице.</span><span class="sxs-lookup"><span data-stu-id="fb568-109">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="fb568-110">Чтобы указать имя для таблицы, необходимо сначала создать таблицу, а затем задать ее `name` свойство, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="fb568-110">To specify a name for a table, you must first create the table and then set its `name` property, as shown in the following example.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-111">**Новая таблица**</span><span class="sxs-lookup"><span data-stu-id="fb568-111">**New table**</span></span>

![Новая таблица в Excel](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a><span data-ttu-id="fb568-113">Добавление строк в таблицу</span><span class="sxs-lookup"><span data-stu-id="fb568-113">Add rows to a table</span></span>

<span data-ttu-id="fb568-114">В примере ниже показано, как добавить семь новых строк в таблицу **ExpensesTable** (Таблица расходов) на листе **Sample** (Пример).</span><span class="sxs-lookup"><span data-stu-id="fb568-114">The following code sample adds seven new rows to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="fb568-115">Новые строки будут добавлены в конец таблицы.</span><span class="sxs-lookup"><span data-stu-id="fb568-115">The new rows are added to the end of the table.</span></span> <span data-ttu-id="fb568-116">Если приложение Excel, в котором выполняется код, поддерживает [набор требований](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1,2**, ширина столбцов и высота строк задаются в соответствии с текущими данными в таблице.</span><span class="sxs-lookup"><span data-stu-id="fb568-116">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="fb568-117">`index`Свойство объекта [TableRow](/javascript/api/excel/excel.tablerow) указывает номер индекса строки в коллекции Rows таблицы.</span><span class="sxs-lookup"><span data-stu-id="fb568-117">The `index` property of a [TableRow](/javascript/api/excel/excel.tablerow) object indicates the index number of the row within the rows collection of the table.</span></span> <span data-ttu-id="fb568-118">`TableRow`Объект не содержит `id` свойство, которое можно использовать в качестве уникального ключа для идентификации строки.</span><span class="sxs-lookup"><span data-stu-id="fb568-118">A `TableRow` object does not contain an `id` property that can be used as a unique key to identify the row.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/16/2017", "THE PHONE COMPANY", "Communications", "$120"],
        ["1/20/2017", "NORTHWIND ELECTRIC CARS", "Transportation", "$142"],
        ["1/20/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$27"],
        ["1/21/2017", "COHO VINEYARD", "Restaurant", "$33"],
        ["1/25/2017", "BELLOWS COLLEGE", "Education", "$350"],
        ["1/28/2017", "TREY RESEARCH", "Other", "$135"],
        ["1/31/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-119">**Таблица с новыми строками**</span><span class="sxs-lookup"><span data-stu-id="fb568-119">**Table with new rows**</span></span>

![Таблица с новыми строками в Excel](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a><span data-ttu-id="fb568-121">Добавление столбца в таблицу</span><span class="sxs-lookup"><span data-stu-id="fb568-121">Add a column to a table</span></span>

<span data-ttu-id="fb568-p105">В примерах ниже показано, как добавить столбец в таблицу. В первом примере показано, как заполнить новый столбец статическими значениями, во втором — как заполнить новый столбец формулами.</span><span class="sxs-lookup"><span data-stu-id="fb568-p105">These examples show how to add a column to a table. The first example populates the new column with static values; the second example populates the new column with formulas.</span></span>

> [!NOTE]
> <span data-ttu-id="fb568-p106">Свойство **index** объекта [TableColumn](/javascript/api/excel/excel.tablecolumn) указывает номер индекса столбца в коллекции столбцов таблицы. Свойство **id** объекта **TableColumn** содержит уникальный ключ, идентифицирующий столбец.</span><span class="sxs-lookup"><span data-stu-id="fb568-p106">The **index** property of a [TableColumn](/javascript/api/excel/excel.tablecolumn) object indicates the index number of the column within the columns collection of the table. The **id** property of a **TableColumn** object contains a unique key that identifies the column.</span></span>

### <a name="add-a-column-that-contains-static-values"></a><span data-ttu-id="fb568-126">Добавление столбца, содержащего статические значения</span><span class="sxs-lookup"><span data-stu-id="fb568-126">Add a column that contains static values</span></span>

<span data-ttu-id="fb568-127">В примере кода ниже показано, как добавить новый столбец в таблицу **ExpensesTable** (Таблица расходов) на листе **Sample** (Пример).</span><span class="sxs-lookup"><span data-stu-id="fb568-127">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="fb568-128">Новый столбец будет добавлен после всех существующих столбцов в таблице. Он будет содержать заголовок Day of the Week (День недели), а также данные для заполнения ячеек в столбце.</span><span class="sxs-lookup"><span data-stu-id="fb568-128">The new column is added after all existing columns in the table and contains a header ("Day of the Week") as well as data to populate the cells in the column.</span></span> <span data-ttu-id="fb568-129">Если приложение Excel, в котором выполняется код, поддерживает [набор требований](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1,2**, ширина столбцов и высота строк задаются в соответствии с текущими данными в таблице.</span><span class="sxs-lookup"><span data-stu-id="fb568-129">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Day of the Week"],
        ["Saturday"],
        ["Friday"],
        ["Monday"],
        ["Thursday"],
        ["Sunday"],
        ["Saturday"],
        ["Monday"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-130">**Таблица с новым столбцом**</span><span class="sxs-lookup"><span data-stu-id="fb568-130">**Table with new column**</span></span>

![Таблица с новым столбцом в Excel](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a><span data-ttu-id="fb568-132">Добавление столбца, содержащего формулы</span><span class="sxs-lookup"><span data-stu-id="fb568-132">Add a column that contains formulas</span></span>

<span data-ttu-id="fb568-133">В примере кода ниже показано, как добавить новый столбец в таблицу **ExpensesTable** (Таблица расходов) на листе **Sample** (Пример).</span><span class="sxs-lookup"><span data-stu-id="fb568-133">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="fb568-134">Новый столбец будет добавлен в конец таблицы, будет содержать заголовок Type of the Day (Тип дня), и в нем будет использована формула для заполнения каждой ячейки столбца.</span><span class="sxs-lookup"><span data-stu-id="fb568-134">The new column is added to the end of the table, contains a header ("Type of the Day"), and uses a formula to populate each data cell in the column.</span></span> <span data-ttu-id="fb568-135">Если приложение Excel, в котором выполняется код, поддерживает [набор требований](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1,2**, ширина столбцов и высота строк задаются в соответствии с текущими данными в таблице.</span><span class="sxs-lookup"><span data-stu-id="fb568-135">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Type of the Day"],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")']
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-136">**Таблица с новым столбцом, содержащим вычисленные значения**</span><span class="sxs-lookup"><span data-stu-id="fb568-136">**Table with new calculated column**</span></span>

![Таблица с новым столбцом, содержащим вычисленные значения, в Excel](../images/excel-tables-add-calculated-column.png)

## <a name="update-column-name"></a><span data-ttu-id="fb568-138">Изменение имени столбца</span><span class="sxs-lookup"><span data-stu-id="fb568-138">Update column name</span></span>

<span data-ttu-id="fb568-139">В примере кода ниже показано, как изменить имя первого столбца в таблице на **Purchase date**.</span><span class="sxs-lookup"><span data-stu-id="fb568-139">The following code sample updates the name of the first column in the table to **Purchase date**.</span></span> <span data-ttu-id="fb568-140">Если приложение Excel, в котором выполняется код, поддерживает [набор требований](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1,2**, ширина столбцов и высота строк задаются в соответствии с текущими данными в таблице.</span><span class="sxs-lookup"><span data-stu-id="fb568-140">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    return context.sync()
        .then(function () {
            expensesTable.columns.items[0].name = "Purchase date";

            if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
            }

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-141">**Таблица со столбцом с новым именем**</span><span class="sxs-lookup"><span data-stu-id="fb568-141">**Table with new column name**</span></span>

![Таблица со столбцом с новым именем в Excel](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a><span data-ttu-id="fb568-143">Получение данных из таблицы</span><span class="sxs-lookup"><span data-stu-id="fb568-143">Get data from a table</span></span>

<span data-ttu-id="fb568-144">В примере кода ниже показано, как считать данные из таблицы **ExpensesTable** (Таблица расходов), размещенной на листе **Sample** (Пример), а затем отобразить эти данные под таблицей на том же листе.</span><span class="sxs-lookup"><span data-stu-id="fb568-144">The following code sample reads data from a table named **ExpensesTable** in the worksheet named **Sample** and then outputs that data below the table in the same worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Get data from the header row
    var headerRange = expensesTable.getHeaderRowRange().load("values");

    // Get data from the table
    var bodyRange = expensesTable.getDataBodyRange().load("values");

    // Get data from a single column
    var columnRange = expensesTable.columns.getItem("Merchant").getDataBodyRange().load("values");

    // Get data from a single row
    var rowRange = expensesTable.rows.getItemAt(1).load("values");

    // Sync to populate proxy objects with data from Excel
    return context.sync()
        .then(function () {
            var headerValues = headerRange.values;
            var bodyValues = bodyRange.values;
            var merchantColumnValues = columnRange.values;
            var secondRowValues = rowRange.values;

            // Write data from table back to the sheet
            sheet.getRange("A11:A11").values = [["Results"]];
            sheet.getRange("A13:D13").values = headerValues;
            sheet.getRange("A14:D20").values = bodyValues;
            sheet.getRange("B23:B29").values = merchantColumnValues;
            sheet.getRange("A32:D32").values = secondRowValues;

            // Sync to update the sheet in Excel
            return context.sync();
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-145">**Таблица и выведенные данные**</span><span class="sxs-lookup"><span data-stu-id="fb568-145">**Table and data output**</span></span>

![Данные из таблицы в Excel](../images/excel-tables-get-data.png)

## <a name="detect-data-changes"></a><span data-ttu-id="fb568-147">Обнаружение изменений данных</span><span class="sxs-lookup"><span data-stu-id="fb568-147">Detect data changes</span></span>

<span data-ttu-id="fb568-148">Возможно, надстройке потребуется реагировать на изменения пользователями данных в таблице.</span><span class="sxs-lookup"><span data-stu-id="fb568-148">Your add-in may need to react to users changing the data in a table.</span></span> <span data-ttu-id="fb568-149">Чтобы обнаружить эти изменения, можно [зарегистрировать обработчик событий](excel-add-ins-events.md#register-an-event-handler) для события `onChanged` таблицы.</span><span class="sxs-lookup"><span data-stu-id="fb568-149">To detect these changes, you can [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a table.</span></span> <span data-ttu-id="fb568-150">Обработчики события `onChanged` получают объект [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) при возникновении события.</span><span class="sxs-lookup"><span data-stu-id="fb568-150">Event handlers for the `onChanged` event receive a [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) object when the event fires.</span></span>

<span data-ttu-id="fb568-151">Объект `TableChangedEventArgs` предоставляет сведения об изменениях и источнике.</span><span class="sxs-lookup"><span data-stu-id="fb568-151">The `TableChangedEventArgs` object provides information about the changes and the source.</span></span> <span data-ttu-id="fb568-152">Так как событие `onChanged` возникает при изменении формата или значения данных, может быть полезно, чтобы надстройка проверяла, действительно ли значения изменились.</span><span class="sxs-lookup"><span data-stu-id="fb568-152">Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values have actually changed.</span></span> <span data-ttu-id="fb568-153">Свойство `details` объединяет эти сведения в виде интерфейса [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span><span class="sxs-lookup"><span data-stu-id="fb568-153">The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span></span> <span data-ttu-id="fb568-154">В следующем примере кода показано, как отобразить значения и типы измененной ячейки до и после изменения.</span><span class="sxs-lookup"><span data-stu-id="fb568-154">The following code sample shows how to display the before and after values and types of a cell that has been changed.</span></span>

```js
// This function would be used as an event handler for the Table.onChanged event.
function onTableChanged(eventArgs) {
    Excel.run(function (context) {
        var details = eventArgs.details;
        var address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="sort-data-in-a-table"></a><span data-ttu-id="fb568-155">Сортировка данных в таблице</span><span class="sxs-lookup"><span data-stu-id="fb568-155">Sort data in a table</span></span>

<span data-ttu-id="fb568-156">В примере кода ниже показано, как отсортировать данные по убыванию в четвертом столбце таблицы.</span><span class="sxs-lookup"><span data-stu-id="fb568-156">The following code sample sorts table data in descending order according to the values in the fourth column of the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to sort data by the fourth column of the table (descending)
    var sortRange = expensesTable.getDataBodyRange();
    sortRange.sort.apply([
        {
            key: 3,
            ascending: false,
        },
    ]);

    // Sync to run the queued command in Excel
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-157">**Данные таблицы, отсортированные по столбцу Amount (Сумма) в порядке убывания**</span><span class="sxs-lookup"><span data-stu-id="fb568-157">**Table data sorted by Amount (descending)**</span></span>

![Сортировка табличных данных в Excel](../images/excel-tables-sort.png)

<span data-ttu-id="fb568-159">При сортировке данных на листе создается уведомление о событии.</span><span class="sxs-lookup"><span data-stu-id="fb568-159">When data is sorted in a worksheet, an event notification fires.</span></span> <span data-ttu-id="fb568-160">Дополнительные сведения о событиях, связанных с сортировкой, и о регистрации обработчиков событий надстройкой в ответ на такие события см. в статье [Обработка событий сортировки](excel-add-ins-worksheets.md#handle-sorting-events).</span><span class="sxs-lookup"><span data-stu-id="fb568-160">To learn more about sort-related events and how your add-in can register event handlers to respond to such events, see [Handle sorting events](excel-add-ins-worksheets.md#handle-sorting-events).</span></span>

## <a name="apply-filters-to-a-table"></a><span data-ttu-id="fb568-161">Применение фильтров к таблице</span><span class="sxs-lookup"><span data-stu-id="fb568-161">Apply filters to a table</span></span>

<span data-ttu-id="fb568-p113">В примере кода ниже показано, как применить фильтры для столбцов **Amount** (Сумма) и **Category** (Категория) в таблице. В результате применения фильтров будут отображены только те строки, у которых в столбце **Category** (Категория) содержится одно из указанных значений, а значения в столбце **Amount** (Сумма) меньше среднего значения для всех строк.</span><span class="sxs-lookup"><span data-stu-id="fb568-p113">The following code sample applies filters to the **Amount** column and the **Category** column within a table. As a result of the filters, only rows where **Category** is one of the specified values and **Amount** is below the average value for all rows is shown.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to apply a filter on the Category column
    filter = expensesTable.columns.getItem("Category").filter;
    filter.apply({
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });

    // Queue a command to apply a filter on the Amount column
    var filter = expensesTable.columns.getItem("Amount").filter;
    filter.apply({
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    // Sync to run the queued commands in Excel
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-164">**Таблица данных, в которой применены фильтры для столбцов Category (Категория) и Amount (Сумма)**</span><span class="sxs-lookup"><span data-stu-id="fb568-164">**Table data with filters applied for Category and Amount**</span></span>

![Отфильтрованные данные таблицы в Excel](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a><span data-ttu-id="fb568-166">Удаление фильтров в таблице</span><span class="sxs-lookup"><span data-stu-id="fb568-166">Clear table filters</span></span>

<span data-ttu-id="fb568-167">В примере кода ниже показано, как удалить все фильтры, примененные к таблице.</span><span class="sxs-lookup"><span data-stu-id="fb568-167">The following code sample clears any filters currently applied on the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-168">**Данные таблицы без фильтров**</span><span class="sxs-lookup"><span data-stu-id="fb568-168">**Table data with no filters applied**</span></span>

![Неотфильтрованные данные таблицы в Excel](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a><span data-ttu-id="fb568-170">Получение отображаемого диапазона из отфильтрованной таблицы</span><span class="sxs-lookup"><span data-stu-id="fb568-170">Get the visible range from a filtered table</span></span>

<span data-ttu-id="fb568-171">В примере кода ниже показано, как получить диапазон, содержащий данные только из тех ячеек, которые в данный момент отображаются в указанной таблице, и записать значения из этого диапазона в консоль.</span><span class="sxs-lookup"><span data-stu-id="fb568-171">The following code sample gets a range that contains data only for cells that are currently visible within the specified table, and then writes the values of that range to the console.</span></span> <span data-ttu-id="fb568-172">Вы можете использовать `getVisibleView()` метод, как показано ниже, чтобы получить видимое содержимое таблицы при применении фильтров столбцов.</span><span class="sxs-lookup"><span data-stu-id="fb568-172">You can use the `getVisibleView()` method as shown below to get the visible contents of a table whenever column filters have been applied.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    var visibleRange = expensesTable.getDataBodyRange().getVisibleView();
    visibleRange.load("values");

    return context.sync()
        .then(function() {
            console.log(visibleRange.values);
        });
}).catch(errorHandlerFunction);
```

## <a name="autofilter"></a><span data-ttu-id="fb568-173">Автофильтр</span><span class="sxs-lookup"><span data-stu-id="fb568-173">AutoFilter</span></span>

<span data-ttu-id="fb568-174">Надстройка может использовать объект [AutoFilter](/javascript/api/excel/excel.autofilter) таблицы для фильтрации данных.</span><span class="sxs-lookup"><span data-stu-id="fb568-174">An add-in can use the table's [AutoFilter](/javascript/api/excel/excel.autofilter) object to filter data.</span></span> <span data-ttu-id="fb568-175">Объект `AutoFilter` является целой структурой фильтра таблицы или диапазона.</span><span class="sxs-lookup"><span data-stu-id="fb568-175">An `AutoFilter` object is the entire filter structure of a table or range.</span></span> <span data-ttu-id="fb568-176">Все операции фильтрации, описанные выше в этой статье, совместимы с автофильтром.</span><span class="sxs-lookup"><span data-stu-id="fb568-176">All of the filter operations discussed earlier in this article are compatible with the auto-filter.</span></span> <span data-ttu-id="fb568-177">Единая точка доступа упрощает доступ к нескольким фильтрам и управление ими.</span><span class="sxs-lookup"><span data-stu-id="fb568-177">The single access point does make it easier to access and manage multiple filters.</span></span>

<span data-ttu-id="fb568-178">В следующем примере кода показана такая же [фильтрация данных, как в примере кода выше](#apply-filters-to-a-table), но выполненная полностью с помощью автофильтра.</span><span class="sxs-lookup"><span data-stu-id="fb568-178">The following code sample shows the same [data filtering as the earlier code sample](#apply-filters-to-a-table), but done entirely through the auto-filter.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.autoFilter.apply(expensesTable.getRange(), 2, {
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });
    expensesTable.autoFilter.apply(expensesTable.getRange(), 3, {
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-179">Объект `AutoFilter` можно также применять к диапазону на уровне листа.</span><span class="sxs-lookup"><span data-stu-id="fb568-179">An `AutoFilter` can also be applied to a range at the worksheet level.</span></span> <span data-ttu-id="fb568-180">Дополнительные сведения см. в статье [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md#filter-data).</span><span class="sxs-lookup"><span data-stu-id="fb568-180">See [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#filter-data) for more information.</span></span>

## <a name="format-a-table"></a><span data-ttu-id="fb568-181">Форматирование таблицы</span><span class="sxs-lookup"><span data-stu-id="fb568-181">Format a table</span></span>

<span data-ttu-id="fb568-p117">В примере кода ниже показано, как применить форматирование к таблице. В примере показано, как указать различные цвета заливки для строки заголовков, основной части, второй строки и первого столбца таблицы. Сведения о свойствах, которые вы можете использовать для задания формата, см. в статье [Объект RangeFormat (API JavaScript для Excel)](/javascript/api/excel/excel.rangeformat).</span><span class="sxs-lookup"><span data-stu-id="fb568-p117">The following code sample applies formatting to a table. It specifies different fill colors for the header row of the table, the body of the table, the second row of the table, and the first column of the table. For information about the properties you can use to specify format, see [RangeFormat Object (JavaScript API for Excel)](/javascript/api/excel/excel.rangeformat).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.getHeaderRowRange().format.fill.color = "#C70039";
    expensesTable.getDataBodyRange().format.fill.color = "#DAF7A6";
    expensesTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";
    expensesTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-185">**Таблица после применения форматирования**</span><span class="sxs-lookup"><span data-stu-id="fb568-185">**Table after formatting is applied**</span></span>

![Таблица после применения форматирования в Excel](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a><span data-ttu-id="fb568-187">Преобразование диапазона в таблицу</span><span class="sxs-lookup"><span data-stu-id="fb568-187">Convert a range to a table</span></span>

<span data-ttu-id="fb568-188">В примере кода ниже показано, как создать диапазон данных и преобразовывать его в таблицу.</span><span class="sxs-lookup"><span data-stu-id="fb568-188">The following code sample creates a range of data and then converts that range to a table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Define values for the range
    var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
    ["Frames", 5000, 7000, 6544, 4377],
    ["Saddles", 400, 323, 276, 651],
    ["Brake levers", 12000, 8766, 8456, 9812],
    ["Chains", 1550, 1088, 692, 853],
    ["Mirrors", 225, 600, 923, 544],
    ["Spokes", 6005, 7634, 4589, 8765]];

    // Create the range
    var range = sheet.getRange("A1:E7");
    range.values = values;

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    // Convert the range to a table
    var expensesTable = sheet.tables.add('A1:E7', true);
    expensesTable.name = "ExpensesTable";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-189">**Данные в диапазоне (перед его преобразованием в таблицу)**</span><span class="sxs-lookup"><span data-stu-id="fb568-189">**Data in the range (before the range is converted to a table)**</span></span>

![Данные в диапазоне в Excel](../images/excel-ranges.png)

<span data-ttu-id="fb568-191">**Данные в таблице (после преобразования диапазона в таблицу)**</span><span class="sxs-lookup"><span data-stu-id="fb568-191">**Data in the table (after the range is converted to a table)**</span></span>

![Данные в таблице в Excel](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a><span data-ttu-id="fb568-193">Импорт данных JSON в таблицу</span><span class="sxs-lookup"><span data-stu-id="fb568-193">Import JSON data into a table</span></span>

<span data-ttu-id="fb568-194">В примере кода ниже показано, как создать таблицу на листе **Sample** (Пример), а затем заполнить ее с помощью объекта JSON, который определяет две строки данных.</span><span class="sxs-lookup"><span data-stu-id="fb568-194">The following code sample creates a table in the worksheet named **Sample** and then populates the table by using a JSON object that defines two rows of data.</span></span> <span data-ttu-id="fb568-195">Если приложение Excel, в котором выполняется код, поддерживает [набор требований](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1,2**, ширина столбцов и высота строк задаются в соответствии с текущими данными в таблице.</span><span class="sxs-lookup"><span data-stu-id="fb568-195">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    var transactions = [
      {
        "DATE": "1/1/2017",
        "MERCHANT": "The Phone Company",
        "CATEGORY": "Communications",
        "AMOUNT": "$120"
      },
      {
        "DATE": "1/1/2017",
        "MERCHANT": "Southridge Video",
        "CATEGORY": "Entertainment",
        "AMOUNT": "$40"
      }
    ];

    var newData = transactions.map(item =>
        [item.DATE, item.MERCHANT, item.CATEGORY, item.AMOUNT]);

    expensesTable.rows.add(null, newData);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fb568-196">**Новая таблица**</span><span class="sxs-lookup"><span data-stu-id="fb568-196">**New table**</span></span>

![Новая таблица из импортированных данных JSON в Excel](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a><span data-ttu-id="fb568-198">См. также</span><span class="sxs-lookup"><span data-stu-id="fb568-198">See also</span></span>

- [<span data-ttu-id="fb568-199">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="fb568-199">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)

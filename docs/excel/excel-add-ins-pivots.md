---
title: Работа со сводными таблицами с помощью API JavaScript для Excel
description: Использование Excel JavaScript API для создания сводных таблиц и взаимодействия с их компонентами.
ms.date: 08/17/2018
ms.openlocfilehash: aa6da2e82ab9b0c255208a86012d51db77982934
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2018
ms.locfileid: "22493980"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="93df1-103">Работа со сводными таблицами с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="93df1-103">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="93df1-104">Сводные таблицы упрощают большие наборы данных.</span><span class="sxs-lookup"><span data-stu-id="93df1-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="93df1-105">Они позволяют производить быструю манипуляцию сгруппированных данных.</span><span class="sxs-lookup"><span data-stu-id="93df1-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="93df1-106">API JavaScript для Excel позволяет надстройке создавать сводные таблицы и взаимодействовать с их компонентами.</span><span class="sxs-lookup"><span data-stu-id="93df1-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="93df1-107">Если вы не знакомы с функциональностью сводных таблиц, рекомендуем рассмотреть их в качестве конечного пользователя.</span><span class="sxs-lookup"><span data-stu-id="93df1-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="93df1-108">См. [Создание сводной таблицы для анализа данных листа](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) как учебник для начинающих на этих инструментах.</span><span class="sxs-lookup"><span data-stu-id="93df1-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="93df1-109">В этой статье приведены примеры кода для распространенных сценариев.</span><span class="sxs-lookup"><span data-stu-id="93df1-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="93df1-110"> [Excel OpenSpec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel) предоставляет полный справочную документацию по этой функции предварительного просмотра.</span><span class="sxs-lookup"><span data-stu-id="93df1-110">The [Excel OpenSpec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel) provides full reference documentation for this preview feature.</span></span> 

<span data-ttu-id="93df1-111">Для дальнейшего понимания API сводной таблицы см. [**PivotTable**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottable.md) и [**PivotTableCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottablecollection.md).</span><span class="sxs-lookup"><span data-stu-id="93df1-111">To further your understanding of the PivotTable API, see [**PivotTable**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottable.md) and [**PivotTableCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottablecollection.md).</span></span>

> [!NOTE]
> <span data-ttu-id="93df1-112">Эти примеры использования API доступны в настоящее время только в общедоступной предварительной версии (бета-версия).</span><span class="sxs-lookup"><span data-stu-id="93df1-112">These samples use APIs currently available only in public preview (beta).</span></span> <span data-ttu-id="93df1-113">В этих примерах требуется предварительная версия сборки для запуска.</span><span class="sxs-lookup"><span data-stu-id="93df1-113">These samples require preview builds to run.</span></span> <span data-ttu-id="93df1-114">Воспользуйтесь бета-версией библиотеки [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)или примите участие в программе[предварительная оценка Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="93df1-114">Either use the beta library of the [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) or join the [Office Insider program](https://products.office.com/office-insider).</span></span> <span data-ttu-id="93df1-115">Возможности сводных таблиц в настоящее время доступны в сборке 16.0.10801.20004.</span><span class="sxs-lookup"><span data-stu-id="93df1-115">PivotTable features are currently available in build 16.0.10801.20004.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="93df1-116">Иерархии</span><span class="sxs-lookup"><span data-stu-id="93df1-116">PKI Hierarchies</span></span>

<span data-ttu-id="93df1-117">Сводные таблицы организованы на основе четырёх категорий иерархии: строка, столбец, данные и фильтр.</span><span class="sxs-lookup"><span data-stu-id="93df1-117">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="93df1-118">В этой статье будут использоваться следующие данные, описывающие продажи фруктов из различных ферм.</span><span class="sxs-lookup"><span data-stu-id="93df1-118">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Коллекция продажи фруктов различных типов из нескольких ферм.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="93df1-120">Эти данные имеют пять иерархий: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span><span class="sxs-lookup"><span data-stu-id="93df1-120">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="93df1-121">Каждая иерархия может существовать только в одной из четырех категорий.</span><span class="sxs-lookup"><span data-stu-id="93df1-121">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="93df1-122">Если **Type** добавлен в иерархии столбца, затем он добавляется в иерархии строки и остается только в последних.</span><span class="sxs-lookup"><span data-stu-id="93df1-122">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="93df1-123">Иерархии строк и столбцов определяют группировку данных.</span><span class="sxs-lookup"><span data-stu-id="93df1-123">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="93df1-124">Например иерархии строки **Farms** будет сгруппировать все наборы данных из той же фермы.</span><span class="sxs-lookup"><span data-stu-id="93df1-124">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="93df1-125">Выбор между иерархиями строк и столбцов определяет ориентацию сводных таблиц.</span><span class="sxs-lookup"><span data-stu-id="93df1-125">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="93df1-126">Иерархии данных представляют собой значения, которые нужно сгруппировать в зависимости от иерархии строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="93df1-126">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="93df1-127">Сводная таблица с иерархиями строки **Farms** и данных **Crates Sold Wholesale** показывает общую сумму (по умолчанию) разных фруктов для каждой фермы.</span><span class="sxs-lookup"><span data-stu-id="93df1-127">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="93df1-128">Иерархии фильтра включают или исключаютданных изсводного документа, на основе значений в рамках отфильтрованного типа.</span><span class="sxs-lookup"><span data-stu-id="93df1-128">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="93df1-129">Иерархия фильтра **Classification** с только выбранным типом**Organic** отображает данные для органических фруктов.</span><span class="sxs-lookup"><span data-stu-id="93df1-129">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="93df1-130">Вот данные фермы, вместе с сводной таблицей.</span><span class="sxs-lookup"><span data-stu-id="93df1-130">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="93df1-131">Сводная таблица использует **Farm** и **Type**в качестве иерархий строк, **Crates Sold at Farm** и **Crates Sold Wholesale** - иерархий данных (с помощью статистической функции суммы по умолчанию) и **Classification** в качестве иерархий фильтра (с выбранным**Organic**).</span><span class="sxs-lookup"><span data-stu-id="93df1-131">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Выбор данных о продажах фруктов рядом со Сводной таблицей с иерархиями строки, данных и фильтра.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="93df1-133">Эта сводная таблица может быть создана через API JavaScript или с помощью пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="93df1-133">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="93df1-134">Оба параметра разрешают дальнейшую манипуляцию посредством надстроек.</span><span class="sxs-lookup"><span data-stu-id="93df1-134">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="93df1-135">Создание сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="93df1-135">Create a PivotTable or PivotChart report</span></span>

<span data-ttu-id="93df1-136">Сводной таблице требуются имя, источник и местом назначения.</span><span class="sxs-lookup"><span data-stu-id="93df1-136">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="93df1-137">Источником может быть адрес диапазона или имя таблицы (передается как `Range`, `string`, или `Table` типа).</span><span class="sxs-lookup"><span data-stu-id="93df1-137">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="93df1-138">Местом назначения является адреса диапазона (представленный в виде`Range` или `string`).</span><span class="sxs-lookup"><span data-stu-id="93df1-138">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="93df1-139">Следующие примеры показывают различные способы создания сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="93df1-139">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="93df1-140">Создание сводной таблицы с помощью адресов диапазона</span><span class="sxs-lookup"><span data-stu-id="93df1-140">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="93df1-141">Создание сводной таблицы с помощью объектов диапазона</span><span class="sxs-lookup"><span data-stu-id="93df1-141">Create a PivotTable with Range objects</span></span>

```typescript
await Excel.run(async (context) => {    
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);
    
    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="93df1-142">Создание сводной таблицы на уровне рабочей книги</span><span class="sxs-lookup"><span data-stu-id="93df1-142">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="93df1-143">Использование существующей сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="93df1-143">Use an existing PivotTable</span></span>

<span data-ttu-id="93df1-144">Созданные вручную сводные таблицы, также доступны через коллекцию сводной таблицы рабочей книги или отдельных листов.</span><span class="sxs-lookup"><span data-stu-id="93df1-144">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="93df1-145">Следующий код получает первую сводную таблицу в рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="93df1-145">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="93df1-146">Затем таблице дается  имя для простоты последующего использования.</span><span class="sxs-lookup"><span data-stu-id="93df1-146">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="93df1-147">Добавление строк и столбцов сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="93df1-147">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="93df1-148">Строки и столбцы сводной таблицы: данные, применимые к тем значениям полей.</span><span class="sxs-lookup"><span data-stu-id="93df1-148">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="93df1-149">Добавление сводной таблицы столбца всех продаж, применимых к каждой ферме **Farm**.</span><span class="sxs-lookup"><span data-stu-id="93df1-149">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="93df1-150">Добавление строк **Type** и **Classification** дополнительно разбивает данные на основе какие фрукты были проданы и была ли они органическими или нет.</span><span class="sxs-lookup"><span data-stu-id="93df1-150">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![Сводная таблица со столбцом фермы и строками типа и классификации.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

<span data-ttu-id="93df1-152">Вы также можете иметь сводную таблицу только строк или столбцов.</span><span class="sxs-lookup"><span data-stu-id="93df1-152">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="93df1-153">Добавление иерархий данных сводным таблицам</span><span class="sxs-lookup"><span data-stu-id="93df1-153">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="93df1-154">Иерархии данных заполняю т сводную таблицу сведениямина основе объединения строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="93df1-154">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="93df1-155">Добавление иерархий данных **Crates Sol at Farm** и **Crates Sold Wholesale** дает суммы эти цифр для каждой строки и столбца.</span><span class="sxs-lookup"><span data-stu-id="93df1-155">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="93df1-156">В примере, обе **Farm** и **Type** являются строками, с данными продаж ящиков.</span><span class="sxs-lookup"><span data-stu-id="93df1-156">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![Сводная таблица показывает сумму всех продаж разных фруктов в ферме, в зависимости от фермы их происхождения.](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the heirarchies that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="93df1-158">Изменение статистической функции</span><span class="sxs-lookup"><span data-stu-id="93df1-158">Change aggregation function</span></span>

<span data-ttu-id="93df1-159">Иерархии данных имеют свои агрегированные значения.</span><span class="sxs-lookup"><span data-stu-id="93df1-159">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="93df1-160">Для наборов данных чисел, это сумма по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="93df1-160">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="93df1-161">Свойство`summarizeBy` определяет это поведение на основе `AggregrationFunction` типа.</span><span class="sxs-lookup"><span data-stu-id="93df1-161">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="93df1-162">В настоящее время поддерживаются следующие типы статистической функции `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, и `Automatic` (по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="93df1-162">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="93df1-163">В следующих примерах кода изменяется агрегирование для средних значений данных.</span><span class="sxs-lookup"><span data-stu-id="93df1-163">The following code samples changes the aggregation to be averages of the data.</span></span>

```typescript
    await Excel.run(async (context) => {
        const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
        pivotTable.dataHierarchies.load("no-properties-needed");
        await context.sync();

        // changing the aggregation from the default sum to an average of all the values in the hierarchy
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        await context.sync();
    });
```

## <a name="pivottable-layouts"></a><span data-ttu-id="93df1-164">Макеты сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="93df1-164">PivotTable layouts</span></span>

<span data-ttu-id="93df1-165">Макет сводной таблицы определяет положение иерархий и их данных.</span><span class="sxs-lookup"><span data-stu-id="93df1-165">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="93df1-166">Доступ к макету для определения диапазонов, где хранятся данные.</span><span class="sxs-lookup"><span data-stu-id="93df1-166">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="93df1-167">На следующей диаграмме показано, какие функции макета соответствуют каким диапазонам из сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="93df1-167">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Диаграмма, показывающая, какие части сводной таблицы возвращают функции get диапазона макета.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="93df1-169">Приведенный ниже код показывает получение последней строки сводной таблицы данных через макет.</span><span class="sxs-lookup"><span data-stu-id="93df1-169">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="93df1-170">Затем эти значения суммируются друг с другом для общего итога.</span><span class="sxs-lookup"><span data-stu-id="93df1-170">Those values are then summed together for a grand total.</span></span>


```typescript
    await Excel.run(async (context) => {
        const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
        
        // get the totals for each data hierarchy from the layout
        const range = pivotTable.layout.getDataBodyRange();
        const grandTotalRange = range.getLastRow();
        grandTotalRange.load("address");
        await context.sync();
        
        // sum the totals from the PivotTable data hierarchies and place them in a new range
        const masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
        masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
        await context.sync();
    });
```

<span data-ttu-id="93df1-171">Сводные таблицы используются три стиля макета: сжатый, структурный и табличный.</span><span class="sxs-lookup"><span data-stu-id="93df1-171">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="93df1-172">Мы видели сжатый стиль в предыдущих примерах.</span><span class="sxs-lookup"><span data-stu-id="93df1-172">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="93df1-173">В приведенных ниже примерах используется, соответственнo, структурный и табличный стили, соответственно.</span><span class="sxs-lookup"><span data-stu-id="93df1-173">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="93df1-174">В примере кода показано, как переключаться между различные макетами.</span><span class="sxs-lookup"><span data-stu-id="93df1-174">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="93df1-175">Структурный макет</span><span class="sxs-lookup"><span data-stu-id="93df1-175">Outline layout</span></span>

![Использование макета структуры сводной таблицы.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="93df1-177">Табличный макет</span><span class="sxs-lookup"><span data-stu-id="93df1-177">Tabular layout</span></span>

![Использование макета таблицы сводной таблицы.](../images/excel-pivots-tabular-layout.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    await context.sync();
    
    // cycling through layout styles
    if (pivotTable.layout.layoutType === "Compact") {
        pivotTable.layout.layoutType = "Outline";
    } else if (pivotTable.layout.layoutType === "Outline") {
        pivotTable.layout.layoutType = "Tabular";
    } else {
        pivotTable.layout.layoutType = "Compact";
    }
    
    await context.sync();
});
```

## <a name="change-hierarchy-names"></a><span data-ttu-id="93df1-179">Изменение имен иерархий</span><span class="sxs-lookup"><span data-stu-id="93df1-179">Change hierarchy names</span></span>

<span data-ttu-id="93df1-180">Поля иерархии доступны для редактирования.</span><span class="sxs-lookup"><span data-stu-id="93df1-180">Hierarchy fields are editable.</span></span> <span data-ttu-id="93df1-181">Следующий код демонстрирует как изменить отображаемые имена двух иерархий данных.</span><span class="sxs-lookup"><span data-stu-id="93df1-181">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

```typescript
await Excel.run(async (context) => {
    const dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    await context.sync();
    
    // changing the displayed names of these entries
    dataHierarchies.items[0].name = "Farm Sales";
    dataHierarchies.items[1].name = "Wholesale";
    await context.sync();
});
```

## <a name="delete-a-pivottable"></a><span data-ttu-id="93df1-182">Удаление сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="93df1-182">Delete a PivotTable</span></span>

<span data-ttu-id="93df1-183">Сводные таблицы удаляются с помощью их имени.</span><span class="sxs-lookup"><span data-stu-id="93df1-183">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

> [!NOTE]
> <span data-ttu-id="93df1-184">Мы будем рады вашим отзывам и предложениям для разработок предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="93df1-184">We welcome feedback on our preview designs.</span></span> <span data-ttu-id="93df1-185">Если у вас есть комментарии, предложения или проблемы, связанные с новой API сводной таблицы, оставьте свои комментарии на репозитории[UserVoice](https://officespdev.uservoice.com/forums/224641-feature-requests-and-feedback?category_id=163563) или [repo OpenSpec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec).</span><span class="sxs-lookup"><span data-stu-id="93df1-185">If you have comments, suggestions, or issues with the new PivotTable API, please leave your comments on [UserVoice](https://officespdev.uservoice.com/forums/224641-feature-requests-and-feedback?category_id=163563) or on the [OpenSpec GitHub repo](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec).</span></span>

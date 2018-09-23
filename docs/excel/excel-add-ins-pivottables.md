---
title: Работа со сводными таблицами с помощью API JavaScript для Excel
description: Использование Excel JavaScript API для создания сводных таблиц и взаимодействия с их компонентами.
ms.date: 09/21/2018
ms.openlocfilehash: b8704389ced3686858f488b2a50f80c22b1b8bd6
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967671"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="6c662-103">Работа со сводными таблицами с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="6c662-103">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="6c662-104">Сводные таблицы упрощают большие наборы данных.</span><span class="sxs-lookup"><span data-stu-id="6c662-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="6c662-105">Они позволяют производить быструю манипуляцию сгруппированных данных.</span><span class="sxs-lookup"><span data-stu-id="6c662-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="6c662-106">API JavaScript для Excel позволяет надстройке создавать сводные таблицы и взаимодействовать с их компонентами.</span><span class="sxs-lookup"><span data-stu-id="6c662-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="6c662-107">Если вы не знакомы с функциональностью сводных таблиц, рекомендуем рассмотреть их в качестве конечного пользователя.</span><span class="sxs-lookup"><span data-stu-id="6c662-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="6c662-108">См. [Создание сводной таблицы для анализа данных листа](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) как учебник для начинающих на этих инструментах.</span><span class="sxs-lookup"><span data-stu-id="6c662-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="6c662-109">В этой статье приведены примеры кода для распространенных сценариев.</span><span class="sxs-lookup"><span data-stu-id="6c662-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="6c662-110">Подробнее об API сводных таблиц см. статьи [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) и [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="6c662-110">To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6c662-111">Сводные таблицы, созданные с помощью OLAP, в настоящее время не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="6c662-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="6c662-112">Иерархии</span><span class="sxs-lookup"><span data-stu-id="6c662-112">PKI Hierarchies</span></span>

<span data-ttu-id="6c662-113">Сводные таблицы организованы на основе четырёх категорий иерархии: строка, столбец, данные и фильтр.</span><span class="sxs-lookup"><span data-stu-id="6c662-113">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="6c662-114">В этой статье будут использоваться следующие данные, описывающие продажи фруктов из различных ферм.</span><span class="sxs-lookup"><span data-stu-id="6c662-114">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Коллекция продажи фруктов различных типов из нескольких ферм.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="6c662-116">Эти данные имеют пять иерархий: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span><span class="sxs-lookup"><span data-stu-id="6c662-116">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="6c662-117">Каждая иерархия может существовать только в одной из четырех категорий.</span><span class="sxs-lookup"><span data-stu-id="6c662-117">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="6c662-118">Если **Type** добавлен в иерархии столбца, затем он добавляется в иерархии строки и остается только в последних.</span><span class="sxs-lookup"><span data-stu-id="6c662-118">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="6c662-119">Иерархии строк и столбцов определяют группировку данных.</span><span class="sxs-lookup"><span data-stu-id="6c662-119">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="6c662-120">Например иерархии строки **Farms** будет сгруппировать все наборы данных из той же фермы.</span><span class="sxs-lookup"><span data-stu-id="6c662-120">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="6c662-121">Выбор между иерархиями строк и столбцов определяет ориентацию сводных таблиц.</span><span class="sxs-lookup"><span data-stu-id="6c662-121">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="6c662-122">Иерархии данных представляют собой значения, которые нужно сгруппировать в зависимости от иерархии строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="6c662-122">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="6c662-123">Сводная таблица с иерархиями строки **Farms** и данных **Crates Sold Wholesale** показывает общую сумму (по умолчанию) разных фруктов для каждой фермы.</span><span class="sxs-lookup"><span data-stu-id="6c662-123">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="6c662-124">Иерархии фильтра включают или исключаютданных изсводного документа, на основе значений в рамках отфильтрованного типа.</span><span class="sxs-lookup"><span data-stu-id="6c662-124">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="6c662-125">Иерархия фильтра **Classification** с только выбранным типом**Organic** отображает данные для органических фруктов.</span><span class="sxs-lookup"><span data-stu-id="6c662-125">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="6c662-126">Вот данные фермы, вместе с сводной таблицей.</span><span class="sxs-lookup"><span data-stu-id="6c662-126">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="6c662-127">Сводная таблица использует **Farm** и **Type**в качестве иерархий строк, **Crates Sold at Farm** и **Crates Sold Wholesale** - иерархий данных (с помощью статистической функции суммы по умолчанию) и **Classification** в качестве иерархий фильтра (с выбранным**Organic**).</span><span class="sxs-lookup"><span data-stu-id="6c662-127">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Выбор данных о продажах фруктов рядом со Сводной таблицей с иерархиями строки, данных и фильтра.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="6c662-129">Эта сводная таблица может быть создана через API JavaScript или с помощью пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="6c662-129">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="6c662-130">Оба параметра разрешают дальнейшую манипуляцию посредством надстроек.</span><span class="sxs-lookup"><span data-stu-id="6c662-130">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="6c662-131">Создание сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="6c662-131">Create a PivotTable with Range objects</span></span>

<span data-ttu-id="6c662-132">Сводной таблице требуются имя, источник и местом назначения.</span><span class="sxs-lookup"><span data-stu-id="6c662-132">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="6c662-133">Источником может быть адрес диапазона или имя таблицы (передается как `Range`, `string`, или `Table` типа).</span><span class="sxs-lookup"><span data-stu-id="6c662-133">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="6c662-134">Местом назначения является адреса диапазона (представленный в виде`Range` или `string`).</span><span class="sxs-lookup"><span data-stu-id="6c662-134">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="6c662-135">Следующие примеры показывают различные способы создания сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="6c662-135">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="6c662-136">Создание сводной таблицы с помощью адресов диапазона</span><span class="sxs-lookup"><span data-stu-id="6c662-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="6c662-137">Создание сводной таблицы с помощью объектов диапазона</span><span class="sxs-lookup"><span data-stu-id="6c662-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="6c662-138">Создание сводной таблицы на уровне рабочей книги</span><span class="sxs-lookup"><span data-stu-id="6c662-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="6c662-139">Использование существующей сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="6c662-139">Use an existing PivotTable</span></span>

<span data-ttu-id="6c662-140">Созданные вручную сводные таблицы, также доступны через коллекцию сводной таблицы рабочей книги или отдельных листов.</span><span class="sxs-lookup"><span data-stu-id="6c662-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="6c662-141">Следующий код получает первую сводную таблицу в рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="6c662-141">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="6c662-142">Затем таблице дается  имя для простоты последующего использования.</span><span class="sxs-lookup"><span data-stu-id="6c662-142">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="6c662-143">Добавление строк и столбцов сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="6c662-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="6c662-144">Строки и столбцы сводной таблицы: данные, применимые к тем значениям полей.</span><span class="sxs-lookup"><span data-stu-id="6c662-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="6c662-145">Добавление сводной таблицы столбца всех продаж, применимых к каждой ферме **Farm**.</span><span class="sxs-lookup"><span data-stu-id="6c662-145">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="6c662-146">Добавление строк **Type** и **Classification** дополнительно разбивает данные на основе какие фрукты были проданы и была ли они органическими или нет.</span><span class="sxs-lookup"><span data-stu-id="6c662-146">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="6c662-148">Вы также можете иметь сводную таблицу только строк или столбцов.</span><span class="sxs-lookup"><span data-stu-id="6c662-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="6c662-149">Добавление иерархий данных сводным таблицам</span><span class="sxs-lookup"><span data-stu-id="6c662-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="6c662-150">Иерархии данных заполняю т сводную таблицу сведениямина основе объединения строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="6c662-150">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="6c662-151">Добавление иерархий данных **Crates Sol at Farm** и **Crates Sold Wholesale** дает суммы эти цифр для каждой строки и столбца.</span><span class="sxs-lookup"><span data-stu-id="6c662-151">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="6c662-152">В примере, обе **Farm** и **Type** являются строками, с данными продаж ящиков.</span><span class="sxs-lookup"><span data-stu-id="6c662-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="6c662-154">Изменение статистической функции</span><span class="sxs-lookup"><span data-stu-id="6c662-154">Change aggregation function</span></span>

<span data-ttu-id="6c662-155">Иерархии данных имеют свои агрегированные значения.</span><span class="sxs-lookup"><span data-stu-id="6c662-155">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="6c662-156">Для наборов данных чисел, это сумма по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="6c662-156">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="6c662-157">Свойство`summarizeBy` определяет это поведение на основе `AggregrationFunction` типа.</span><span class="sxs-lookup"><span data-stu-id="6c662-157">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="6c662-158">В настоящее время поддерживаются следующие типы статистической функции `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, и `Automatic` (по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="6c662-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="6c662-159">В следующих примерах кода изменяется агрегирование для средних значений данных.</span><span class="sxs-lookup"><span data-stu-id="6c662-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="pivottable-layouts"></a><span data-ttu-id="6c662-160">Макеты сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="6c662-160">PivotTable layouts</span></span>

<span data-ttu-id="6c662-161">Макет сводной таблицы определяет положение иерархий и их данных.</span><span class="sxs-lookup"><span data-stu-id="6c662-161">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="6c662-162">Доступ к макету для определения диапазонов, где хранятся данные.</span><span class="sxs-lookup"><span data-stu-id="6c662-162">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="6c662-163">На следующей диаграмме показано, какие функции макета соответствуют каким диапазонам из сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="6c662-163">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Диаграмма, показывающая, какие части сводной таблицы возвращают функции get диапазона макета.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="6c662-165">Приведенный ниже код показывает получение последней строки сводной таблицы данных через макет.</span><span class="sxs-lookup"><span data-stu-id="6c662-165">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="6c662-166">Затем эти значения суммируются друг с другом для общего итога.</span><span class="sxs-lookup"><span data-stu-id="6c662-166">Those values are then summed together for a grand total.</span></span>


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

<span data-ttu-id="6c662-167">Сводные таблицы используются три стиля макета: сжатый, структурный и табличный.</span><span class="sxs-lookup"><span data-stu-id="6c662-167">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="6c662-168">Мы видели сжатый стиль в предыдущих примерах.</span><span class="sxs-lookup"><span data-stu-id="6c662-168">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="6c662-169">В приведенных ниже примерах используется, соответственнo, структурный и табличный стили, соответственно.</span><span class="sxs-lookup"><span data-stu-id="6c662-169">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="6c662-170">В примере кода показано, как переключаться между различные макетами.</span><span class="sxs-lookup"><span data-stu-id="6c662-170">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="6c662-171">Структурный макет</span><span class="sxs-lookup"><span data-stu-id="6c662-171">Outline layout</span></span>

![Использование макета структуры сводной таблицы.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="6c662-173">Табличный макет</span><span class="sxs-lookup"><span data-stu-id="6c662-173">Tabular layout</span></span>

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="6c662-175">Изменение имен иерархий</span><span class="sxs-lookup"><span data-stu-id="6c662-175">Change hierarchy names</span></span>

<span data-ttu-id="6c662-176">Поля иерархии доступны для редактирования.</span><span class="sxs-lookup"><span data-stu-id="6c662-176">Hierarchy fields are editable.</span></span> <span data-ttu-id="6c662-177">Следующий код демонстрирует как изменить отображаемые имена двух иерархий данных.</span><span class="sxs-lookup"><span data-stu-id="6c662-177">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="6c662-178">Удаление сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="6c662-178">Delete a PivotTable</span></span>

<span data-ttu-id="6c662-179">Сводная таблица удаляется по имени.</span><span class="sxs-lookup"><span data-stu-id="6c662-179">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="6c662-180">См. также</span><span class="sxs-lookup"><span data-stu-id="6c662-180">See also</span></span>

- [<span data-ttu-id="6c662-181">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="6c662-181">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6c662-182">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="6c662-182">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)

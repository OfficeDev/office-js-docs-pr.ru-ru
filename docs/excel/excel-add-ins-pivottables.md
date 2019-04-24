---
title: Работать со сводными таблицами с помощью API JavaScript для Excel
description: Используйте API JavaScript для Excel, чтобы создавать сводные таблицы и взаимодействовать с их компонентами.
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: b53d734e676417a6438f1008bac720a38a244d1f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449386"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="26c03-103">Работать со сводными таблицами с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="26c03-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="26c03-104">Сводные таблицы упрощают работу с большими наборами данных.</span><span class="sxs-lookup"><span data-stu-id="26c03-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="26c03-105">Они позволяют быстро управлять группированием данных.</span><span class="sxs-lookup"><span data-stu-id="26c03-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="26c03-106">API JavaScript для Excel позволяет надстройке создавать сводные таблицы и взаимодействовать с их компонентами.</span><span class="sxs-lookup"><span data-stu-id="26c03-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span>

<span data-ttu-id="26c03-107">Если вы не знакомы с функциями сводных таблиц, рассмотрите возможность их изучения в качестве конечного пользователя.</span><span class="sxs-lookup"><span data-stu-id="26c03-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span> <span data-ttu-id="26c03-108">Ознакомьтесь со статьей [Создание сводной таблицы, чтобы проанализировать данные листа](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) для хорошего учебника по этим средствам.</span><span class="sxs-lookup"><span data-stu-id="26c03-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="26c03-109">В этой статье приведены примеры кода для распространенных сценариев.</span><span class="sxs-lookup"><span data-stu-id="26c03-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="26c03-110">Подробнее об API сводных таблиц можно узнать в статье [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="26c03-110">To further your understanding of the PivotTable API, see [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="26c03-111">Сводные таблицы, созданные с помощью OLAP, в настоящее время не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="26c03-111">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="26c03-112">Кроме того, отсутствует поддержка Power Pivot.</span><span class="sxs-lookup"><span data-stu-id="26c03-112">There is also no support for Power Pivot.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="26c03-113">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="26c03-113">Hierarchies</span></span>

<span data-ttu-id="26c03-114">Сводные таблицы организованы в соответствии с четырьмя категориями иерархии: строкой, столбцом, данными и фильтром.</span><span class="sxs-lookup"><span data-stu-id="26c03-114">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="26c03-115">В этой статье будут использоваться следующие данные, описывающие продажи фруктов из различных ферм.</span><span class="sxs-lookup"><span data-stu-id="26c03-115">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Коллекция продаж фруктов различных типов из различных ферм.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="26c03-117">Эти данные имеют пять иерархий: **ферм**, **типов**, **классификаций**, ящиков, **проданных в ферме**, и ящики, продаваемые **оптовой торговлей**.</span><span class="sxs-lookup"><span data-stu-id="26c03-117">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="26c03-118">Каждая иерархия может существовать только в одной из четырех категорий.</span><span class="sxs-lookup"><span data-stu-id="26c03-118">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="26c03-119">Если **тип** добавляется к иерархиям столбцов и затем добавляется к иерархиям строк, он остается только последним.</span><span class="sxs-lookup"><span data-stu-id="26c03-119">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="26c03-120">Иерархии строк и столбцов определяют, как группируются данные.</span><span class="sxs-lookup"><span data-stu-id="26c03-120">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="26c03-121">Например, иерархия **ферм фермы** объединяет все наборы данных из одной фермы.</span><span class="sxs-lookup"><span data-stu-id="26c03-121">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="26c03-122">Выбор между строкой и иерархией столбцов определяет ориентацию сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="26c03-122">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="26c03-123">Иерархии данных — это значения, которые должны быть объединены на основе иерархий строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="26c03-123">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="26c03-124">Сводная таблица с иерархией **ферм** и иерархией данных для ящиков, проданных в **оптовой торговле** , показывает общую сумму (по умолчанию) всех различных Fruits для каждой фермы.</span><span class="sxs-lookup"><span data-stu-id="26c03-124">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="26c03-125">Иерархии фильтров включают или исключают данные из сводной таблицы на основе значений в этом типе фильтрации.</span><span class="sxs-lookup"><span data-stu-id="26c03-125">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="26c03-126">Иерархия фильтров **классификации** с типом "не \*\*\*\* только выбранные" показывает только данные для придля себя фруктов.</span><span class="sxs-lookup"><span data-stu-id="26c03-126">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="26c03-127">Далее представлены данные фермы, вместе со сводной таблицей.</span><span class="sxs-lookup"><span data-stu-id="26c03-127">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="26c03-128">В сводной таблице используется **ферма** и **тип** в качестве иерархий строк, ящики, проданные **на ферме** и ящики, проданные по **оптовой торговле** в виде иерархий данных (с статистической функцией статистической обработки по умолчанию Sum), а **классификация** — как фильтр. иерархия ( \*\*\*\* с выделенным параметром).</span><span class="sxs-lookup"><span data-stu-id="26c03-128">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Выбор данных о продажах для фруктов рядом со сводной таблицей со строками, данными и иерархиями фильтров.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="26c03-130">Эту сводную таблицу можно создать с помощью API JavaScript или ПОЛЬЗОВАТЕЛЬСКОГО интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="26c03-130">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="26c03-131">Оба варианта позволяют осуществлять дальнейшую обработку надстроек.</span><span class="sxs-lookup"><span data-stu-id="26c03-131">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="26c03-132">Создание сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="26c03-132">Create a PivotTable</span></span>

<span data-ttu-id="26c03-133">Для сводных таблиц требуются имя, источник и назначение.</span><span class="sxs-lookup"><span data-stu-id="26c03-133">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="26c03-134">Источником может быть адрес диапазона или имя таблицы (передается как тип `Range`, `string`или `Table` тип).</span><span class="sxs-lookup"><span data-stu-id="26c03-134">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="26c03-135">Назначение является адресом диапазона ( `Range` или `string`).</span><span class="sxs-lookup"><span data-stu-id="26c03-135">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="26c03-136">В следующих примерах показаны различные методы создания сводных таблиц.</span><span class="sxs-lookup"><span data-stu-id="26c03-136">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="26c03-137">Создание сводной таблицы с адресами диапазона</span><span class="sxs-lookup"><span data-stu-id="26c03-137">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="26c03-138">Создание сводной таблицы с объектами Range</span><span class="sxs-lookup"><span data-stu-id="26c03-138">Create a PivotTable with Range objects</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
        "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="26c03-139">Создание сводной таблицы на уровне книги</span><span class="sxs-lookup"><span data-stu-id="26c03-139">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="26c03-140">Использование существующей сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="26c03-140">Use an existing PivotTable</span></span>

<span data-ttu-id="26c03-141">Вы также можете получить доступ к сводным таблицам, созданным вручную, с помощью сводной таблицы книги или отдельных листов.</span><span class="sxs-lookup"><span data-stu-id="26c03-141">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="26c03-142">Приведенный ниже код получает первую сводную таблицу в книге.</span><span class="sxs-lookup"><span data-stu-id="26c03-142">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="26c03-143">Затем имя таблицы придается имени для упрощения справочных материалов.</span><span class="sxs-lookup"><span data-stu-id="26c03-143">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="26c03-144">Добавление строк и столбцов в сводную таблицу</span><span class="sxs-lookup"><span data-stu-id="26c03-144">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="26c03-145">Строки и столбцы поворачивают данные вокруг этих значений полей.</span><span class="sxs-lookup"><span data-stu-id="26c03-145">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="26c03-146">При добавлении столбца **фермы** все продажи для каждой фермы отворачиваются.</span><span class="sxs-lookup"><span data-stu-id="26c03-146">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="26c03-147">Добавление строк **типа** и **классификации** дополнительно разделяет данные на основании того, сколько фруктов было продано, и не было ли оно согласовано.</span><span class="sxs-lookup"><span data-stu-id="26c03-147">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![Сводная таблица со столбцами фермы, а также строками типов и классификации.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

<span data-ttu-id="26c03-149">Кроме того, можно создать сводную таблицу, используя только строки или столбцы.</span><span class="sxs-lookup"><span data-stu-id="26c03-149">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="26c03-150">Добавление иерархий данных в сводную таблицу</span><span class="sxs-lookup"><span data-stu-id="26c03-150">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="26c03-151">Иерархии данных заполняют сводную таблицу со сведениями, которые необходимо объединить в зависимости от строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="26c03-151">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="26c03-152">Добавление иерархий данных ящиков, проданных **в ферме** и ящиков, продаваемых в **оптовой торговле** , приводит к суммированию этих значений для каждой строки и столбца.</span><span class="sxs-lookup"><span data-stu-id="26c03-152">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="26c03-153">В этом примере **ферма** и **тип** представляют собой строки, в которых продажи ящиков являются данными.</span><span class="sxs-lookup"><span data-stu-id="26c03-153">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![Сводная таблица, в которой показаны общие продажи разных фруктов на основе фермы, из которой они получены.](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="26c03-155">Изменение статистической функции</span><span class="sxs-lookup"><span data-stu-id="26c03-155">Change aggregation function</span></span>

<span data-ttu-id="26c03-156">Иерархия данных содержит статистические значения.</span><span class="sxs-lookup"><span data-stu-id="26c03-156">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="26c03-157">Для наборов данных Numbers это сумма по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="26c03-157">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="26c03-158">`summarizeBy` Свойство определяет это поведение на основе типа [аггрегатионфунктион](/javascript/api/excel/excel.aggregationfunction) .</span><span class="sxs-lookup"><span data-stu-id="26c03-158">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="26c03-159">`Sum`В настоящее время поддерживаются типы статистической `Count`функции `Average`, `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP`,,,,,,,, и `Automatic` (значение по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="26c03-159">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="26c03-160">В приведенных ниже примерах кода статистическая схема изменяется для средних значений данных.</span><span class="sxs-lookup"><span data-stu-id="26c03-160">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="26c03-161">Изменение вычислений с помощью Шовасруле</span><span class="sxs-lookup"><span data-stu-id="26c03-161">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="26c03-162">Сводные таблицы по умолчанию объединяют данные иерархий строк и столбцов независимо друг от друга.</span><span class="sxs-lookup"><span data-stu-id="26c03-162">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="26c03-163">[Шовасруле](/javascript/api/excel/excel.showasrule) изменяет иерархию данных на выходные значения на основе других элементов в сводной таблице.</span><span class="sxs-lookup"><span data-stu-id="26c03-163">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="26c03-164">У `ShowAsRule` объекта есть три свойства:</span><span class="sxs-lookup"><span data-stu-id="26c03-164">The `ShowAsRule` object has three properties:</span></span>

-   <span data-ttu-id="26c03-165">`calculation`: Тип относительного вычисления, применяемого к иерархии данных (значение по умолчанию — `none`).</span><span class="sxs-lookup"><span data-stu-id="26c03-165">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="26c03-166">`baseField`: Поле в иерархии, содержащее базовые данные перед применением вычисления.</span><span class="sxs-lookup"><span data-stu-id="26c03-166">`baseField`: The field within the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="26c03-167">[PivotField](/javascript/api/excel/excel.pivotfield) обычно имеет то же имя, что и его родительская иерархия.</span><span class="sxs-lookup"><span data-stu-id="26c03-167">The [PivotField](/javascript/api/excel/excel.pivotfield) usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="26c03-168">`baseItem`: Отдельные [PivotItem](/javascript/api/excel/excel.pivotitem) по сравнению со значениями базовых полей на основе типа вычисления.</span><span class="sxs-lookup"><span data-stu-id="26c03-168">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="26c03-169">Для этого поля требуется не все вычисления.</span><span class="sxs-lookup"><span data-stu-id="26c03-169">Not all calculations require this field.</span></span>

<span data-ttu-id="26c03-170">В следующем примере показана настройка вычисления **суммы ящиков, проданных в** иерархии данных фермы, в процентах от общей суммы по столбцу.</span><span class="sxs-lookup"><span data-stu-id="26c03-170">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span> <span data-ttu-id="26c03-171">Мы по-прежнему хотим, чтобы гранулярность была расширена до уровня типа фруктов, поэтому мы будем использовать иерархию **типов** строк и базовое поле.</span><span class="sxs-lookup"><span data-stu-id="26c03-171">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span> <span data-ttu-id="26c03-172">В примере также используется **ферма** в качестве первой иерархии строк, поэтому записи итоговой фермы отображаются в процентах, ответственных за изготовление.</span><span class="sxs-lookup"><span data-stu-id="26c03-172">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![Сводная таблица, в которой показаны процентные доли продаж фруктов относительно общего итога для отдельных ферм и отдельных типов фруктов в каждой ферме.](../images/excel-pivots-showas-percentage.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the crates of each fruit type sold at the farm as a percentage of the column's total
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
    farmDataHierarchy.showAs = farmShowAs; 
    farmDataHierarchy.name = "Percentage of Total Farm Sales";

    await context.sync();
});
```

<span data-ttu-id="26c03-174">В предыдущем примере показано, как задать вычисление для столбца относительно иерархии отдельных строк.</span><span class="sxs-lookup"><span data-stu-id="26c03-174">The previous example set the calculation to the column, relative to an individual row hierarchy.</span></span> <span data-ttu-id="26c03-175">Когда расчет относится к отдельному элементу, используйте `baseItem` свойство.</span><span class="sxs-lookup"><span data-stu-id="26c03-175">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="26c03-176">В приведенном ниже примере `differenceFrom` показано вычисление.</span><span class="sxs-lookup"><span data-stu-id="26c03-176">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="26c03-177">В нем отображается разность записей иерархии данных о продажах в ферме, относящихся к параметрам "фермы".</span><span class="sxs-lookup"><span data-stu-id="26c03-177">It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”.</span></span>
<span data-ttu-id="26c03-178">Ферма `baseField` состоит \*\*\*\* в том, что мы видим различия между другими фермами, а также подразделение для каждого типа вроде фруктов (**тип** также является иерархией строк в данном примере).</span><span class="sxs-lookup"><span data-stu-id="26c03-178">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Сводная таблица, в которой показаны различия продаж фруктов между "фермами" и другими.](../images/excel-pivots-showas-differencefrom.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the difference between crate sales of the "A Farms" and the other farms
    // this difference is both aggregated and shown for individual fruit types (where applicable)
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
    farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
    farmDataHierarchy.showAs = farmShowAs;
    farmDataHierarchy.name = "Difference from A Farms";
    await context.sync();
});
```

## <a name="pivottable-layouts"></a><span data-ttu-id="26c03-182">Макеты сводных таблиц</span><span class="sxs-lookup"><span data-stu-id="26c03-182">PivotTable layouts</span></span>

<span data-ttu-id="26c03-183">[PivotLayout](/javascript/api/excel/excel.pivotlayout) определяет размещение иерархий и их данных.</span><span class="sxs-lookup"><span data-stu-id="26c03-183">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="26c03-184">Вы можете получить доступ к макету, чтобы определить диапазоны, в которых хранятся данные.</span><span class="sxs-lookup"><span data-stu-id="26c03-184">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="26c03-185">На следующей схеме показано, какие вызовы функций макета соответствуют какому диапазону сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="26c03-185">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Схема, на которой показано, какие разделы сводной таблицы возвращаются функциями диапазона получения в макете.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="26c03-187">В приведенном ниже коде показано, как получить последнюю строку данных сводной таблицы, прополнив макет.</span><span class="sxs-lookup"><span data-stu-id="26c03-187">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="26c03-188">Затем эти значения суммируются вместе для общего итога.</span><span class="sxs-lookup"><span data-stu-id="26c03-188">Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="26c03-189">В сводных таблицах есть три стиля макета: компактный, структурированный и табличный.</span><span class="sxs-lookup"><span data-stu-id="26c03-189">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="26c03-190">В предыдущих примерах показан стиль "Компактный".</span><span class="sxs-lookup"><span data-stu-id="26c03-190">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="26c03-191">В приведенных ниже примерах используются структурированные и табличные стили соответственно.</span><span class="sxs-lookup"><span data-stu-id="26c03-191">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="26c03-192">В примере кода показано, как циклически переключаться между различными макетами.</span><span class="sxs-lookup"><span data-stu-id="26c03-192">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="26c03-193">Макет структуры</span><span class="sxs-lookup"><span data-stu-id="26c03-193">Outline layout</span></span>

![Сводная таблица с использованием структуры.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="26c03-195">Табличный макет</span><span class="sxs-lookup"><span data-stu-id="26c03-195">Tabular layout</span></span>

![Сводная таблица с использованием табличного макета.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="26c03-197">Изменение имен иерархий</span><span class="sxs-lookup"><span data-stu-id="26c03-197">Change hierarchy names</span></span>

<span data-ttu-id="26c03-198">Поля иерархии можно редактировать.</span><span class="sxs-lookup"><span data-stu-id="26c03-198">Hierarchy fields are editable.</span></span> <span data-ttu-id="26c03-199">В приведенном ниже коде показано, как изменить отображаемые имена двух иерархий данных.</span><span class="sxs-lookup"><span data-stu-id="26c03-199">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="26c03-200">Удаление сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="26c03-200">Delete a PivotTable</span></span>

<span data-ttu-id="26c03-201">Сводные таблицы удаляются с использованием их имени.</span><span class="sxs-lookup"><span data-stu-id="26c03-201">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="26c03-202">См. также</span><span class="sxs-lookup"><span data-stu-id="26c03-202">See also</span></span>

- [<span data-ttu-id="26c03-203">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="26c03-203">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="26c03-204">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="26c03-204">Excel JavaScript API Reference</span></span>](/javascript/api/excel)

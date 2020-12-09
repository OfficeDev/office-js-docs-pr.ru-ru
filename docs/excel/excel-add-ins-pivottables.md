---
title: Работать со сводными таблицами с помощью API JavaScript для Excel
description: Используйте API JavaScript для Excel, чтобы создавать сводные таблицы и взаимодействовать с их компонентами.
ms.date: 12/07/2020
localization_priority: Normal
ms.openlocfilehash: 0a1fefa6a855ab9ee1ccd71fd0dc60f282d2944b
ms.sourcegitcommit: fecad2afa7938d7178456c11ba52b558224813b4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/09/2020
ms.locfileid: "49603801"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="9418b-103">Работать со сводными таблицами с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="9418b-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="9418b-104">Сводные таблицы упрощают работу с большими наборами данных.</span><span class="sxs-lookup"><span data-stu-id="9418b-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="9418b-105">Они позволяют быстро управлять группированием данных.</span><span class="sxs-lookup"><span data-stu-id="9418b-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="9418b-106">API JavaScript для Excel позволяет надстройке создавать сводные таблицы и взаимодействовать с их компонентами.</span><span class="sxs-lookup"><span data-stu-id="9418b-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="9418b-107">В этой статье описывается, как сводные таблицы представлены с помощью API JavaScript для Office, а также приведены примеры кода для ключевых сценариев.</span><span class="sxs-lookup"><span data-stu-id="9418b-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="9418b-108">Если вы не знакомы с функциями сводных таблиц, рассмотрите возможность их изучения в качестве конечного пользователя.</span><span class="sxs-lookup"><span data-stu-id="9418b-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="9418b-109">Ознакомьтесь со статьей [Создание сводной таблицы, чтобы проанализировать данные листа](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) для хорошего учебника по этим средствам.</span><span class="sxs-lookup"><span data-stu-id="9418b-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9418b-110">Сводные таблицы, созданные с помощью OLAP, в настоящее время не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="9418b-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="9418b-111">Кроме того, отсутствует поддержка Power Pivot.</span><span class="sxs-lookup"><span data-stu-id="9418b-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="9418b-112">Объектная модель</span><span class="sxs-lookup"><span data-stu-id="9418b-112">Object model</span></span>

<span data-ttu-id="9418b-113">[Сводная таблица](/javascript/api/excel/excel.pivottable) является центральным объектом для сводных ТАБЛИЦ в API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="9418b-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="9418b-114">`Workbook.pivotTables` и `Worksheet.pivotTables` — это [пивоттаблеколлектионс](/javascript/api/excel/excel.pivottablecollection) , которые содержат [Сводные таблицы](/javascript/api/excel/excel.pivottable) в книге и листе соответственно.</span><span class="sxs-lookup"><span data-stu-id="9418b-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="9418b-115">[Сводная таблица](/javascript/api/excel/excel.pivottable) содержит [Пивосиерарчиколлектион](/javascript/api/excel/excel.pivothierarchycollection) с несколькими [пивосиерарчиес](/javascript/api/excel/excel.pivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="9418b-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="9418b-116">Эти [пивосиерарчиес](/javascript/api/excel/excel.pivothierarchy) можно добавить в конкретные коллекции иерархий, чтобы определить, как данные будут сведены в сводную таблицу (как описано в [следующем разделе](#hierarchies)).</span><span class="sxs-lookup"><span data-stu-id="9418b-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="9418b-117">[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) содержит [пивотфиелдколлектион](/javascript/api/excel/excel.pivotfieldcollection) , в котором есть ровно один [PivotField](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="9418b-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="9418b-118">Если проект разворачивается для включения сводных таблиц OLAP, это может измениться.</span><span class="sxs-lookup"><span data-stu-id="9418b-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="9418b-119">К [PivotField](/javascript/api/excel/excel.pivotfield) может быть применено одно или несколько [PivotFilters](/javascript/api/excel/excel.pivotfilters) , если [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) поля назначено категории иерархии.</span><span class="sxs-lookup"><span data-stu-id="9418b-119">A [PivotField](/javascript/api/excel/excel.pivotfield) can have one or more [PivotFilters](/javascript/api/excel/excel.pivotfilters) applied, as long as the field's [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) is assigned to a hierarchy category.</span></span> 
- <span data-ttu-id="9418b-120">[PivotField](/javascript/api/excel/excel.pivotfield) содержит [Пивотитемколлектион](/javascript/api/excel/excel.pivotitemcollection) с несколькими [PivotItems](/javascript/api/excel/excel.pivotitem).</span><span class="sxs-lookup"><span data-stu-id="9418b-120">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="9418b-121">[Сводная таблица](/javascript/api/excel/excel.pivottable) содержит объект [PivotLayout](/javascript/api/excel/excel.pivotlayout) , определяющий, где на листе отображаются [PivotFields](/javascript/api/excel/excel.pivotfield) и [PivotItems](/javascript/api/excel/excel.pivotitem) .</span><span class="sxs-lookup"><span data-stu-id="9418b-121">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span>

<span data-ttu-id="9418b-122">Рассмотрим, как эти отношения применяются к некоторым примерам данных.</span><span class="sxs-lookup"><span data-stu-id="9418b-122">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="9418b-123">В приведенных ниже данных описываются продажи фруктов из различных ферм.</span><span class="sxs-lookup"><span data-stu-id="9418b-123">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="9418b-124">Это будет пример во всей этой статье.</span><span class="sxs-lookup"><span data-stu-id="9418b-124">It will be the example throughout this article.</span></span>

![Коллекция продаж фруктов различных типов из различных ферм.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="9418b-126">Данные продаж фермы фруктов будут использоваться для создания сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="9418b-126">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="9418b-127">Каждый столбец, например **types**, — это `PivotHierarchy` .</span><span class="sxs-lookup"><span data-stu-id="9418b-127">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="9418b-128">Иерархия **types** содержит поле **типы** .</span><span class="sxs-lookup"><span data-stu-id="9418b-128">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="9418b-129">Поле **типы** содержит элементы **Apple**, **киви**, **Лемон**, **травяные** и **оранжевые**.</span><span class="sxs-lookup"><span data-stu-id="9418b-129">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="9418b-130">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="9418b-130">Hierarchies</span></span>

<span data-ttu-id="9418b-131">Сводные таблицы организованы в соответствии с четырьмя категориями иерархии: [строкой](/javascript/api/excel/excel.rowcolumnpivothierarchy), [столбцом](/javascript/api/excel/excel.rowcolumnpivothierarchy), [данными](/javascript/api/excel/excel.datapivothierarchy)и [фильтром](/javascript/api/excel/excel.filterpivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="9418b-131">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="9418b-132">Приведенные выше данные фермы имеют пять иерархий: **фермы**, **типы**, **классификации**, **ящики**, проданные в ферме и **ящики, продаваемые оптовой торговлей**.</span><span class="sxs-lookup"><span data-stu-id="9418b-132">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="9418b-133">Каждая иерархия может существовать только в одной из четырех категорий.</span><span class="sxs-lookup"><span data-stu-id="9418b-133">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="9418b-134">Если **тип** добавляется к иерархиям столбцов, он также не может находиться в иерархиях "строка", "данные" или "Фильтрация".</span><span class="sxs-lookup"><span data-stu-id="9418b-134">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="9418b-135">Если впоследствии **тип** добавляется к иерархиям строк, он удаляется из иерархий столбцов.</span><span class="sxs-lookup"><span data-stu-id="9418b-135">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="9418b-136">Такое поведение аналогично тому, как выполняется назначение иерархии с помощью пользовательского интерфейса Excel или API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="9418b-136">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="9418b-137">Иерархии строк и столбцов определяют, как группируются данные.</span><span class="sxs-lookup"><span data-stu-id="9418b-137">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="9418b-138">Например, иерархия **ферм фермы** объединяет все наборы данных из одной фермы.</span><span class="sxs-lookup"><span data-stu-id="9418b-138">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="9418b-139">Выбор между строкой и иерархией столбцов определяет ориентацию сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="9418b-139">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="9418b-140">Иерархии данных — это значения, которые должны быть объединены на основе иерархий строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="9418b-140">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="9418b-141">Сводная таблица с иерархией **ферм** и иерархией данных для ящиков, проданных в **оптовой торговле** , показывает общую сумму (по умолчанию) всех различных Fruits для каждой фермы.</span><span class="sxs-lookup"><span data-stu-id="9418b-141">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="9418b-142">Иерархии фильтров включают или исключают данные из сводной таблицы на основе значений в этом типе фильтрации.</span><span class="sxs-lookup"><span data-stu-id="9418b-142">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="9418b-143">Иерархия фильтров **классификации** **с типом "** не только выбранные" показывает только данные для придля себя фруктов.</span><span class="sxs-lookup"><span data-stu-id="9418b-143">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="9418b-144">Далее представлены данные фермы, вместе со сводной таблицей.</span><span class="sxs-lookup"><span data-stu-id="9418b-144">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="9418b-145">В сводной таблице используется **ферма** и **тип** в качестве иерархий строк, **ящики** , проданные в ферме и ящики, проданные в ферме, а также **продаются по оптовой торговле** в виде иерархий данных (с использованием статистической функции по умолчанию Sum) и **классификации** в качестве иерархии фильтров ( **с выбранным** параметром "</span><span class="sxs-lookup"><span data-stu-id="9418b-145">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![Выбор данных о продажах для фруктов рядом со сводной таблицей со строками, данными и иерархиями фильтров.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="9418b-147">Эту сводную таблицу можно создать с помощью API JavaScript или пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="9418b-147">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="9418b-148">Оба варианта позволяют осуществлять дальнейшую обработку надстроек.</span><span class="sxs-lookup"><span data-stu-id="9418b-148">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="9418b-149">Создание сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="9418b-149">Create a PivotTable</span></span>

<span data-ttu-id="9418b-150">Для сводных таблиц требуются имя, источник и назначение.</span><span class="sxs-lookup"><span data-stu-id="9418b-150">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="9418b-151">Источником может быть адрес диапазона или имя таблицы (передается как `Range` тип, `string` или `Table` тип).</span><span class="sxs-lookup"><span data-stu-id="9418b-151">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="9418b-152">Назначение является адресом диапазона ( `Range` или `string` ).</span><span class="sxs-lookup"><span data-stu-id="9418b-152">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="9418b-153">В следующих примерах показаны различные методы создания сводных таблиц.</span><span class="sxs-lookup"><span data-stu-id="9418b-153">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="9418b-154">Создание сводной таблицы с адресами диапазона</span><span class="sxs-lookup"><span data-stu-id="9418b-154">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="9418b-155">Создание сводной таблицы с объектами Range</span><span class="sxs-lookup"><span data-stu-id="9418b-155">Create a PivotTable with Range objects</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21.
    var rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    var rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
      "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    return context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="9418b-156">Создание сводной таблицы на уровне книги</span><span class="sxs-lookup"><span data-stu-id="9418b-156">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="9418b-157">Использование существующей сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="9418b-157">Use an existing PivotTable</span></span>

<span data-ttu-id="9418b-158">Вы также можете получить доступ к сводным таблицам, созданным вручную, с помощью сводной таблицы книги или отдельных листов.</span><span class="sxs-lookup"><span data-stu-id="9418b-158">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="9418b-159">В следующем коде показано получение сводной таблицы с именем **My Pivot** из книги.</span><span class="sxs-lookup"><span data-stu-id="9418b-159">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="9418b-160">Добавление строк и столбцов в сводную таблицу</span><span class="sxs-lookup"><span data-stu-id="9418b-160">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="9418b-161">Строки и столбцы поворачивают данные вокруг этих значений полей.</span><span class="sxs-lookup"><span data-stu-id="9418b-161">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="9418b-162">При добавлении столбца **фермы** все продажи для каждой фермы отворачиваются.</span><span class="sxs-lookup"><span data-stu-id="9418b-162">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="9418b-163">Добавление строк **типа** и **классификации** дополнительно разделяет данные на основании того, сколько фруктов было продано, и не было ли оно согласовано.</span><span class="sxs-lookup"><span data-stu-id="9418b-163">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![Сводная таблица со столбцами фермы, а также строками типов и классификации.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

<span data-ttu-id="9418b-165">Кроме того, можно создать сводную таблицу, используя только строки или столбцы.</span><span class="sxs-lookup"><span data-stu-id="9418b-165">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="9418b-166">Добавление иерархий данных в сводную таблицу</span><span class="sxs-lookup"><span data-stu-id="9418b-166">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="9418b-167">Иерархии данных заполняют сводную таблицу со сведениями, которые необходимо объединить в зависимости от строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="9418b-167">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="9418b-168">Добавление иерархий данных ящиков, проданных **в ферме** и **ящиков, продаваемых в оптовой торговле** , приводит к суммированию этих значений для каждой строки и столбца.</span><span class="sxs-lookup"><span data-stu-id="9418b-168">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="9418b-169">В этом примере **ферма** и **тип** представляют собой строки, в которых продажи ящиков являются данными.</span><span class="sxs-lookup"><span data-stu-id="9418b-169">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

![Сводная таблица, в которой показаны общие продажи разных фруктов на основе фермы, из которой они получены.](../images/excel-pivots-data-hierarchy.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based.
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case).
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    return context.sync();
});
```

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="9418b-171">Макеты сводных таблиц и извлечение сведенных данных</span><span class="sxs-lookup"><span data-stu-id="9418b-171">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="9418b-172">[PivotLayout](/javascript/api/excel/excel.pivotlayout) определяет размещение иерархий и их данных.</span><span class="sxs-lookup"><span data-stu-id="9418b-172">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="9418b-173">Вы можете получить доступ к макету, чтобы определить диапазоны, в которых хранятся данные.</span><span class="sxs-lookup"><span data-stu-id="9418b-173">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="9418b-174">На следующей схеме показано, какие вызовы функций макета соответствуют какому диапазону сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="9418b-174">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Схема, на которой показано, какие разделы сводной таблицы возвращаются функциями диапазона получения в макете.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="9418b-176">Получение данных из сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="9418b-176">Get data from the PivotTable</span></span>

<span data-ttu-id="9418b-177">Макет определяет способ отображения сводной таблицы на листе.</span><span class="sxs-lookup"><span data-stu-id="9418b-177">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="9418b-178">Это означает, что `PivotLayout` объект управляет диапазонами, используемыми для элементов сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="9418b-178">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="9418b-179">Используйте диапазоны, предоставленные макетом, для получения данных, собранных и агрегированных сводной таблицей.</span><span class="sxs-lookup"><span data-stu-id="9418b-179">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="9418b-180">В частности, используйте `PivotLayout.getDataBodyRange` для доступа к тем, что делает Сводная таблица.</span><span class="sxs-lookup"><span data-stu-id="9418b-180">In particular, use `PivotLayout.getDataBodyRange` to access what the PivotTable produces.</span></span>

<span data-ttu-id="9418b-181">В приведенном ниже коде показано, как получить последнюю строку данных сводной таблицы, посвященную макету ( **общему** количеству **ящиков, проданных в ферме** , и **сумме ящиков, проданных** в одной колонке в предыдущем примере).</span><span class="sxs-lookup"><span data-stu-id="9418b-181">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="9418b-182">Затем эти значения суммируются вместе для итогового итога, который отображается в ячейке **E30** (вне сводной таблицы).</span><span class="sxs-lookup"><span data-stu-id="9418b-182">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range, outside of the PivotTable.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("E30");
        masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

### <a name="layout-types"></a><span data-ttu-id="9418b-183">Типы макетов</span><span class="sxs-lookup"><span data-stu-id="9418b-183">Layout types</span></span>

<span data-ttu-id="9418b-184">В сводных таблицах есть три стиля макета: компактный, структурированный и табличный.</span><span class="sxs-lookup"><span data-stu-id="9418b-184">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="9418b-185">В предыдущих примерах показан стиль "Компактный".</span><span class="sxs-lookup"><span data-stu-id="9418b-185">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="9418b-186">В приведенных ниже примерах используются структурированные и табличные стили соответственно.</span><span class="sxs-lookup"><span data-stu-id="9418b-186">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="9418b-187">В примере кода показано, как циклически переключаться между различными макетами.</span><span class="sxs-lookup"><span data-stu-id="9418b-187">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="9418b-188">Макет структуры</span><span class="sxs-lookup"><span data-stu-id="9418b-188">Outline layout</span></span>

![Сводная таблица с использованием структуры.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="9418b-190">Табличный макет</span><span class="sxs-lookup"><span data-stu-id="9418b-190">Tabular layout</span></span>

![Сводная таблица с использованием табличного макета.](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a><span data-ttu-id="9418b-192">Удаление сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="9418b-192">Delete a PivotTable</span></span>

<span data-ttu-id="9418b-193">Сводные таблицы удаляются с использованием их имени.</span><span class="sxs-lookup"><span data-stu-id="9418b-193">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a><span data-ttu-id="9418b-194">Фильтрация сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="9418b-194">Filter a PivotTable</span></span>

<span data-ttu-id="9418b-195">Основным методом фильтрации данных сводной таблицы является PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="9418b-195">The primary method for filtering PivotTable data is with PivotFilters.</span></span> <span data-ttu-id="9418b-196">Срезы предоставляют альтернативный и менее гибкий метод фильтрации.</span><span class="sxs-lookup"><span data-stu-id="9418b-196">Slicers offer an alternate, less flexible filtering method.</span></span> 

<span data-ttu-id="9418b-197">[PivotFilters](/javascript/api/excel/excel.pivotfilters) фильтрация данных на основе четырех [иерархических категорий](#hierarchies) сводной таблицы (фильтров, столбцов, строк и значений).</span><span class="sxs-lookup"><span data-stu-id="9418b-197">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filter data based on a PivotTable's four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="9418b-198">Существует четыре типа PivotFilters, позволяющие использовать фильтрацию на основе дат, анализ строк, сравнение чисел и фильтрацию на основе настраиваемого ввода.</span><span class="sxs-lookup"><span data-stu-id="9418b-198">There are four types of PivotFilters, allowing calendar date-based filtering, string parsing, number comparison, and filtering based on a custom input.</span></span> 

<span data-ttu-id="9418b-199">[Срезы](/javascript/api/excel/excel.slicer) можно применять как к сводным таблицам, так и к обычным таблицам Excel.</span><span class="sxs-lookup"><span data-stu-id="9418b-199">[Slicers](/javascript/api/excel/excel.slicer) can be applied to both PivotTables and regular Excel tables.</span></span> <span data-ttu-id="9418b-200">При применении к сводной таблице срезы функционируют так же, как и [пивотмануалфилтер](#pivotmanualfilter) , и позволяют выполнять фильтрацию на основе настраиваемого ввода.</span><span class="sxs-lookup"><span data-stu-id="9418b-200">When applied to a PivotTable, slicers function like a [PivotManualFilter](#pivotmanualfilter) and allow filtering based on a custom input.</span></span> <span data-ttu-id="9418b-201">В отличие от PivotFilters, срезы имеют [компонент пользовательского интерфейса Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span><span class="sxs-lookup"><span data-stu-id="9418b-201">Unlike PivotFilters, slicers have an [Excel UI component](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span></span> <span data-ttu-id="9418b-202">С помощью `Slicer` класса вы создадите этот компонент пользовательского интерфейса, управляете фильтрацией и контролируйте его внешний вид.</span><span class="sxs-lookup"><span data-stu-id="9418b-202">With the `Slicer` class, you create this UI component, manage filtering, and control its visual appearance.</span></span> 

### <a name="filter-with-pivotfilters"></a><span data-ttu-id="9418b-203">Фильтрация с помощью PivotFilters</span><span class="sxs-lookup"><span data-stu-id="9418b-203">Filter with PivotFilters</span></span>

<span data-ttu-id="9418b-204">[PivotFilters](/javascript/api/excel/excel.pivotfilters) позволяют фильтровать данные сводной таблицы на основе четырех [категорий иерархии](#hierarchies) (фильтров, столбцов, строк и значений).</span><span class="sxs-lookup"><span data-stu-id="9418b-204">[PivotFilters](/javascript/api/excel/excel.pivotfilters) allow you to filter PivotTable data based on the four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="9418b-205">В объектной модели сводной таблицы `PivotFilters` применяются к [PivotField](/javascript/api/excel/excel.pivotfield), и у каждого из них `PivotField` может быть один или несколько назначенных `PivotFilters` .</span><span class="sxs-lookup"><span data-stu-id="9418b-205">In the PivotTable object model, `PivotFilters` are applied to a [PivotField](/javascript/api/excel/excel.pivotfield), and each `PivotField` can have one or more assigned `PivotFilters`.</span></span> <span data-ttu-id="9418b-206">Чтобы применить PivotFilters к PivotField, соответствующему [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) поля необходимо назначить категории иерархии.</span><span class="sxs-lookup"><span data-stu-id="9418b-206">To apply PivotFilters to a PivotField, the field's corresponding [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) must be assigned to a hierarchy category.</span></span> 

#### <a name="types-of-pivotfilters"></a><span data-ttu-id="9418b-207">Типы PivotFilters</span><span class="sxs-lookup"><span data-stu-id="9418b-207">Types of PivotFilters</span></span>

| <span data-ttu-id="9418b-208">Тип фильтра</span><span class="sxs-lookup"><span data-stu-id="9418b-208">Filter type</span></span> | <span data-ttu-id="9418b-209">Назначение фильтра</span><span class="sxs-lookup"><span data-stu-id="9418b-209">Filter purpose</span></span> | <span data-ttu-id="9418b-210">Справочные материалы по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="9418b-210">Excel JavaScript API reference</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="9418b-211">датефилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-211">DateFilter</span></span> | <span data-ttu-id="9418b-212">Фильтрация на основе даты в календаре.</span><span class="sxs-lookup"><span data-stu-id="9418b-212">Calendar date-based filtering.</span></span> | [<span data-ttu-id="9418b-213">пивотдатефилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-213">PivotDateFilter</span></span>](/javascript/api/excel/excel.pivotdatefilter) |
| <span data-ttu-id="9418b-214">лабелфилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-214">LabelFilter</span></span> | <span data-ttu-id="9418b-215">Фильтрация по текстовому сравнению.</span><span class="sxs-lookup"><span data-stu-id="9418b-215">Text comparison filtering.</span></span> | [<span data-ttu-id="9418b-216">пивотлабелфилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-216">PivotLabelFilter</span></span>](/javascript/api/excel/excel.pivotlabelfilter) |
| <span data-ttu-id="9418b-217">мануалфилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-217">ManualFilter</span></span> | <span data-ttu-id="9418b-218">Настраиваемый фильтр ввода.</span><span class="sxs-lookup"><span data-stu-id="9418b-218">Custom input filtering.</span></span> | [<span data-ttu-id="9418b-219">пивотмануалфилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-219">PivotManualFilter</span></span>](/javascript/api/excel/excel.pivotmanualfilter) |
| <span data-ttu-id="9418b-220">валуефилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-220">ValueFilter</span></span> | <span data-ttu-id="9418b-221">Фильтрация сравнения номеров.</span><span class="sxs-lookup"><span data-stu-id="9418b-221">Number comparison filtering.</span></span> | [<span data-ttu-id="9418b-222">пивотвалуефилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-222">PivotValueFilter</span></span>](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a><span data-ttu-id="9418b-223">Создание PivotFilter</span><span class="sxs-lookup"><span data-stu-id="9418b-223">Create a PivotFilter</span></span>

<span data-ttu-id="9418b-224">Чтобы отфильтровать данные сводной таблицы с помощью сводного фильтра (например, Пивотдатефилтер), примените фильтр к [PivotField](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="9418b-224">To filter PivotTable data with a Pivot\*Filter (such as a PivotDateFilter), apply the filter to a [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="9418b-225">В следующих четырех примерах кода показано, как использовать каждый из четырех типов PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="9418b-225">The following four code samples show how to use each of the four types of PivotFilters.</span></span> 

##### <a name="pivotdatefilter"></a><span data-ttu-id="9418b-226">пивотдатефилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-226">PivotDateFilter</span></span>

<span data-ttu-id="9418b-227">Первый пример кода применяет [пивотдатефилтер](/javascript/api/excel/excel.pivotdatefilter) к **дате обновления** PivotField, скрывая все данные до **2020-08-01**.</span><span class="sxs-lookup"><span data-stu-id="9418b-227">The first code sample applies a [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) to the **Date Updated** PivotField, hiding any data prior to **2020-08-01**.</span></span> 

> [!IMPORTANT] 
> <span data-ttu-id="9418b-228">Фильтр PIVOT нельзя применить к PivotField, если это поле PivotHierarchy не назначено категории иерархии.</span><span class="sxs-lookup"><span data-stu-id="9418b-228">A Pivot\*Filter can't be applied to a PivotField unless that field's PivotHierarchy is assigned to a hierarchy category.</span></span> <span data-ttu-id="9418b-229">В следующем примере кода `dateHierarchy` необходимо добавить в категорию сводной таблицы, `rowHierarchies` прежде чем его можно будет использовать для фильтрации.</span><span class="sxs-lookup"><span data-stu-id="9418b-229">In the following code sample, the `dateHierarchy` must be added to the PivotTable's `rowHierarchies` category before it can be used for filtering.</span></span>

```js
Excel.run(function (context) {
    // Get the PivotTable and the date hierarchy.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date Updated");
    
    return context.sync().then(function () {
        // PivotFilters can only be applied to PivotHierarchies that are being used for pivoting.
        // If it's not already there, add "Date Updated" to the hierarchies.
        if (dateHierarchy.isNullObject) {
          dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));
        }

        // Apply a date filter to filter out anything logged before August.
        var filterField = dateHierarchy.fields.getItem("Date Updated");
        var dateFilter = {
          condition: Excel.DateFilterCondition.afterOrEqualTo,
          comparator: {
            date: "2020-08-01",
            specificity: Excel.FilterDatetimeSpecificity.month
          }
        };
        filterField.applyFilter({ dateFilter: dateFilter });
        
        return context.sync();
    });
});
```

> [!NOTE]
> <span data-ttu-id="9418b-230">В следующих трех фрагментах кода отображаются только отрывок, относящиеся к фильтрам, а не полные `Excel.run` вызовы.</span><span class="sxs-lookup"><span data-stu-id="9418b-230">The following three code snippets only display filter-specific excerpts, instead of full `Excel.run` calls.</span></span>

##### <a name="pivotlabelfilter"></a><span data-ttu-id="9418b-231">пивотлабелфилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-231">PivotLabelFilter</span></span>

<span data-ttu-id="9418b-232">Во втором фрагменте кода показано, как применить [пивотлабелфилтер](/javascript/api/excel/excel.pivotlabelfilter) к **типу** PivotField, используя свойство, `LabelFilterCondition.beginsWith` чтобы исключить метки, начинающиеся с буквы **L**.</span><span class="sxs-lookup"><span data-stu-id="9418b-232">The second code snippet demonstrates how to apply a [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) to the **Type** PivotField, using the `LabelFilterCondition.beginsWith` property to exclude labels that start with the letter **L**.</span></span> 

```js
    // Get the "Type" field.
    var filterField = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");

    // Filter out any types that start with "L" ("Lemons" and "Limes" in this case).
    var filter: Excel.PivotLabelFilter = {
      condition: Excel.LabelFilterCondition.beginsWith,
      substring: "L",
      exclusive: true
    };

    // Apply the label filter to the field.
    filterField.applyFilter({ labelFilter: filter });
```

##### <a name="pivotmanualfilter"></a><span data-ttu-id="9418b-233">пивотмануалфилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-233">PivotManualFilter</span></span>

<span data-ttu-id="9418b-234">Третий фрагмент кода применяет к полю **классификации** вручную фильтр с [пивотмануалфилтер](/javascript/api/excel/excel.pivotmanualfilter) , отфильтровывая данные, которые не включают согласованности классификации **.**</span><span class="sxs-lookup"><span data-stu-id="9418b-234">The third code snippet applies a manual filter with [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) to the the **Classification** field, filtering out data that doesn't include the classification **Organic**.</span></span> 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a><span data-ttu-id="9418b-235">пивотвалуефилтер</span><span class="sxs-lookup"><span data-stu-id="9418b-235">PivotValueFilter</span></span>

<span data-ttu-id="9418b-236">Чтобы сравнить числа, используйте фильтр значений с [пивотвалуефилтер](/javascript/api/excel/excel.pivotvaluefilter), как показано в последнем фрагменте кода.</span><span class="sxs-lookup"><span data-stu-id="9418b-236">To compare numbers, use a value filter with [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), as shown in the final code snippet.</span></span> <span data-ttu-id="9418b-237">В этом `PivotValueFilter` разделе сравниваются данные в **ферме** PivotField с данными в рабочих ящиках, проданных в **оптовой торговле** PivotField, включая только те фермы, сумма которых проданных ящиков превышает значение **500**.</span><span class="sxs-lookup"><span data-stu-id="9418b-237">The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Wholesale** PivotField, including only farms whose sum of crates sold exceeds the value **500**.</span></span> 

```js
    // Get the "Farm" field.
    var filterField = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");
    
    // Filter to only include rows with more than 500 wholesale crates sold.
    var filter: Excel.PivotValueFilter = {
      condition: Excel.ValueFilterCondition.greaterThan,
      comparator: 500,
      value: "Sum of Crates Sold Wholesale"
    };
    
    // Apply the value filter to the field.
    filterField.applyFilter({ valueFilter: filter });
```

#### <a name="remove-pivotfilters"></a><span data-ttu-id="9418b-238">Удаление PivotFilters</span><span class="sxs-lookup"><span data-stu-id="9418b-238">Remove PivotFilters</span></span>

<span data-ttu-id="9418b-239">Чтобы удалить все PivotFilters, примените `clearAllFilters` метод к каждому PivotField, как показано в следующем примере кода.</span><span class="sxs-lookup"><span data-stu-id="9418b-239">To remove all PivotFilters, apply the `clearAllFilters` method to each PivotField, as shown in the following code sample.</span></span> 

```js
Excel.run(function (context) {
    // Get the PivotTable.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.hierarchies.load("name");
    
    return context.sync().then(function () {
        // Clear the filters on each PivotField.
        pivotTable.hierarchies.items.forEach(function (hierarchy) {
          hierarchy.fields.getItem(hierarchy.name).clearAllFilters();
        });
        return context.sync();
    });
});
```

### <a name="filter-with-slicers"></a><span data-ttu-id="9418b-240">Фильтрация с помощью срезов</span><span class="sxs-lookup"><span data-stu-id="9418b-240">Filter with slicers</span></span>

<span data-ttu-id="9418b-241">[Срезы](/javascript/api/excel/excel.slicer) позволяют фильтровать данные из сводной таблицы или таблицы Excel.</span><span class="sxs-lookup"><span data-stu-id="9418b-241">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="9418b-242">Срез использует значения из указанного столбца или PivotField для фильтрации соответствующих строк.</span><span class="sxs-lookup"><span data-stu-id="9418b-242">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="9418b-243">Эти значения хранятся в виде объектов [SlicerItem](/javascript/api/excel/excel.sliceritem) в `Slicer` .</span><span class="sxs-lookup"><span data-stu-id="9418b-243">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="9418b-244">Надстройка может настраивать эти фильтры, как это могут делать пользователи ([через пользовательский интерфейс Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span><span class="sxs-lookup"><span data-stu-id="9418b-244">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="9418b-245">Срез располагается вверху листа в графическом слое, как показано на следующем снимке экрана.</span><span class="sxs-lookup"><span data-stu-id="9418b-245">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![Фильтрация данных среза в сводной таблице.](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="9418b-247">Методы, описанные в этом разделе, касаются использования срезов, подключенных к сводным таблицам.</span><span class="sxs-lookup"><span data-stu-id="9418b-247">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="9418b-248">Те же методы применяются и для использования срезов, подключенных к таблицам.</span><span class="sxs-lookup"><span data-stu-id="9418b-248">The same techniques also apply to using slicers connected to tables.</span></span>

#### <a name="create-a-slicer"></a><span data-ttu-id="9418b-249">Создание среза</span><span class="sxs-lookup"><span data-stu-id="9418b-249">Create a slicer</span></span>

<span data-ttu-id="9418b-250">Вы можете создать срез в книге или листе с помощью `Workbook.slicers.add` метода или `Worksheet.slicers.add` метода.</span><span class="sxs-lookup"><span data-stu-id="9418b-250">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="9418b-251">Это приведет к добавлению среза в [слицерколлектион](/javascript/api/excel/excel.slicercollection) указанного `Workbook` или `Worksheet` объекта.</span><span class="sxs-lookup"><span data-stu-id="9418b-251">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="9418b-252">`SlicerCollection.add`Метод имеет три параметра:</span><span class="sxs-lookup"><span data-stu-id="9418b-252">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="9418b-253">`slicerSource`: Источник данных, на котором основан новый срез.</span><span class="sxs-lookup"><span data-stu-id="9418b-253">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="9418b-254">Это может быть `PivotTable` , `Table` или строка, представляющая имя или идентификатор или `PivotTable` `Table` .</span><span class="sxs-lookup"><span data-stu-id="9418b-254">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="9418b-255">`sourceField`: Поле в источнике данных, с помощью которого выполняется фильтрация.</span><span class="sxs-lookup"><span data-stu-id="9418b-255">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="9418b-256">Это может быть `PivotField` , `TableColumn` или строка, представляющая имя или идентификатор или `PivotField` `TableColumn` .</span><span class="sxs-lookup"><span data-stu-id="9418b-256">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="9418b-257">`slicerDestination`: Лист, на котором будет создан новый срез.</span><span class="sxs-lookup"><span data-stu-id="9418b-257">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="9418b-258">Это может быть `Worksheet` объект или имя или идентификатор объекта `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="9418b-258">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="9418b-259">Этот параметр не является обязательным при `SlicerCollection` доступе к `Worksheet.slicers` .</span><span class="sxs-lookup"><span data-stu-id="9418b-259">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="9418b-260">В этом случае лист коллекции используется в качестве назначения.</span><span class="sxs-lookup"><span data-stu-id="9418b-260">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="9418b-261">В приведенном ниже примере кода в **сводную** таблицу добавляется новый срез.</span><span class="sxs-lookup"><span data-stu-id="9418b-261">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="9418b-262">Источник среза — это сводная таблица и фильтры **продаж фермы** с использованием данных **типа** .</span><span class="sxs-lookup"><span data-stu-id="9418b-262">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="9418b-263">Срез также называется **срезом фруктов** для дальнейшего использования.</span><span class="sxs-lookup"><span data-stu-id="9418b-263">The slicer is also named **Fruit Slicer** for future reference.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Pivot");
    var slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    return context.sync();
});
```

#### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="9418b-264">Фильтрация элементов с помощью среза</span><span class="sxs-lookup"><span data-stu-id="9418b-264">Filter items with a slicer</span></span>

<span data-ttu-id="9418b-265">Срез фильтрует сводную таблицу с элементами из `sourceField` .</span><span class="sxs-lookup"><span data-stu-id="9418b-265">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="9418b-266">`Slicer.selectItems`Метод задает элементы, остающиеся в срезе.</span><span class="sxs-lookup"><span data-stu-id="9418b-266">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="9418b-267">Эти элементы передаются в метод как объект `string[]` , представляющий ключи элементов.</span><span class="sxs-lookup"><span data-stu-id="9418b-267">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="9418b-268">Все строки, содержащие эти элементы, сохраняются в статистической обработке сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="9418b-268">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="9418b-269">Последующие вызовы `selectItems` задают для списка ключи, указанные в этих вызовах.</span><span class="sxs-lookup"><span data-stu-id="9418b-269">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="9418b-270">Если `Slicer.selectItems` передается элемент, который не находится в источнике данных, `InvalidArgument` возникает ошибка.</span><span class="sxs-lookup"><span data-stu-id="9418b-270">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="9418b-271">Содержимое можно проверить с помощью `Slicer.slicerItems` свойства, которое является [слицеритемколлектион](/javascript/api/excel/excel.sliceritemcollection).</span><span class="sxs-lookup"><span data-stu-id="9418b-271">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="9418b-272">В приведенном ниже примере кода показаны три выбранных для среза элементов: **Лемон**, **травяной** и **оранжевый**.</span><span class="sxs-lookup"><span data-stu-id="9418b-272">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="9418b-273">Чтобы удалить все фильтры из среза, используйте `Slicer.clearFilters` метод, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="9418b-273">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a><span data-ttu-id="9418b-274">Стиль и форматирование среза</span><span class="sxs-lookup"><span data-stu-id="9418b-274">Style and format a slicer</span></span>

<span data-ttu-id="9418b-275">Надстройка может настраивать параметры отображения среза с помощью `Slicer` свойств.</span><span class="sxs-lookup"><span data-stu-id="9418b-275">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="9418b-276">В приведенном ниже примере кода для стиля задается значение **SlicerStyleLight6**, в верхней части среза задается **Тип фруктов**, помещается срез в позицию **(395, 15)** на уровне рисунка и задается размер среза **135x150** пикселей.</span><span class="sxs-lookup"><span data-stu-id="9418b-276">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.caption = "Fruit Types";
    slicer.left = 395;
    slicer.top = 15;
    slicer.height = 135;
    slicer.width = 150;
    slicer.style = "SlicerStyleLight6";
    return context.sync();
});
```

#### <a name="delete-a-slicer"></a><span data-ttu-id="9418b-277">Удаление среза</span><span class="sxs-lookup"><span data-stu-id="9418b-277">Delete a slicer</span></span>

<span data-ttu-id="9418b-278">Чтобы удалить срез, вызовите `Slicer.delete` метод.</span><span class="sxs-lookup"><span data-stu-id="9418b-278">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="9418b-279">В примере кода ниже показано, как удалить первый срез из текущего листа.</span><span class="sxs-lookup"><span data-stu-id="9418b-279">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="9418b-280">Изменение статистической функции</span><span class="sxs-lookup"><span data-stu-id="9418b-280">Change aggregation function</span></span>

<span data-ttu-id="9418b-281">Иерархия данных содержит статистические значения.</span><span class="sxs-lookup"><span data-stu-id="9418b-281">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="9418b-282">Для наборов данных Numbers это сумма по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9418b-282">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="9418b-283">`summarizeBy`Свойство определяет это поведение на основе типа [аггрегатионфунктион](/javascript/api/excel/excel.aggregationfunction) .</span><span class="sxs-lookup"><span data-stu-id="9418b-283">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="9418b-284">В настоящее время поддерживаются типы статистической функции,,,,,,,,, `Sum` `Count` `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` и `Automatic` (значение по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="9418b-284">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="9418b-285">В приведенных ниже примерах кода статистическая схема изменяется для средних значений данных.</span><span class="sxs-lookup"><span data-stu-id="9418b-285">The following code samples changes the aggregation to be averages of the data.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    return context.sync().then(function() {

        // Change the aggregation from the default sum to an average of all the values in the hierarchy.
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        return context.sync();
    });
});
```

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="9418b-286">Изменение вычислений с помощью Шовасруле</span><span class="sxs-lookup"><span data-stu-id="9418b-286">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="9418b-287">Сводные таблицы по умолчанию объединяют данные иерархий строк и столбцов независимо друг от друга.</span><span class="sxs-lookup"><span data-stu-id="9418b-287">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="9418b-288">[Шовасруле](/javascript/api/excel/excel.showasrule) изменяет иерархию данных на выходные значения на основе других элементов в сводной таблице.</span><span class="sxs-lookup"><span data-stu-id="9418b-288">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="9418b-289">`ShowAsRule`У объекта есть три свойства:</span><span class="sxs-lookup"><span data-stu-id="9418b-289">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="9418b-290">`calculation`: Тип относительного вычисления, применяемого к иерархии данных (значение по умолчанию — `none` ).</span><span class="sxs-lookup"><span data-stu-id="9418b-290">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="9418b-291">`baseField`: [PivotField](/javascript/api/excel/excel.pivotfield) в иерархии, содержащей базовые данные перед применением вычисления.</span><span class="sxs-lookup"><span data-stu-id="9418b-291">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="9418b-292">Так как сводные таблицы Excel имеют сопоставление "один к одному" в поле "иерархия", для доступа к иерархии и полю используется то же имя.</span><span class="sxs-lookup"><span data-stu-id="9418b-292">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="9418b-293">`baseItem`: Отдельные [PivotItem](/javascript/api/excel/excel.pivotitem) по сравнению со значениями базовых полей на основе типа вычисления.</span><span class="sxs-lookup"><span data-stu-id="9418b-293">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="9418b-294">Для этого поля требуется не все вычисления.</span><span class="sxs-lookup"><span data-stu-id="9418b-294">Not all calculations require this field.</span></span>

<span data-ttu-id="9418b-295">В следующем примере показана настройка вычисления **суммы ящиков, проданных в** иерархии данных фермы, в процентах от общей суммы по столбцу.</span><span class="sxs-lookup"><span data-stu-id="9418b-295">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="9418b-296">Мы по-прежнему хотим, чтобы гранулярность была расширена до уровня типа фруктов, поэтому мы будем использовать иерархию **типов** строк и базовое поле.</span><span class="sxs-lookup"><span data-stu-id="9418b-296">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="9418b-297">В примере также используется **ферма** в качестве первой иерархии строк, поэтому записи итоговой фермы отображаются в процентах, ответственных за изготовление.</span><span class="sxs-lookup"><span data-stu-id="9418b-297">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![Сводная таблица, в которой показаны процентные доли продаж фруктов относительно общего итога для отдельных ферм и отдельных типов фруктов в каждой ферме.](../images/excel-pivots-showas-percentage.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {

        // Show the crates of each fruit type sold at the farm as a percentage of the column's total.
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Percentage of Total Farm Sales";
    });
});
```

<span data-ttu-id="9418b-299">В предыдущем примере показано, как задать вычисление для столбца относительно поля отдельной иерархии строк.</span><span class="sxs-lookup"><span data-stu-id="9418b-299">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="9418b-300">Когда расчет относится к отдельному элементу, используйте `baseItem` свойство.</span><span class="sxs-lookup"><span data-stu-id="9418b-300">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="9418b-301">В приведенном ниже примере показано `differenceFrom` вычисление.</span><span class="sxs-lookup"><span data-stu-id="9418b-301">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="9418b-302">В нем отображается разность записей иерархии данных о продажах в ферме, относящихся к параметрам **ферм**.</span><span class="sxs-lookup"><span data-stu-id="9418b-302">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="9418b-303">`baseField` **Ферма** состоит в том, что мы видим различия между другими фермами, а также подразделение для каждого типа вроде фруктов (**тип** также является иерархией строк в данном примере).</span><span class="sxs-lookup"><span data-stu-id="9418b-303">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Сводная таблица, в которой показаны различия продаж фруктов между "фермами" и другими.](../images/excel-pivots-showas-differencefrom.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {
        // Show the difference between crate sales of the "A Farms" and the other farms.
        // This difference is both aggregated and shown for individual fruit types (where applicable).
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
        farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Difference from A Farms";
    });
});
```

## <a name="change-hierarchy-names"></a><span data-ttu-id="9418b-307">Изменение имен иерархий</span><span class="sxs-lookup"><span data-stu-id="9418b-307">Change hierarchy names</span></span>

<span data-ttu-id="9418b-308">Поля иерархии можно редактировать.</span><span class="sxs-lookup"><span data-stu-id="9418b-308">Hierarchy fields are editable.</span></span> <span data-ttu-id="9418b-309">В приведенном ниже коде показано, как изменить отображаемые имена двух иерархий данных.</span><span class="sxs-lookup"><span data-stu-id="9418b-309">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

```js
Excel.run(function (context) {
    var dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    return context.sync().then(function () {
        // changing the displayed names of these entries
        dataHierarchies.items[0].name = "Farm Sales";
        dataHierarchies.items[1].name = "Wholesale";
    });
});
```

## <a name="see-also"></a><span data-ttu-id="9418b-310">См. также</span><span class="sxs-lookup"><span data-stu-id="9418b-310">See also</span></span>

- [<span data-ttu-id="9418b-311">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="9418b-311">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9418b-312">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="9418b-312">Excel JavaScript API Reference</span></span>](/javascript/api/excel)

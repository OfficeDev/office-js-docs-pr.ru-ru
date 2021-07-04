---
title: Работа с pivotTables с помощью Excel API JavaScript
description: Используйте API Excel JavaScript для создания pivotTables и взаимодействия с их компонентами.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 8c8917f57b7546694e12380fc4369847be24ceac
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290742"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="bf707-103">Работа с pivotTables с помощью Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="bf707-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="bf707-104">PivotTables упрощают большие наборы данных.</span><span class="sxs-lookup"><span data-stu-id="bf707-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="bf707-105">Они позволяют быстро манипулировать сгруппными данными.</span><span class="sxs-lookup"><span data-stu-id="bf707-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="bf707-106">API Excel JavaScript позволяет надстройки создавать pivotTables и взаимодействовать с их компонентами.</span><span class="sxs-lookup"><span data-stu-id="bf707-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="bf707-107">В этой статье описывается, как pivotTables представлены API javaScript Office JavaScript и представлены примеры кода для ключевых сценариев.</span><span class="sxs-lookup"><span data-stu-id="bf707-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="bf707-108">Если вы не знакомы с функциями PivotTables, рассмотрите их в качестве конечного пользователя.</span><span class="sxs-lookup"><span data-stu-id="bf707-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="bf707-109">См. [в этой ссылке Создание pivotTable](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) для анализа данных таблицы для хорошей грунтовки на этих средствах.</span><span class="sxs-lookup"><span data-stu-id="bf707-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bf707-110">В настоящее время не поддерживаются pivotTables, созданные с помощью OLAP.</span><span class="sxs-lookup"><span data-stu-id="bf707-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="bf707-111">Также не поддерживается power Pivot.</span><span class="sxs-lookup"><span data-stu-id="bf707-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="bf707-112">Объектная модель</span><span class="sxs-lookup"><span data-stu-id="bf707-112">Object model</span></span>

<span data-ttu-id="bf707-113">[PivotTable](/javascript/api/excel/excel.pivottable) — это центральный объект для pivotTables в API javaScript Office JavaScript.</span><span class="sxs-lookup"><span data-stu-id="bf707-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="bf707-114">`Workbook.pivotTables` и `Worksheet.pivotTables` [являются pivotTableCollections,](/javascript/api/excel/excel.pivottablecollection) которые содержат [pivotTables](/javascript/api/excel/excel.pivottable) в книге и таблице, соответственно.</span><span class="sxs-lookup"><span data-stu-id="bf707-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="bf707-115">[PivotTable содержит](/javascript/api/excel/excel.pivottable) [pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) с несколькими [pivotHierarchies.](/javascript/api/excel/excel.pivothierarchy)</span><span class="sxs-lookup"><span data-stu-id="bf707-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="bf707-116">Эти [pivotHierarchies](/javascript/api/excel/excel.pivothierarchy) можно добавить в определенные коллекции иерархии, чтобы определить, как pivotTable определяет данные поворотов (как по объяснению в [следующем разделе).](#hierarchies)</span><span class="sxs-lookup"><span data-stu-id="bf707-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="bf707-117">[PivotHierarchy содержит](/javascript/api/excel/excel.pivothierarchy) [PivotFieldCollection,](/javascript/api/excel/excel.pivotfieldcollection) который имеет ровно один [PivotField](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="bf707-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="bf707-118">Если проект расширяется и включает pivotTables OLAP, это может измениться.</span><span class="sxs-lookup"><span data-stu-id="bf707-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="bf707-119">В [PivotField](/javascript/api/excel/excel.pivotfield) может применяться один или несколько [pivotFilters,](/javascript/api/excel/excel.pivotfilters) если [pivotHierarchy](/javascript/api/excel/excel.pivothierarchy) поля назначено в категорию иерархии.</span><span class="sxs-lookup"><span data-stu-id="bf707-119">A [PivotField](/javascript/api/excel/excel.pivotfield) can have one or more [PivotFilters](/javascript/api/excel/excel.pivotfilters) applied, as long as the field's [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) is assigned to a hierarchy category.</span></span>
- <span data-ttu-id="bf707-120">[PivotField содержит](/javascript/api/excel/excel.pivotfield) [pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) с несколькими [pivotItems.](/javascript/api/excel/excel.pivotitem)</span><span class="sxs-lookup"><span data-stu-id="bf707-120">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="bf707-121">[PivotTable](/javascript/api/excel/excel.pivottable) содержит [pivotLayout,](/javascript/api/excel/excel.pivotlayout) который определяет, где в таблице отображаются [PivotFields](/javascript/api/excel/excel.pivotfield) и [PivotItems.](/javascript/api/excel/excel.pivotitem)</span><span class="sxs-lookup"><span data-stu-id="bf707-121">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span> <span data-ttu-id="bf707-122">Макет также управляет некоторыми настройками отображения для PivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-122">The layout also controls some display settings for the PivotTable.</span></span>

<span data-ttu-id="bf707-123">Давайте рассмотрим, как эти отношения применяются к некоторым примерным данным.</span><span class="sxs-lookup"><span data-stu-id="bf707-123">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="bf707-124">В следующих данных описываются продажи фруктов из различных ферм.</span><span class="sxs-lookup"><span data-stu-id="bf707-124">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="bf707-125">Это будет пример всей этой статьи.</span><span class="sxs-lookup"><span data-stu-id="bf707-125">It will be the example throughout this article.</span></span>

![Коллекция продаж фруктов разных типов из разных ферм.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="bf707-127">Эти данные о продажах фермы будут использоваться для того, чтобы сделать PivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-127">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="bf707-128">Каждый столбец, например **Types,** является `PivotHierarchy` .</span><span class="sxs-lookup"><span data-stu-id="bf707-128">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="bf707-129">Иерархия **Типов** содержит поле **Типы.**</span><span class="sxs-lookup"><span data-stu-id="bf707-129">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="bf707-130">Поле **Types** содержит элементы **Apple,** **Kiwi,** **Lemon,** **Lime** и **Orange.**</span><span class="sxs-lookup"><span data-stu-id="bf707-130">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="bf707-131">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="bf707-131">Hierarchies</span></span>

<span data-ttu-id="bf707-132">PivotTables организованы на основе четырех категорий иерархии: [строка,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [столбец,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [данные](/javascript/api/excel/excel.datapivothierarchy)и [фильтр](/javascript/api/excel/excel.filterpivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="bf707-132">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="bf707-133">Данные фермы, показанные ранее, имеет пять иерархий: **фермы,** **тип,** **классификация,** ящики, проданные на **ферме,** и **ящики продаются** оптом.</span><span class="sxs-lookup"><span data-stu-id="bf707-133">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="bf707-134">Каждая иерархия может существовать только в одной из четырех категорий.</span><span class="sxs-lookup"><span data-stu-id="bf707-134">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="bf707-135">Если **тип** добавляется в иерархии столбцов, он также не может быть в строке, данных или иерархиях фильтрации.</span><span class="sxs-lookup"><span data-stu-id="bf707-135">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="bf707-136">Если **type** впоследствии добавляется в иерархии строк, он удаляется из иерархий столбцов.</span><span class="sxs-lookup"><span data-stu-id="bf707-136">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="bf707-137">Это поведение одинаково, если назначение иерархии Excel пользовательского интерфейса или Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="bf707-137">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="bf707-138">Иерархии строк и столбцов определяют группу данных.</span><span class="sxs-lookup"><span data-stu-id="bf707-138">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="bf707-139">Например, иерархия строк  Фермы сгруппит все наборы данных из одной фермы.</span><span class="sxs-lookup"><span data-stu-id="bf707-139">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="bf707-140">Выбор иерархии строк и столбцов определяет ориентацию pivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-140">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="bf707-141">Иерархии данных — это значения, которые будут агрегироваться на основе иерархий строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="bf707-141">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="bf707-142">PivotTable с иерархией строк  ферм и иерархией  данных оптовой продажи ящиков показывает общую сумму (по умолчанию) всех различных фруктов для каждой фермы.</span><span class="sxs-lookup"><span data-stu-id="bf707-142">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="bf707-143">Иерархии фильтров включают или исключают данные из поворота на основе значений этого фильтрованного типа.</span><span class="sxs-lookup"><span data-stu-id="bf707-143">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="bf707-144">Иерархия фильтров **классификации** с выбранным типом **Органический** показывает только данные для органических фруктов.</span><span class="sxs-lookup"><span data-stu-id="bf707-144">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="bf707-145">Вот еще раз данные фермы, а также pivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-145">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="bf707-146">PivotTable использует  Ферму  и Тип в качестве иерархий **строк,** Ящики, проданные в ферме, а ящики продаются  оптом в качестве иерархий данных (с функцией суммы агрегации по умолчанию) и Классификация как иерархия фильтров (с органическим выбранным).  </span><span class="sxs-lookup"><span data-stu-id="bf707-146">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![Выбор данных о продажах фруктов рядом с pivotTable с иерархиями строк, данных и фильтров.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="bf707-148">Этот pivotTable может быть создан с помощью API JavaScript или Excel пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="bf707-148">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="bf707-149">Оба варианта позволяют дальнейшие манипуляции с помощью надстройок.</span><span class="sxs-lookup"><span data-stu-id="bf707-149">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="bf707-150">Создание pivotTable</span><span class="sxs-lookup"><span data-stu-id="bf707-150">Create a PivotTable</span></span>

<span data-ttu-id="bf707-151">Для pivotTables необходимо имя, источник и пункт назначения.</span><span class="sxs-lookup"><span data-stu-id="bf707-151">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="bf707-152">Источником может быть адрес диапазона или имя таблицы (передается как `Range` , `string` или `Table` тип).</span><span class="sxs-lookup"><span data-stu-id="bf707-152">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="bf707-153">Пункт назначения — это адрес диапазона (дается как a `Range` `string` или).</span><span class="sxs-lookup"><span data-stu-id="bf707-153">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="bf707-154">В следующих примерах покажут различные методы создания pivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-154">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="bf707-155">Создание pivotTable с адресами диапазона</span><span class="sxs-lookup"><span data-stu-id="bf707-155">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="bf707-156">Создание pivotTable с объектами Range</span><span class="sxs-lookup"><span data-stu-id="bf707-156">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="bf707-157">Создание pivotTable на уровне книги</span><span class="sxs-lookup"><span data-stu-id="bf707-157">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="bf707-158">Использование существующего pivotTable</span><span class="sxs-lookup"><span data-stu-id="bf707-158">Use an existing PivotTable</span></span>

<span data-ttu-id="bf707-159">Созданные вручную pivotTables также доступны через коллекцию PivotTable книги или отдельных таблиц.</span><span class="sxs-lookup"><span data-stu-id="bf707-159">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="bf707-160">Следующий код получает pivotTable с именем **My Pivot из** книги.</span><span class="sxs-lookup"><span data-stu-id="bf707-160">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="bf707-161">Добавление строк и столбцов в pivotTable</span><span class="sxs-lookup"><span data-stu-id="bf707-161">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="bf707-162">Строки и столбцы совмещут данные вокруг значений этих полей.</span><span class="sxs-lookup"><span data-stu-id="bf707-162">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="bf707-163">Добавление **столбца Ферма** является поворотным для всех продаж каждой фермы.</span><span class="sxs-lookup"><span data-stu-id="bf707-163">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="bf707-164">Добавление строк **Тип** и **Классификация** еще больше разбивает данные, основанные на том, какие фрукты были проданы и были ли они органическими или нет.</span><span class="sxs-lookup"><span data-stu-id="bf707-164">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![PivotTable с столбцом Фермы и строками типа и классификации.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

<span data-ttu-id="bf707-166">Вы также можете иметь pivotTable только с строками или столбцами.</span><span class="sxs-lookup"><span data-stu-id="bf707-166">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="bf707-167">Добавление иерархий данных в PivotTable</span><span class="sxs-lookup"><span data-stu-id="bf707-167">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="bf707-168">Иерархии данных заполняют PivotTable информацией, которую необходимо объединить на основе строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="bf707-168">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="bf707-169">Добавление иерархий данных ящиков, проданных в **farm** и **Crates Sold Wholesale,** дает суммы этих цифр для каждой строки и столбца.</span><span class="sxs-lookup"><span data-stu-id="bf707-169">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="bf707-170">В примере **и Farm,** и **Type** — строки, а в качестве данных — объем продаж ящика.</span><span class="sxs-lookup"><span data-stu-id="bf707-170">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

![A PivotTable showing the total sales of different fruit based on the farm they came from.](../images/excel-pivots-data-hierarchy.png)

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="bf707-172">Макеты pivotTable и получение pivoted данных</span><span class="sxs-lookup"><span data-stu-id="bf707-172">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="bf707-173">[PivotLayout](/javascript/api/excel/excel.pivotlayout) определяет размещение иерархий и их данных.</span><span class="sxs-lookup"><span data-stu-id="bf707-173">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="bf707-174">Вы можете получить доступ к макету, чтобы определить диапазоны хранения данных.</span><span class="sxs-lookup"><span data-stu-id="bf707-174">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="bf707-175">На следующей схеме показано, какие вызовы функции макета соответствуют диапазонам pivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-175">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Схема, показывающая, какие разделы pivotTable возвращаются функциями диапазона макета.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="bf707-177">Получать данные из PivotTable</span><span class="sxs-lookup"><span data-stu-id="bf707-177">Get data from the PivotTable</span></span>

<span data-ttu-id="bf707-178">Макет определяет отображение pivotTable в таблице.</span><span class="sxs-lookup"><span data-stu-id="bf707-178">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="bf707-179">Это означает, `PivotLayout` что объект управляет диапазонами, используемыми для элементов PivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-179">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="bf707-180">Используйте диапазоны, предоставляемые макетом, чтобы получить данные, собранные и агрегированные сводной.</span><span class="sxs-lookup"><span data-stu-id="bf707-180">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="bf707-181">В частности, используйте `PivotLayout.getDataBodyRange` для доступа к данным, производимым pivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-181">In particular, use `PivotLayout.getDataBodyRange` to access the data produced by the PivotTable.</span></span>

<span data-ttu-id="bf707-182">В следующем коде показано, как получить последнюю строку данных PivotTable, проехав макет (общее  общее количество как суммы ящиков, проданных на ферме, так и суммы столбцов, проданных в начале примера).  </span><span class="sxs-lookup"><span data-stu-id="bf707-182">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="bf707-183">Затем эти значения суммируется для итогового итогового значения, отображаемого в ячейке **E30** (за пределами PivotTable).</span><span class="sxs-lookup"><span data-stu-id="bf707-183">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

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

### <a name="layout-types"></a><span data-ttu-id="bf707-184">Типы макетов</span><span class="sxs-lookup"><span data-stu-id="bf707-184">Layout types</span></span>

<span data-ttu-id="bf707-185">PivotTables имеют три стиля макета: Compact, Outline и Tabular.</span><span class="sxs-lookup"><span data-stu-id="bf707-185">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="bf707-186">В предыдущих примерах мы видели компактный стиль.</span><span class="sxs-lookup"><span data-stu-id="bf707-186">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="bf707-187">В следующих примерах используются, соответственно, схемы и табулярные стили.</span><span class="sxs-lookup"><span data-stu-id="bf707-187">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="bf707-188">Пример кода показывает цикл между различными макетами.</span><span class="sxs-lookup"><span data-stu-id="bf707-188">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="bf707-189">Макет схемы</span><span class="sxs-lookup"><span data-stu-id="bf707-189">Outline layout</span></span>

![PivotTable с помощью макета схемы.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="bf707-191">Макет табуляра</span><span class="sxs-lookup"><span data-stu-id="bf707-191">Tabular layout</span></span>

![PivotTable с помощью табулярного макета.](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a><span data-ttu-id="bf707-193">Пример кода коммутатора типа PivotLayout</span><span class="sxs-lookup"><span data-stu-id="bf707-193">PivotLayout type switch code sample</span></span>

```js
Excel.run(function (context) {
    // Change the PivotLayout.type to a new type.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    return context.sync().then(function () {
        // Cycle between the three layout types.
        if (pivotTable.layout.layoutType === "Compact") {
            pivotTable.layout.layoutType = "Outline";
        } else if (pivotTable.layout.layoutType === "Outline") {
            pivotTable.layout.layoutType = "Tabular";
        } else {
            pivotTable.layout.layoutType = "Compact";
        }
    
        return context.sync();
    });
});
```

### <a name="other-pivotlayout-functions"></a><span data-ttu-id="bf707-194">Другие функции PivotLayout</span><span class="sxs-lookup"><span data-stu-id="bf707-194">Other PivotLayout functions</span></span>

<span data-ttu-id="bf707-195">По умолчанию pivotTables корректирует размер строки и столбца по мере необходимости.</span><span class="sxs-lookup"><span data-stu-id="bf707-195">By default, PivotTables adjust row and column sizes as needed.</span></span> <span data-ttu-id="bf707-196">Это делается при обновлении PivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-196">This is done when the PivotTable is refreshed.</span></span> <span data-ttu-id="bf707-197">`PivotLayout.autoFormat` указывает такое поведение.</span><span class="sxs-lookup"><span data-stu-id="bf707-197">`PivotLayout.autoFormat` specifies that behavior.</span></span> <span data-ttu-id="bf707-198">Любые изменения размера строки или столбца, внесенные вашей надстройки, сохраняются, `autoFormat` когда `false` это .</span><span class="sxs-lookup"><span data-stu-id="bf707-198">Any row or column size changes made by your add-in persist when `autoFormat` is `false`.</span></span> <span data-ttu-id="bf707-199">Кроме того, параметры pivotTable по умолчанию сохраняют настраиваемый форматирование в PivotTable (например, изменения заливок и шрифтов).</span><span class="sxs-lookup"><span data-stu-id="bf707-199">Additionally, the default settings of a PivotTable keep any custom formatting in the PivotTable (such as fills and font changes).</span></span> <span data-ttu-id="bf707-200">Установите `PivotLayout.preserveFormatting` для `false` применения формата по умолчанию при обновлении.</span><span class="sxs-lookup"><span data-stu-id="bf707-200">Set `PivotLayout.preserveFormatting` to `false` to apply the default format when refreshed.</span></span>

<span data-ttu-id="bf707-201">Кроме того, элемент управления загонами и параметров общей строки, отображение пустых ячеек данных `PivotLayout` и [параметры текста alt.](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669)</span><span class="sxs-lookup"><span data-stu-id="bf707-201">A `PivotLayout` also controls header and total row settings, how empty data cells are displayed, and [alt text](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) options.</span></span> <span data-ttu-id="bf707-202">Ссылка [PivotLayout](/javascript/api/excel/excel.pivotlayout) содержит полный список этих функций.</span><span class="sxs-lookup"><span data-stu-id="bf707-202">The [PivotLayout](/javascript/api/excel/excel.pivotlayout) reference provides a complete list of these features.</span></span>

<span data-ttu-id="bf707-203">В следующем примере кода пустые ячейки данных отображают строку, форматировать диапазон тела до согласованного горизонтального выравнивания и гарантировать, что изменения форматирования остаются даже после обновления `"--"` PivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-203">The following code sample makes empty data cells display the string `"--"`, formats the body range to a consistent horizontal alignment, and ensures that the formatting changes remain even after the PivotTable is refreshed.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("Farm Sales");
    var pivotLayout = pivotTable.layout;

    // Set a default value for an empty cell in the PivotTable. This doesn't include cells left blank by the layout.
    pivotLayout.emptyCellText = "--";

    // Set the text alignment to match the rest of the PivotTable.
    pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;

    // Ensure empty cells are filled with a default value.
    pivotLayout.fillEmptyCells = true;

    // Ensure that the format settings persist, even after the PivotTable is refreshed and recalculated.
    pivotLayout.preserveFormatting = true;
    return context.sync();
});
```

## <a name="delete-a-pivottable"></a><span data-ttu-id="bf707-204">Удаление pivotTable</span><span class="sxs-lookup"><span data-stu-id="bf707-204">Delete a PivotTable</span></span>

<span data-ttu-id="bf707-205">PivotTables удаляются с помощью их имени.</span><span class="sxs-lookup"><span data-stu-id="bf707-205">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a><span data-ttu-id="bf707-206">Фильтр pivotTable</span><span class="sxs-lookup"><span data-stu-id="bf707-206">Filter a PivotTable</span></span>

<span data-ttu-id="bf707-207">Основной метод фильтрации данных pivotTable используется с помощью PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="bf707-207">The primary method for filtering PivotTable data is with PivotFilters.</span></span> <span data-ttu-id="bf707-208">Слайсеры предлагают альтернативный, менее гибкий метод фильтрации.</span><span class="sxs-lookup"><span data-stu-id="bf707-208">Slicers offer an alternate, less flexible filtering method.</span></span>

<span data-ttu-id="bf707-209">[PivotFilters](/javascript/api/excel/excel.pivotfilters) фильтрует данные на основе четырех категорий [](#hierarchies) иерархии PivotTable (фильтры, столбцы, строки и значения).</span><span class="sxs-lookup"><span data-stu-id="bf707-209">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filter data based on a PivotTable's four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="bf707-210">Существует четыре типа pivotFilters, которые позволяют фильтрацию на основе дат календаря, размыв строк, сравнение номеров и фильтрацию на основе настраиваемого ввода.</span><span class="sxs-lookup"><span data-stu-id="bf707-210">There are four types of PivotFilters, allowing calendar date-based filtering, string parsing, number comparison, and filtering based on a custom input.</span></span>

<span data-ttu-id="bf707-211">[Срезы](/javascript/api/excel/excel.slicer) можно применять как к таблицам PivotTables, так и к Excel таблицам.</span><span class="sxs-lookup"><span data-stu-id="bf707-211">[Slicers](/javascript/api/excel/excel.slicer) can be applied to both PivotTables and regular Excel tables.</span></span> <span data-ttu-id="bf707-212">При применении к pivotTable срезы функционируют как [PivotManualFilter](#pivotmanualfilter) и позволяют фильтрацию на основе настраиваемого ввода.</span><span class="sxs-lookup"><span data-stu-id="bf707-212">When applied to a PivotTable, slicers function like a [PivotManualFilter](#pivotmanualfilter) and allow filtering based on a custom input.</span></span> <span data-ttu-id="bf707-213">В отличие от PivotFilters, срезеры имеют [компонент Excel пользовательского интерфейса.](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)</span><span class="sxs-lookup"><span data-stu-id="bf707-213">Unlike PivotFilters, slicers have an [Excel UI component](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span></span> <span data-ttu-id="bf707-214">С помощью `Slicer` класса вы создаете этот компонент пользовательского интерфейса, управляете фильтрацией и контролируете его внешний вид.</span><span class="sxs-lookup"><span data-stu-id="bf707-214">With the `Slicer` class, you create this UI component, manage filtering, and control its visual appearance.</span></span>

### <a name="filter-with-pivotfilters"></a><span data-ttu-id="bf707-215">Фильтр с помощью pivotFilters</span><span class="sxs-lookup"><span data-stu-id="bf707-215">Filter with PivotFilters</span></span>

<span data-ttu-id="bf707-216">[PivotFilters](/javascript/api/excel/excel.pivotfilters) позволяют фильтровать данные pivotTable на основе четырех категорий [иерархии](#hierarchies) (фильтры, столбцы, строки и значения).</span><span class="sxs-lookup"><span data-stu-id="bf707-216">[PivotFilters](/javascript/api/excel/excel.pivotfilters) allow you to filter PivotTable data based on the four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="bf707-217">В объектной модели PivotTable применяются к `PivotFilters` [PivotField,](/javascript/api/excel/excel.pivotfield)и каждому из них может быть назначено одно `PivotField` или несколько `PivotFilters` объектов.</span><span class="sxs-lookup"><span data-stu-id="bf707-217">In the PivotTable object model, `PivotFilters` are applied to a [PivotField](/javascript/api/excel/excel.pivotfield), and each `PivotField` can have one or more assigned `PivotFilters`.</span></span> <span data-ttu-id="bf707-218">Чтобы применить PivotFilters к PivotField, необходимо приурочеть соответствующую [pivotHierarchy](/javascript/api/excel/excel.pivothierarchy) поля к категории иерархии.</span><span class="sxs-lookup"><span data-stu-id="bf707-218">To apply PivotFilters to a PivotField, the field's corresponding [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) must be assigned to a hierarchy category.</span></span>

#### <a name="types-of-pivotfilters"></a><span data-ttu-id="bf707-219">Типы pivotFilters</span><span class="sxs-lookup"><span data-stu-id="bf707-219">Types of PivotFilters</span></span>

| <span data-ttu-id="bf707-220">Тип фильтра</span><span class="sxs-lookup"><span data-stu-id="bf707-220">Filter type</span></span> | <span data-ttu-id="bf707-221">Назначение фильтра</span><span class="sxs-lookup"><span data-stu-id="bf707-221">Filter purpose</span></span> | <span data-ttu-id="bf707-222">Справочные материалы по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="bf707-222">Excel JavaScript API reference</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="bf707-223">DateFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-223">DateFilter</span></span> | <span data-ttu-id="bf707-224">Фильтрация на основе даты календаря.</span><span class="sxs-lookup"><span data-stu-id="bf707-224">Calendar date-based filtering.</span></span> | [<span data-ttu-id="bf707-225">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-225">PivotDateFilter</span></span>](/javascript/api/excel/excel.pivotdatefilter) |
| <span data-ttu-id="bf707-226">LabelFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-226">LabelFilter</span></span> | <span data-ttu-id="bf707-227">Фильтрация сравнения текста.</span><span class="sxs-lookup"><span data-stu-id="bf707-227">Text comparison filtering.</span></span> | [<span data-ttu-id="bf707-228">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-228">PivotLabelFilter</span></span>](/javascript/api/excel/excel.pivotlabelfilter) |
| <span data-ttu-id="bf707-229">ManualFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-229">ManualFilter</span></span> | <span data-ttu-id="bf707-230">Настраиваемая фильтрация входных данных.</span><span class="sxs-lookup"><span data-stu-id="bf707-230">Custom input filtering.</span></span> | [<span data-ttu-id="bf707-231">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-231">PivotManualFilter</span></span>](/javascript/api/excel/excel.pivotmanualfilter) |
| <span data-ttu-id="bf707-232">ValueFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-232">ValueFilter</span></span> | <span data-ttu-id="bf707-233">Фильтрация сравнения номеров.</span><span class="sxs-lookup"><span data-stu-id="bf707-233">Number comparison filtering.</span></span> | [<span data-ttu-id="bf707-234">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-234">PivotValueFilter</span></span>](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a><span data-ttu-id="bf707-235">Создание pivotFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-235">Create a PivotFilter</span></span>

<span data-ttu-id="bf707-236">Чтобы фильтровать pivotTable данные с помощью (например, a), применить фильтр `Pivot*Filter` к `PivotDateFilter` [PivotField](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="bf707-236">To filter PivotTable data with a `Pivot*Filter` (such as a `PivotDateFilter`), apply the filter to a [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="bf707-237">В следующих четырех примерах кода покажите, как использовать каждый из четырех типов PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="bf707-237">The following four code samples show how to use each of the four types of PivotFilters.</span></span>

##### <a name="pivotdatefilter"></a><span data-ttu-id="bf707-238">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-238">PivotDateFilter</span></span>

<span data-ttu-id="bf707-239">Первый пример кода применяет [pivotDateFilter к PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) с обновленным pivotField, скрывая все данные до **2020-08-01**. </span><span class="sxs-lookup"><span data-stu-id="bf707-239">The first code sample applies a [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) to the **Date Updated** PivotField, hiding any data prior to **2020-08-01**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bf707-240">A не может применяться к PivotField, если только pivotHierarchy этого поля не назначена в `Pivot*Filter` категорию иерархии.</span><span class="sxs-lookup"><span data-stu-id="bf707-240">A `Pivot*Filter` can't be applied to a PivotField unless that field's PivotHierarchy is assigned to a hierarchy category.</span></span> <span data-ttu-id="bf707-241">В следующем примере кода необходимо добавить его в категорию PivotTable, прежде чем его можно будет использовать `dateHierarchy` `rowHierarchies` для фильтрации.</span><span class="sxs-lookup"><span data-stu-id="bf707-241">In the following code sample, the `dateHierarchy` must be added to the PivotTable's `rowHierarchies` category before it can be used for filtering.</span></span>

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
> <span data-ttu-id="bf707-242">В следующих трех фрагментах кода отображаются только выдержки, определенные фильтром, а не полные `Excel.run` вызовы.</span><span class="sxs-lookup"><span data-stu-id="bf707-242">The following three code snippets only display filter-specific excerpts, instead of full `Excel.run` calls.</span></span>

##### <a name="pivotlabelfilter"></a><span data-ttu-id="bf707-243">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-243">PivotLabelFilter</span></span>

<span data-ttu-id="bf707-244">Второй фрагмент кода демонстрирует, как применить [pivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) к **Типу** PivotField, используя свойство для исключения меток, которые начинаются с буквы `LabelFilterCondition.beginsWith` **L**.</span><span class="sxs-lookup"><span data-stu-id="bf707-244">The second code snippet demonstrates how to apply a [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) to the **Type** PivotField, using the `LabelFilterCondition.beginsWith` property to exclude labels that start with the letter **L**.</span></span>

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

##### <a name="pivotmanualfilter"></a><span data-ttu-id="bf707-245">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-245">PivotManualFilter</span></span>

<span data-ttu-id="bf707-246">Третий фрагмент кода применяет ручной фильтр [с PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) в поле **Классификация,** отфильтровыв данные, не включающие классификацию **Organic.**</span><span class="sxs-lookup"><span data-stu-id="bf707-246">The third code snippet applies a manual filter with [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) to the the **Classification** field, filtering out data that doesn't include the classification **Organic**.</span></span>

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a><span data-ttu-id="bf707-247">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="bf707-247">PivotValueFilter</span></span>

<span data-ttu-id="bf707-248">Чтобы сравнить числа, используйте фильтр значения [с PivotValueFilter,](/javascript/api/excel/excel.pivotvaluefilter)как показано в заключительном фрагменте кода.</span><span class="sxs-lookup"><span data-stu-id="bf707-248">To compare numbers, use a value filter with [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), as shown in the final code snippet.</span></span> <span data-ttu-id="bf707-249">Эти данные сравнивают в pivotField фермы с данными в Оптовом pivotField, в том числе только фермами, сумма проданных ящиков превышает `PivotValueFilter` **значение 500**.  </span><span class="sxs-lookup"><span data-stu-id="bf707-249">The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Wholesale** PivotField, including only farms whose sum of crates sold exceeds the value **500**.</span></span>

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

#### <a name="remove-pivotfilters"></a><span data-ttu-id="bf707-250">Удаление pivotFilters</span><span class="sxs-lookup"><span data-stu-id="bf707-250">Remove PivotFilters</span></span>

<span data-ttu-id="bf707-251">Чтобы удалить все pivotFilters, применяем метод к каждому `clearAllFilters` PivotField, как показано в следующем примере кода.</span><span class="sxs-lookup"><span data-stu-id="bf707-251">To remove all PivotFilters, apply the `clearAllFilters` method to each PivotField, as shown in the following code sample.</span></span>

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

### <a name="filter-with-slicers"></a><span data-ttu-id="bf707-252">Фильтр с помощью срезов</span><span class="sxs-lookup"><span data-stu-id="bf707-252">Filter with slicers</span></span>

<span data-ttu-id="bf707-253">[Срезы](/javascript/api/excel/excel.slicer) позволяют фильтровать данные из Excel pivotTable или таблицы.</span><span class="sxs-lookup"><span data-stu-id="bf707-253">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="bf707-254">Срезер использует значения из указанного столбца или PivotField для фильтрации соответствующих строк.</span><span class="sxs-lookup"><span data-stu-id="bf707-254">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="bf707-255">Эти значения хранятся в [качестве объектов SlicerItem](/javascript/api/excel/excel.sliceritem) в `Slicer` .</span><span class="sxs-lookup"><span data-stu-id="bf707-255">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="bf707-256">Ваша надстройка может регулировать эти фильтры, как и пользователи (Excel[пользовательского интерфейса).](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)</span><span class="sxs-lookup"><span data-stu-id="bf707-256">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="bf707-257">Срез находится на вершине таблицы в слое рисования, как показано на следующем скриншоте.</span><span class="sxs-lookup"><span data-stu-id="bf707-257">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![Фильтрующий срез данных на pivotTable.](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="bf707-259">Методы, описанные в этом разделе, посвящены использованию срезов, подключенных к PivotTables.</span><span class="sxs-lookup"><span data-stu-id="bf707-259">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="bf707-260">Те же методы применяются и к использованию срезов, подключенных к таблицам.</span><span class="sxs-lookup"><span data-stu-id="bf707-260">The same techniques also apply to using slicers connected to tables.</span></span>

#### <a name="create-a-slicer"></a><span data-ttu-id="bf707-261">Создание среза</span><span class="sxs-lookup"><span data-stu-id="bf707-261">Create a slicer</span></span>

<span data-ttu-id="bf707-262">С помощью метода или метода можно создать срез в книге или на `Workbook.slicers.add` `Worksheet.slicers.add` таблице.</span><span class="sxs-lookup"><span data-stu-id="bf707-262">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="bf707-263">Это добавляет срез в [slicerCollection](/javascript/api/excel/excel.slicercollection) указанного или `Workbook` `Worksheet` объекта.</span><span class="sxs-lookup"><span data-stu-id="bf707-263">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="bf707-264">Метод `SlicerCollection.add` имеет три параметра:</span><span class="sxs-lookup"><span data-stu-id="bf707-264">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="bf707-265">`slicerSource`Источник данных, на котором основан новый срез.</span><span class="sxs-lookup"><span data-stu-id="bf707-265">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="bf707-266">Это может быть строка или строка, представляющая `PivotTable` `Table` имя или ID или `PivotTable` `Table` .</span><span class="sxs-lookup"><span data-stu-id="bf707-266">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="bf707-267">`sourceField`: Поле в источнике данных, с помощью которого необходимо фильтровать.</span><span class="sxs-lookup"><span data-stu-id="bf707-267">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="bf707-268">Это может быть строка или строка, представляющая `PivotField` `TableColumn` имя или ID или `PivotField` `TableColumn` .</span><span class="sxs-lookup"><span data-stu-id="bf707-268">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="bf707-269">`slicerDestination`. Таблица, на которой будет создан новый срез.</span><span class="sxs-lookup"><span data-stu-id="bf707-269">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="bf707-270">Это может быть объект или имя или `Worksheet` ИД `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="bf707-270">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="bf707-271">Этот параметр не является необходимым при `SlicerCollection` доступе через `Worksheet.slicers` .</span><span class="sxs-lookup"><span data-stu-id="bf707-271">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="bf707-272">В этом случае в качестве назначения используется таблица коллекции.</span><span class="sxs-lookup"><span data-stu-id="bf707-272">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="bf707-273">В следующем примере кода в таблицу **Pivot** добавляется новый срез.</span><span class="sxs-lookup"><span data-stu-id="bf707-273">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="bf707-274">Источником среза является pivotTable продаж фермы и фильтры с помощью **данных Type.** </span><span class="sxs-lookup"><span data-stu-id="bf707-274">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="bf707-275">Срезер также называется **Fruit Slicer для** будущей ссылки.</span><span class="sxs-lookup"><span data-stu-id="bf707-275">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

#### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="bf707-276">Фильтрация элементов с помощью среза</span><span class="sxs-lookup"><span data-stu-id="bf707-276">Filter items with a slicer</span></span>

<span data-ttu-id="bf707-277">Срезник фильтрует pivotTable с элементами из `sourceField` .</span><span class="sxs-lookup"><span data-stu-id="bf707-277">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="bf707-278">Метод `Slicer.selectItems` задает элементы, которые остаются в срезе.</span><span class="sxs-lookup"><span data-stu-id="bf707-278">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="bf707-279">Эти элементы передаются методу в качестве `string[]` , представляющего ключи элементов.</span><span class="sxs-lookup"><span data-stu-id="bf707-279">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="bf707-280">Все строки, содержащие эти элементы, остаются в агрегации PivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-280">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="bf707-281">Последующие `selectItems` вызовы для набора списка к клавишам, указанным в этих вызовах.</span><span class="sxs-lookup"><span data-stu-id="bf707-281">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="bf707-282">Если передается элемент, который не находится в источнике `Slicer.selectItems` данных, будет `InvalidArgument` выброшена ошибка.</span><span class="sxs-lookup"><span data-stu-id="bf707-282">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="bf707-283">Содержимое можно проверить с помощью `Slicer.slicerItems` свойства [SlicerItemCollection.](/javascript/api/excel/excel.sliceritemcollection)</span><span class="sxs-lookup"><span data-stu-id="bf707-283">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="bf707-284">В следующем примере кода показаны три пункта, выбранные для среза: **лимон,** **лайм** и **оранжевый**.</span><span class="sxs-lookup"><span data-stu-id="bf707-284">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="bf707-285">Чтобы удалить все фильтры из среза, используйте `Slicer.clearFilters` метод, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="bf707-285">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a><span data-ttu-id="bf707-286">Стиль и формат среза</span><span class="sxs-lookup"><span data-stu-id="bf707-286">Style and format a slicer</span></span>

<span data-ttu-id="bf707-287">Надстройка может настраивать параметры отображения среза с помощью `Slicer` свойств.</span><span class="sxs-lookup"><span data-stu-id="bf707-287">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="bf707-288">Следующий пример кода задает стиль **SlicerStyleLight6,** задает текст в верхней части среза для **типов** фруктов, помещает срез в положение **(395, 15)** на уровне рисования и задает размер среза до **135x150** пикселей.</span><span class="sxs-lookup"><span data-stu-id="bf707-288">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

#### <a name="delete-a-slicer"></a><span data-ttu-id="bf707-289">Удаление среза</span><span class="sxs-lookup"><span data-stu-id="bf707-289">Delete a slicer</span></span>

<span data-ttu-id="bf707-290">Чтобы удалить срез, позвоните по `Slicer.delete` методу.</span><span class="sxs-lookup"><span data-stu-id="bf707-290">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="bf707-291">Следующий пример кода удаляет первый срез из текущего таблицы.</span><span class="sxs-lookup"><span data-stu-id="bf707-291">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="bf707-292">Функция агрегирования изменений</span><span class="sxs-lookup"><span data-stu-id="bf707-292">Change aggregation function</span></span>

<span data-ttu-id="bf707-293">Иерархии данных объединяют свои значения.</span><span class="sxs-lookup"><span data-stu-id="bf707-293">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="bf707-294">Для наборов данных номеров это сумма по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="bf707-294">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="bf707-295">Свойство определяет такое поведение на основе типа `summarizeBy` [AggregationFunction.](/javascript/api/excel/excel.aggregationfunction)</span><span class="sxs-lookup"><span data-stu-id="bf707-295">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="bf707-296">Поддерживаемые в настоящее время типы функций агрегации: `Sum` , , , , , , , `Count` и `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` (по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="bf707-296">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="bf707-297">В следующих примерах кода агрегация изменяется на средние значения данных.</span><span class="sxs-lookup"><span data-stu-id="bf707-297">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="bf707-298">Изменение вычислений с помощью ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="bf707-298">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="bf707-299">Сводки по умолчанию объединяют данные иерархий строки и столбцов независимо друг от друга.</span><span class="sxs-lookup"><span data-stu-id="bf707-299">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="bf707-300">[ShowAsRule](/javascript/api/excel/excel.showasrule) изменяет иерархию данных на значения вывода на основе других элементов в PivotTable.</span><span class="sxs-lookup"><span data-stu-id="bf707-300">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="bf707-301">Объект `ShowAsRule` имеет три свойства:</span><span class="sxs-lookup"><span data-stu-id="bf707-301">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="bf707-302">`calculation`: Тип относительного вычисления, применяемого к иерархии данных (по `none` умолчанию).</span><span class="sxs-lookup"><span data-stu-id="bf707-302">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="bf707-303">`baseField`. [PivotField](/javascript/api/excel/excel.pivotfield) в иерархии, содержащей базовые данные перед применив расчет.</span><span class="sxs-lookup"><span data-stu-id="bf707-303">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="bf707-304">Поскольку Excel pivotTables имеют сопоставление иерархии в поле один к одному, вы будете использовать одно и то же имя для доступа как к иерархии, так и к полю.</span><span class="sxs-lookup"><span data-stu-id="bf707-304">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="bf707-305">`baseItem`: Отдельный [pivotItem](/javascript/api/excel/excel.pivotitem) сравнивается со значениями базовых полей, основанными на типе вычисления.</span><span class="sxs-lookup"><span data-stu-id="bf707-305">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="bf707-306">Не все вычисления требуют этого поля.</span><span class="sxs-lookup"><span data-stu-id="bf707-306">Not all calculations require this field.</span></span>

<span data-ttu-id="bf707-307">В следующем примере вычисление суммы ящиков, проданных в иерархии данных **фермы,** определяется в процентах от общего числа столбцов.</span><span class="sxs-lookup"><span data-stu-id="bf707-307">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="bf707-308">Мы по-прежнему хотим, чтобы детализация распространила на уровень типа плода, поэтому мы будем использовать иерархию строк **Type** и ее поле.</span><span class="sxs-lookup"><span data-stu-id="bf707-308">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="bf707-309">В примере также **имеется иерархия** фермы в качестве первой строки, поэтому общие записи фермы отображают процент, который каждая ферма отвечает за производство.</span><span class="sxs-lookup"><span data-stu-id="bf707-309">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![A PivotTable showing the percentages of fruit sales relative to the grand total for both individual farms and individual fruit types within each farm.](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="bf707-311">В предыдущем примере за установите вычисление столбца по отношению к полю отдельной иерархии строк.</span><span class="sxs-lookup"><span data-stu-id="bf707-311">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="bf707-312">Если расчет относится к отдельному элементу, используйте `baseItem` свойство.</span><span class="sxs-lookup"><span data-stu-id="bf707-312">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="bf707-313">В следующем примере показан `differenceFrom` расчет.</span><span class="sxs-lookup"><span data-stu-id="bf707-313">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="bf707-314">Он отображает разницу записей иерархии данных о продажах фермы по сравнению с записями **a Farms.**</span><span class="sxs-lookup"><span data-stu-id="bf707-314">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="bf707-315">Это ферма, поэтому мы видим различия между другими фермами, а также разбивки для каждого типа как `baseField` фрукты **(Тип** также иерархия строки в этом примере).</span><span class="sxs-lookup"><span data-stu-id="bf707-315">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![PivotTable с указанием различий в продажах фруктов между "Фермами" и другими.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="bf707-319">Изменение имен иерархии</span><span class="sxs-lookup"><span data-stu-id="bf707-319">Change hierarchy names</span></span>

<span data-ttu-id="bf707-320">Области иерархии можно изменить.</span><span class="sxs-lookup"><span data-stu-id="bf707-320">Hierarchy fields are editable.</span></span> <span data-ttu-id="bf707-321">В следующем коде показано, как изменить отображаемые имена двух иерархий данных.</span><span class="sxs-lookup"><span data-stu-id="bf707-321">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="bf707-322">См. также</span><span class="sxs-lookup"><span data-stu-id="bf707-322">See also</span></span>

- [<span data-ttu-id="bf707-323">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="bf707-323">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="bf707-324">Excel Ссылка на API JavaScript</span><span class="sxs-lookup"><span data-stu-id="bf707-324">Excel JavaScript API Reference</span></span>](/javascript/api/excel)

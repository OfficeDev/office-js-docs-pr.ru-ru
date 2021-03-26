---
title: Основные концепции программирования с помощью API JavaScript для Excel
description: Создание надстроек для Excel с помощью API JavaScript для Excel.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: dde7dc66e0746fc4d9cf91ed3df824fab05c109d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292604"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="e5991-103">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e5991-103">Fundamental programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="e5991-104">В этой статье описано, как создавать надстройки для Excel 2016 или более поздней версии с помощью [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md).</span><span class="sxs-lookup"><span data-stu-id="e5991-104">This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="e5991-105">В статье изложены основные принципы, которые являются фундаментальными при использовании этого API, а также имеются рекомендации по выполнению определенных задач, например чтению данных из большого диапазона или записи данных в него, изменения всех ячеек в диапазоне и много другого.</span><span class="sxs-lookup"><span data-stu-id="e5991-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e5991-106">Сведения об асинхронном типе интерфейсов API Excel и принципах их работы с книгой см. в статье [Использование модели API, зависящей от приложения](../develop/application-specific-api-model.md).</span><span class="sxs-lookup"><span data-stu-id="e5991-106">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Excel APIs and how they work with the workbook.</span></span>  

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="e5991-107">Интерфейсы API Office.js для Excel</span><span class="sxs-lookup"><span data-stu-id="e5991-107">Office.js APIs for Excel</span></span>

<span data-ttu-id="e5991-108">Надстройка Excel взаимодействует с объектами в Excel с помощью API JavaScript для Office, включающего две объектных модели JavaScript:</span><span class="sxs-lookup"><span data-stu-id="e5991-108">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="e5991-109">**API JavaScript для Excel**. Появившийся в Office 2016 [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам.</span><span class="sxs-lookup"><span data-stu-id="e5991-109">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="e5991-110">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="e5991-110">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="e5991-111">Скорее всего, вы будете разрабатывать большую часть функций надстроек для Excel 2016 или более поздней версии с помощью API JavaScript для Excel, но вам также потребуются объекты из общего API.</span><span class="sxs-lookup"><span data-stu-id="e5991-111">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="e5991-112">Например:</span><span class="sxs-lookup"><span data-stu-id="e5991-112">For example:</span></span>

* <span data-ttu-id="e5991-p103">[Context](/javascript/api/office/office.context). Объект `Context` представляет среду выполнения надстройки и предоставляет доступ к ключевым объектам API. Он состоит из данных конфигурации книги, например `contentLanguage` и `officeTheme`, а также предоставляет сведения о среде выполнения надстройки, например `host` и `platform`. Кроме того, он предоставляет метод `requirements.isSetSupported()`, с помощью которого можно проверить, поддерживается ли указанный набор обязательных элементов приложением Excel, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="e5991-p103">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="e5991-116">[Document](/javascript/api/office/office.document). Объект `Document` предоставляет метод `getFileAsync()`, позволяющий скачать файл Excel, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="e5991-116">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="e5991-117">На рисунке ниже показано, когда можно использовать API JavaScript для Excel или общие API.</span><span class="sxs-lookup"><span data-stu-id="e5991-117">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Изображение различий между API JS для Excel и общими API](../images/excel-js-api-common-api.png)

## <a name="object-model"></a><span data-ttu-id="e5991-119">Объектная модель</span><span class="sxs-lookup"><span data-stu-id="e5991-119">Object model</span></span>

<span data-ttu-id="e5991-120">Чтобы понять API-интерфейсы Excel, вы должны понимать, как компоненты рабочей книги связаны друг с другом.</span><span class="sxs-lookup"><span data-stu-id="e5991-120">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

* <span data-ttu-id="e5991-121">**Рабочая книга** содержит одну или несколько **рабочих листов**.</span><span class="sxs-lookup"><span data-stu-id="e5991-121">A **Workbook** contains one or more **Worksheets**.</span></span>
* <span data-ttu-id="e5991-122">**Рабочий лист** предоставляет доступ к ячейкам через объекты **Range**.</span><span class="sxs-lookup"><span data-stu-id="e5991-122">A **Worksheet** gives access to cells through **Range** objects.</span></span>
* <span data-ttu-id="e5991-123">**Range** представляет группу смежных клеток.</span><span class="sxs-lookup"><span data-stu-id="e5991-123">A **Range** represents a group of contiguous cells.</span></span>
* <span data-ttu-id="e5991-124">**Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур** и других объектов визуализации данных или организации.</span><span class="sxs-lookup"><span data-stu-id="e5991-124">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
* <span data-ttu-id="e5991-125">**Рабочий лист** содержит коллекции тех объектов данных, которые присутствуют на отдельном листе.</span><span class="sxs-lookup"><span data-stu-id="e5991-125">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
* <span data-ttu-id="e5991-126">**Рабочие книги** содержат коллекции некоторых из этих объектов данных (таких как **таблицы**) для всей **рабочей книги**.</span><span class="sxs-lookup"><span data-stu-id="e5991-126">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="e5991-127">Диапазоны</span><span class="sxs-lookup"><span data-stu-id="e5991-127">Ranges</span></span>

<span data-ttu-id="e5991-128">Диапазон - это группа непрерывных ячеек в рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="e5991-128">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="e5991-129">В надстройках обычно используется нотация в стиле A1 (например, **B3** для отдельной ячейки в столбце **B** и строке **3** или **C2:F4** для ячеек из столбцов с **C** по **F** и строк с **2** по **4**) для определения диапазонов.</span><span class="sxs-lookup"><span data-stu-id="e5991-129">Add-ins typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="e5991-130">Диапазоны имеют три основных свойства: `values`, `formulas`, и `format`.</span><span class="sxs-lookup"><span data-stu-id="e5991-130">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="e5991-131">Эти свойства получают или устанавливают значения ячеек, формулы для оценки и визуальное форматирование ячеек.</span><span class="sxs-lookup"><span data-stu-id="e5991-131">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="e5991-132">Образец диапазона</span><span class="sxs-lookup"><span data-stu-id="e5991-132">Range sample</span></span>

<span data-ttu-id="e5991-133">В следующем примере показано, как создавать записи продаж.</span><span class="sxs-lookup"><span data-stu-id="e5991-133">The following sample shows how to create sales records.</span></span> <span data-ttu-id="e5991-134">Эта функция использует объекты `Range` для установки значений, формул и форматов.</span><span class="sxs-lookup"><span data-stu-id="e5991-134">This function uses `Range` objects to set the values, formulas, and formats.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

<span data-ttu-id="e5991-135">В этом примере создаются следующие данные в текущем листе:</span><span class="sxs-lookup"><span data-stu-id="e5991-135">This sample creates the following data in the current worksheet:</span></span>

![Запись о продажах, показывающая строки значений, столбец формулы и отформатированные заголовки.](../images/excel-overview-range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="e5991-137">Диаграммы, таблицы и другие объекты данных</span><span class="sxs-lookup"><span data-stu-id="e5991-137">Charts, tables, and other data objects</span></span>

<span data-ttu-id="e5991-138">API JavaScript для Excel могут создавать и управлять структурами данных и визуализациями в Excel.</span><span class="sxs-lookup"><span data-stu-id="e5991-138">The Excel JavaScript APIs can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="e5991-139">Таблицы и диаграммы являются двумя наиболее часто используемыми объектами, но API поддерживают сводные таблицы, фигуры, изображения и многое другое.</span><span class="sxs-lookup"><span data-stu-id="e5991-139">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="e5991-140">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="e5991-140">Creating a table</span></span>

<span data-ttu-id="e5991-141">Создавайте таблицы, используя заполненные данными диапазоны.</span><span class="sxs-lookup"><span data-stu-id="e5991-141">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="e5991-142">Элементы управления форматированием и таблицами (например, фильтры) автоматически применяются к диапазону.</span><span class="sxs-lookup"><span data-stu-id="e5991-142">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="e5991-143">В следующем примере создается таблица с использованием диапазонов из предыдущего примера.</span><span class="sxs-lookup"><span data-stu-id="e5991-143">The following sample creates a table using the ranges from the previous sample.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

<span data-ttu-id="e5991-144">Использование этого примера кода на листе с предыдущими данными создает следующую таблицу:</span><span class="sxs-lookup"><span data-stu-id="e5991-144">Using this sample code on the worksheet with the previous data creates the following table:</span></span>

![Таблица сделана из предыдущего рекорда продаж.](../images/excel-overview-table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="e5991-146">Создание диаграммы</span><span class="sxs-lookup"><span data-stu-id="e5991-146">Creating a chart</span></span>

<span data-ttu-id="e5991-147">Создайте диаграммы для визуализации данных в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="e5991-147">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="e5991-148">API поддерживают десятки разновидностей диаграмм, каждая из которых может быть настроена в соответствии с вашими потребностями.</span><span class="sxs-lookup"><span data-stu-id="e5991-148">The APIs support dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="e5991-149">В следующем примере создается простая гистограмма для трех элементов, которая размещается на 100 пикселей ниже верхней части листа.</span><span class="sxs-lookup"><span data-stu-id="e5991-149">The following sample creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

<span data-ttu-id="e5991-150">Выполнение этого примера на листе с предыдущей таблицей создает следующую диаграмму:</span><span class="sxs-lookup"><span data-stu-id="e5991-150">Running this sample on the worksheet with the previous table creates the following chart:</span></span>

![Гистограмма, показывающая количества трех элементов из предыдущей записи о продажах.](../images/excel-overview-chart-sample.png)

## <a name="run-options"></a><span data-ttu-id="e5991-152">Параметры выполнения</span><span class="sxs-lookup"><span data-stu-id="e5991-152">Run options</span></span>

<span data-ttu-id="e5991-153">`Excel.run` есть перегрузка, получающая объект [RunOptions](/javascript/api/excel/excel.runoptions).</span><span class="sxs-lookup"><span data-stu-id="e5991-153">`Excel.run` has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="e5991-154">Он содержит набор свойств, влияющих на поведение платформы при выполнении функции.</span><span class="sxs-lookup"><span data-stu-id="e5991-154">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="e5991-155">Ниже перечислены поддерживаемые в настоящее время свойства.</span><span class="sxs-lookup"><span data-stu-id="e5991-155">The following property is currently supported:</span></span>

* <span data-ttu-id="e5991-156">`delayForCellEdit`: определяет, откладывает ли Excel пакетный запрос до выхода пользователя из режима правки ячейки.</span><span class="sxs-lookup"><span data-stu-id="e5991-156">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="e5991-157">Если присвоено значение **true**, пакетный запрос откладывается и запускается, когда пользователь выходит из режима правки ячейки.</span><span class="sxs-lookup"><span data-stu-id="e5991-157">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="e5991-158">Если присвоено значение **false**, происходит автоматический сбой пакетного запроса, если пользователь находится в режиме правки ячейки (приводит к ошибке обращения к пользователю).</span><span class="sxs-lookup"><span data-stu-id="e5991-158">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="e5991-159">Поведение по умолчанию при отсутствии заданного свойства `delayForCellEdit` аналогично поведению при значении **false**.</span><span class="sxs-lookup"><span data-stu-id="e5991-159">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="null-or-blank-property-values"></a><span data-ttu-id="e5991-160">Значения null или пустые значения свойств</span><span class="sxs-lookup"><span data-stu-id="e5991-160">null or blank property values</span></span>

<span data-ttu-id="e5991-161">Значения `null` и пустые строки имеют специальные применения в API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="e5991-161">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="e5991-162">Они используются для представления пустых ячеек, отсутствия форматирования или значений по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e5991-162">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="e5991-163">В этом разделе описано использование значения `null` и пустой строки при получении и настройке свойств.</span><span class="sxs-lookup"><span data-stu-id="e5991-163">This section details the use of `null` and empty string when getting and setting properties.</span></span>

### <a name="null-input-in-2-d-array"></a><span data-ttu-id="e5991-164">Входное значение null в двумерном массиве</span><span class="sxs-lookup"><span data-stu-id="e5991-164">null input in 2-D Array</span></span>

<span data-ttu-id="e5991-p113">В Excel диапазон представлен двумерным массивом, в котором первое измерение — это строки, а второе — столбцы. Чтобы задать значения, формат чисел или формулу только для определенных ячеек в диапазоне, укажите значения, формат чисел или формулу для этих ячеек в двумерном массиве, а для всех остальных ячеек в этом массиве укажите значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e5991-p113">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="e5991-p114">Например, чтобы изменить формат чисел только для одной ячейки в диапазоне и сохранить существующий формат чисел для всех остальных ячеек в диапазоне, укажите новый формат чисел для ячейки, которую необходимо изменить, а для всех остальных ячеек укажите значение `null`. Во фрагменте кода ниже показано, как задать новый формат чисел для четвертой ячейки в диапазоне, при этом формат чисел для первых трех ячеек в диапазоне останется неизменным.</span><span class="sxs-lookup"><span data-stu-id="e5991-p114">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a><span data-ttu-id="e5991-169">Входное значение null для свойства</span><span class="sxs-lookup"><span data-stu-id="e5991-169">null input for a property</span></span>

<span data-ttu-id="e5991-p115">`null` не является допустимым входным значением для одного свойства. Например, указанный ниже фрагмент кода не является допустимым, так как свойство `values` диапазона не должно иметь значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e5991-p115">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null;
```

<span data-ttu-id="e5991-172">Аналогично, указанный ниже фрагмент кода не является допустимым, так как `null` — недопустимое значение для свойства `color`.</span><span class="sxs-lookup"><span data-stu-id="e5991-172">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a><span data-ttu-id="e5991-173">Значения свойств null в ответе</span><span class="sxs-lookup"><span data-stu-id="e5991-173">null property values in the response</span></span>

<span data-ttu-id="e5991-p116">Если в указанном диапазоне имеются другие значения, свойства форматирования, например `size` и `color` будут содержать значения `null` в ответе. Например, если вы получаете диапазон и загружаете его свойство `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="e5991-p116">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="e5991-176">Если у всех ячеек в диапазоне один и тот же цвет шрифта, свойство `range.format.font.color` указывает этот цвет.</span><span class="sxs-lookup"><span data-stu-id="e5991-176">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="e5991-177">Если в диапазоне используется несколько цветов шрифтов, свойство `range.format.font.color` имеет значение `null`.</span><span class="sxs-lookup"><span data-stu-id="e5991-177">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

### <a name="blank-input-for-a-property"></a><span data-ttu-id="e5991-178">Пустое входное значение для свойства</span><span class="sxs-lookup"><span data-stu-id="e5991-178">Blank input for a property</span></span>

<span data-ttu-id="e5991-p117">Когда вы указываете пустое значение для свойства (то есть две кавычки подряд без других знаков между `''`), это будет интерпретировано как инструкция по очистке или сбросу свойства. Например:</span><span class="sxs-lookup"><span data-stu-id="e5991-p117">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="e5991-181">Если вы укажете пустое значение для свойства `values` диапазона, содержимое диапазона будет очищено.</span><span class="sxs-lookup"><span data-stu-id="e5991-181">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="e5991-182">Если вы укажете пустое значение для свойства `numberFormat`, формат чисел будет "сброшен" до формата `General`.</span><span class="sxs-lookup"><span data-stu-id="e5991-182">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="e5991-183">Если вы укажете пустое значение для свойств `formula` и `formulaLocale`, значения формул будут очищены.</span><span class="sxs-lookup"><span data-stu-id="e5991-183">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="e5991-184">Значения пустых свойств в ответе</span><span class="sxs-lookup"><span data-stu-id="e5991-184">Blank property values in the response</span></span>

<span data-ttu-id="e5991-p118">Для операций чтения пустое значение свойства в ответе (то есть две кавычки подряд без других знаков между `''`) указывает, что ячейка не содержит данных или значения. В первом примере ниже первая и последняя ячейки в диапазоне не содержат данных. Во втором примере две первые ячейки в диапазоне не содержат формул.</span><span class="sxs-lookup"><span data-stu-id="e5991-p118">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="requirement-sets"></a><span data-ttu-id="e5991-188">Наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="e5991-188">Requirement sets</span></span>

<span data-ttu-id="e5991-189">Наборы требований — это именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="e5991-189">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="e5991-190">Надстройка Office может выполнить проверку в среде выполнения или использовать указанные в манифесте наборы обязательных элементов, чтобы определить, поддерживает ли приложение Office необходимые надстройке API.</span><span class="sxs-lookup"><span data-stu-id="e5991-190">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span> <span data-ttu-id="e5991-191">Сведения о том, какие именно наборы обязательных элементов доступны на каждой поддерживаемой платформе, см. в статье [Наборы обязательных элементов API JavaScript для Excel](../reference/requirement-sets/excel-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e5991-191">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="e5991-192">Проверка поддержки наборов обязательных элементов в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="e5991-192">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="e5991-193">В следующем примере кода показано, как определить, поддерживает ли приложение Office, в котором запускается надстройка, указанный набор обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="e5991-193">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="e5991-194">Определение поддержки наборов обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="e5991-194">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="e5991-195">С помощью [элемента Requirements](../reference/manifest/requirements.md) в манифесте надстройки можно указать минимальные наборы обязательных элементов и/или методы API, необходимые надстройке для активации.</span><span class="sxs-lookup"><span data-stu-id="e5991-195">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="e5991-196">Если платформа или приложение Office не поддерживает наборы обязательных элементов или методы API, указанные в элементе `Requirements` манифеста, надстройка не будет работать в этом приложении или на этой платформе, а также не будет отображаться в списке надстроек в разделе **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="e5991-196">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="e5991-197">В следующем примере кода показан элемент `Requirements` в манифесте надстройки, где указано, что надстройка должна загружаться во всех клиентских приложениях Office, поддерживающих набор обязательных элементов ExcelApi версии 1.3 или выше.</span><span class="sxs-lookup"><span data-stu-id="e5991-197">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="e5991-198">Чтобы надстройка была доступна на всех платформах приложения Office, например Excel в Интернете, Excel для Windows и для iPad, рекомендуем проверять поддержку обязательных элементов в среде выполнения, а не определять поддержку набора обязательных элементов в манифесте.</span><span class="sxs-lookup"><span data-stu-id="e5991-198">To make your add-in available on all platforms of an Office application, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="e5991-199">Наборы обязательных элементов общего API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="e5991-199">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="e5991-200">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](../reference/requirement-sets/office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e5991-200">For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

## <a name="handle-errors"></a><span data-ttu-id="e5991-201">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="e5991-201">Handle errors</span></span>

<span data-ttu-id="e5991-202">При возникновении ошибки в интерфейсе API он возвращает объект `error`, содержащий код и сообщение.</span><span class="sxs-lookup"><span data-stu-id="e5991-202">When an API error occurs, the API returns an `error` object that contains a code and a message.</span></span> <span data-ttu-id="e5991-203">Подробные сведения об обработке ошибок, включая список ошибок API, см. в статье [Обработка ошибок](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="e5991-203">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e5991-204">См. также</span><span class="sxs-lookup"><span data-stu-id="e5991-204">See also</span></span>

* [<span data-ttu-id="e5991-205">Создание первой надстройки Excel</span><span class="sxs-lookup"><span data-stu-id="e5991-205">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
* [<span data-ttu-id="e5991-206">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="e5991-206">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="e5991-207">Оптимизация производительности API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e5991-207">Excel JavaScript API performance optimization</span></span>](../excel/performance.md)
* [<span data-ttu-id="e5991-208">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e5991-208">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

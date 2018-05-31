---
title: Работа с диаграммами с использованием API JavaScript для Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: c0f45892cb937a565a6855390344855f75e7473e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437446"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a><span data-ttu-id="133e8-102">Работа с диаграммами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="133e8-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="133e8-103">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для диаграмм с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="133e8-103">This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API.</span></span> <span data-ttu-id="133e8-104">Полный список свойств и методов, поддерживаемых объектами **Chart** и **ChartCollection**, см. в статьях [Объект Chart (API JavaScript для Excel)](https://dev.office.com/reference/add-ins/excel/chart) и [Объект ChartCollection (API JavaScript для Excel)](https://dev.office.com/reference/add-ins/excel/chartcollection).</span><span class="sxs-lookup"><span data-stu-id="133e8-104">For the complete list of properties and methods that the **Chart** and **ChartCollection** objects support, see [Chart Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/chart) and [Chart Collection Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/chartcollection).</span></span>

## <a name="create-a-chart"></a><span data-ttu-id="133e8-105">Создание диаграммы</span><span class="sxs-lookup"><span data-stu-id="133e8-105">Create a chart</span></span>

<span data-ttu-id="133e8-106">В примере кода ниже показано, как создать диаграмму на листе **Sample** (Пример).</span><span class="sxs-lookup"><span data-stu-id="133e8-106">The following code sample creates a chart in the worksheet named **Sample**.</span></span> <span data-ttu-id="133e8-107">Диаграмма представляет собой **график**, построенный на основе данных из диапазона **A1:B13**.</span><span class="sxs-lookup"><span data-stu-id="133e8-107">The chart is a **Line** chart that is based upon data in the range **A1:B13**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var dataRange = sheet.getRange("A1:B13");
    var chart = sheet.charts.add("Line", dataRange, "auto");

    chart.title.text = "Sales Data";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="133e8-108">**Новый график**</span><span class="sxs-lookup"><span data-stu-id="133e8-108">**New line chart**</span></span>

![Новый график в Excel](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a><span data-ttu-id="133e8-110">Добавление ряда данных в диаграмму</span><span class="sxs-lookup"><span data-stu-id="133e8-110">Add a data series to a chart</span></span>

<span data-ttu-id="133e8-111">В примере кода ниже показано, как добавить ряд данных в первую диаграмму на листе.</span><span class="sxs-lookup"><span data-stu-id="133e8-111">The following code sample adds a data series to the first chart in the worksheet.</span></span> <span data-ttu-id="133e8-112">Новый ряд данных соответствует столбцу **2016** и основан на данных из диапазона **D2:D5**.</span><span class="sxs-lookup"><span data-stu-id="133e8-112">The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.</span></span>

> [!NOTE]
> <span data-ttu-id="133e8-113">В этом примере используются API, которые в данный момент доступны только в виде общедоступной ознакомительной версии (бета-версии).</span><span class="sxs-lookup"><span data-stu-id="133e8-113">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="133e8-114">Чтобы запустить этот образец, вы должны использовать бета-библиотеку CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="133e8-114">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var chart = sheet.charts.getItemAt(0);
    var dataRange = sheet.getRange("D2:D5");

    var newSeries = chart.series.add("2016");
    newSeries.setValues(dataRange);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="133e8-115">**Диаграмма перед добавлением ряда данных 2016**</span><span class="sxs-lookup"><span data-stu-id="133e8-115">**Chart before the 2016 data series is added**</span></span>

![Диаграмма в Excel перед добавлением ряда данных 2016](../images/excel-charts-data-series-before.png)

<span data-ttu-id="133e8-117">**Диаграмма после добавления ряда данных 2016**</span><span class="sxs-lookup"><span data-stu-id="133e8-117">**Chart after the 2016 data series is added**</span></span>

![Диаграмма в Excel после добавления ряда данных 2016](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a><span data-ttu-id="133e8-119">Задание названия диаграммы</span><span class="sxs-lookup"><span data-stu-id="133e8-119">Set chart title</span></span>

<span data-ttu-id="133e8-120">В примере ниже показано, как задать название **Sales Data by Year** (Данные продаж по годам) для первой диаграммы на листе.</span><span class="sxs-lookup"><span data-stu-id="133e8-120">The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="133e8-121">**Диаграмма после задания заголовка**</span><span class="sxs-lookup"><span data-stu-id="133e8-121">**Chart after title is set**</span></span>

![Диаграмма с заголовком в Excel](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a><span data-ttu-id="133e8-123">Задание свойств оси диаграммы</span><span class="sxs-lookup"><span data-stu-id="133e8-123">Set properties of an axis in a chart</span></span>

<span data-ttu-id="133e8-124">Диаграммы, в которых используется [декартова система координат](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), например гистограммы, линейчатые и точечные диаграммы, содержат ось категорий и ось значений.</span><span class="sxs-lookup"><span data-stu-id="133e8-124">Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis.</span></span> <span data-ttu-id="133e8-125">В примерах ниже показано, как задать название и отобразить единицу измерения по оси для диаграммы.</span><span class="sxs-lookup"><span data-stu-id="133e8-125">These examples show how to set the title and display unit of an axis in a chart.</span></span>

### <a name="set-axis-title"></a><span data-ttu-id="133e8-126">Задание названия оси</span><span class="sxs-lookup"><span data-stu-id="133e8-126">Set axis title</span></span>

<span data-ttu-id="133e8-127">В примере кода ниже показано, как задать название **Product** (Продукт) для оси категорий первой диаграммы на листе.</span><span class="sxs-lookup"><span data-stu-id="133e8-127">The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="133e8-128">**Диаграмма после задания названия оси категорий**</span><span class="sxs-lookup"><span data-stu-id="133e8-128">**Chart after title of category axis is set**</span></span>

![Диаграмма с названием оси в Excel](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a><span data-ttu-id="133e8-130">Задание отображаемой единицы измерения оси</span><span class="sxs-lookup"><span data-stu-id="133e8-130">Set axis display unit</span></span>

<span data-ttu-id="133e8-131">В примере ниже показано, как задать отображаемую единицу измерения **Hundreds** (Сотни) для оси значений первой диаграммы на листе.</span><span class="sxs-lookup"><span data-stu-id="133e8-131">The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.</span></span>

> [!NOTE]
> <span data-ttu-id="133e8-132">В этом примере используются API, которые в данный момент доступны только в виде общедоступной ознакомительной версии (бета-версии).</span><span class="sxs-lookup"><span data-stu-id="133e8-132">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="133e8-133">Чтобы запустить этот образец, вы должны использовать бета-библиотеку CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="133e8-133">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="133e8-134">**Диаграмма после задания единицы измерения оси значений**</span><span class="sxs-lookup"><span data-stu-id="133e8-134">**Chart after display unit of value axis is set**</span></span>

![Диаграмма с отображаемой единицей измерения оси значений в Excel](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a><span data-ttu-id="133e8-136">Настройка видимости линий сетки на диаграмме</span><span class="sxs-lookup"><span data-stu-id="133e8-136">Set visibility of gridlines in a chart</span></span>

<span data-ttu-id="133e8-137">В примере ниже показано, как скрыть основные линии сетки для оси значений первой диаграммы на листе.</span><span class="sxs-lookup"><span data-stu-id="133e8-137">The following code sample hides the major gridlines for the value axis of the first chart in the worksheet.</span></span> <span data-ttu-id="133e8-138">Вы можете отобразить основные линии сетки для оси значений диаграммы, задав для свойства `chart.axes.valueAxis.majorGridlines.visible` значение **true**.</span><span class="sxs-lookup"><span data-stu-id="133e8-138">You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to **true**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="133e8-139">**Диаграмма со скрытыми линиями сетки**</span><span class="sxs-lookup"><span data-stu-id="133e8-139">**Chart with gridlines hidden**</span></span>

![Диаграмма со скрытыми линиями сетки в Excel](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a><span data-ttu-id="133e8-141">Линии трендов диаграммы</span><span class="sxs-lookup"><span data-stu-id="133e8-141">Chart trendlines</span></span>

### <a name="add-a-trendline"></a><span data-ttu-id="133e8-142">Добавление линии тренда</span><span class="sxs-lookup"><span data-stu-id="133e8-142">Add a trendline</span></span>

<span data-ttu-id="133e8-p108">В примере кода ниже показано, как добавить линию тренда "скользящее среднее" в первый ряд первой диаграммы на листе **Sample** (Пример). Линия тренда отображает "скользящее среднее" за 5 периодов.</span><span class="sxs-lookup"><span data-stu-id="133e8-p108">The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.</span></span>

> [!NOTE]
> <span data-ttu-id="133e8-145">В этом примере используются API, которые в данный момент доступны только в виде общедоступной ознакомительной версии (бета-версии).</span><span class="sxs-lookup"><span data-stu-id="133e8-145">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="133e8-146">Чтобы запустить этот образец, вы должны использовать бета-библиотеку CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="133e8-146">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="133e8-147">**Диаграмма с линией тренда "скользящее среднее"**</span><span class="sxs-lookup"><span data-stu-id="133e8-147">**Chart with moving average trendline**</span></span>

![Диаграмма с линией тренда "скользящее среднее" в Excel](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a><span data-ttu-id="133e8-149">Изменение линии тренда</span><span class="sxs-lookup"><span data-stu-id="133e8-149">Update a trendline</span></span>

<span data-ttu-id="133e8-150">В примере кода ниже показано, как задать для линии тренда тип **Linear** (Линейная) для первого ряда первой диаграммы на листе **Sample** (Пример).</span><span class="sxs-lookup"><span data-stu-id="133e8-150">The following code sample sets the trendline to type **Linear** for the first series in the first chart in the worksheet named **Sample**.</span></span>

> [!NOTE]
> <span data-ttu-id="133e8-151">В этом примере используются API, которые в данный момент доступны только в виде общедоступной ознакомительной версии (бета-версии).</span><span class="sxs-lookup"><span data-stu-id="133e8-151">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="133e8-152">Чтобы запустить этот образец, вы должны использовать бета-библиотеку CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="133e8-152">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    var series = seriesCollection.getItemAt(0);
    series.trendlines.getItem(0).type = "Linear";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="133e8-153">**Диаграмма с линейной линией тренда**</span><span class="sxs-lookup"><span data-stu-id="133e8-153">**Chart with linear trendline**</span></span>

![Диаграмма с линейной линией тренда в Excel](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a><span data-ttu-id="133e8-155">См. также</span><span class="sxs-lookup"><span data-stu-id="133e8-155">See also</span></span>

- [<span data-ttu-id="133e8-156">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="133e8-156">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="133e8-157">Объект Chart (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="133e8-157">Chart Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/chart) 
- [<span data-ttu-id="133e8-158">Объект ChartCollection (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="133e8-158">Chart Collection Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/chartcollection)
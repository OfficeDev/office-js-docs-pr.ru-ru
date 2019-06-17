---
title: Обзор API JavaScript для Excel
description: ''
ms.date: 06/10/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aa9574a93252c0011b211c39e37cc013beb64432
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910149"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="a6985-102">Обзор API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a6985-102">Excel JavaScript API overview</span></span>

<span data-ttu-id="a6985-103">Вы можете использовать API JavaScript для Excel, чтобы создавать надстройки для Excel 2016 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="a6985-103">You can use the Excel JavaScript API to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="a6985-104">Ниже перечислены объекты Excel высокого уровня, доступные в API.</span><span class="sxs-lookup"><span data-stu-id="a6985-104">The following list shows the high-level Excel objects that are available in the API.</span></span> <span data-ttu-id="a6985-105">Каждая ссылка на страницу объекта содержит описание свойств, событий и методов, доступных для объекта.</span><span class="sxs-lookup"><span data-stu-id="a6985-105">Each object page link contains a description of the properties, events, and methods that are available on the object.</span></span> <span data-ttu-id="a6985-106">Чтобы узнать больше, перейдите по соответствующим ссылкам в меню.</span><span class="sxs-lookup"><span data-stu-id="a6985-106">Explore the links from the menu to learn more.</span></span>

<span data-ttu-id="a6985-107">Для удобства ниже перечислены некоторые из основных объектов Excel.</span><span class="sxs-lookup"><span data-stu-id="a6985-107">Some of the core Excel objects are listed below for convenience:</span></span>

- <span data-ttu-id="a6985-108">[Workbook](/javascript/api/excel/excel.workbook) — объект верхнего уровня, содержащий связанные объекты книг, такие как листы, таблицы, диапазоны и т. д. Его также можно использовать для вывода списка связанных ссылок.</span><span class="sxs-lookup"><span data-stu-id="a6985-108">[Workbook](/javascript/api/excel/excel.workbook): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.</span></span>

- <span data-ttu-id="a6985-109">[Worksheet](/javascript/api/excel/excel.worksheet). Представляет лист в книге.</span><span class="sxs-lookup"><span data-stu-id="a6985-109">[Worksheet](/javascript/api/excel/excel.worksheet): Represents a worksheet in a workbook.</span></span>
  - <span data-ttu-id="a6985-110">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): коллекция объектов **Worksheet** в книге.</span><span class="sxs-lookup"><span data-stu-id="a6985-110">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): A collection of the **Worksheet** objects in a workbook.</span></span>
  - <span data-ttu-id="a6985-111">[Worksheet Protection](/javascript/api/excel/excel.worksheetprotection): защита объекта **Worksheet**.</span><span class="sxs-lookup"><span data-stu-id="a6985-111">[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): Represents the protection of a **Worksheet** object.</span></span>

- <span data-ttu-id="a6985-112">[Range](/javascript/api/excel/excel.range): ячейка, строка, столбец или группа ячеек, содержащая один или несколько смежных блоков ячеек.</span><span class="sxs-lookup"><span data-stu-id="a6985-112">[Range](/javascript/api/excel/excel.range): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.</span></span>
  - <span data-ttu-id="a6985-113">[ConditionalFormat](/javascript/api/excel/excel.conditionalformat): объект, определяющий правило и формат, которые применяются к диапазону при соблюдении условия правила.</span><span class="sxs-lookup"><span data-stu-id="a6985-113">[ConditionalFormat](/javascript/api/excel/excel.conditionalformat): An object defining a rule and a format applied to the range when the rule's condition is met.</span></span>
  - <span data-ttu-id="a6985-114">[DataValidation](/javascript/api/excel/excel.datavalidation): объект, ограничивающий вводимые пользователем данные диапазоном, в основе которого лежит ряд условий.</span><span class="sxs-lookup"><span data-stu-id="a6985-114">[DataValidation](/javascript/api/excel/excel.datavalidation): An object that restricts user input to a range based on a variety of criteria.</span></span>
  - <span data-ttu-id="a6985-115">[RangeSort](/javascript/api/excel/excel.rangesort): объект, управляющий операциями сортировки для диапазона.</span><span class="sxs-lookup"><span data-stu-id="a6985-115">[RangeSort](/javascript/api/excel/excel.rangesort): Represents a object that manages sorting operations on a range.</span></span>

- <span data-ttu-id="a6985-116">[Table](/javascript/api/excel/excel.table): коллекция упорядоченных ячеек для упрощения управления данными.</span><span class="sxs-lookup"><span data-stu-id="a6985-116">[Table](/javascript/api/excel/excel.table): Represents a collection of organized cells designed to make management of the data easy.</span></span>
  - <span data-ttu-id="a6985-117">[TableCollection](/javascript/api/excel/excel.tablecollection). Коллекция таблиц в книге или на листе.</span><span class="sxs-lookup"><span data-stu-id="a6985-117">[TableCollection](/javascript/api/excel/excel.tablecollection): A collection of tables in a workbook or worksheet.</span></span>
  - <span data-ttu-id="a6985-118">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection). Коллекция всех столбцов в таблице.</span><span class="sxs-lookup"><span data-stu-id="a6985-118">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): A collection of all the columns in a table.</span></span>
  - <span data-ttu-id="a6985-119">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection): коллекция всех строк в таблице.</span><span class="sxs-lookup"><span data-stu-id="a6985-119">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection): A collection of all the rows in a table.</span></span>
  - <span data-ttu-id="a6985-120">[TableSort](/javascript/api/excel/excel.tablesort): объект, управляющий операциями сортировки в таблице.</span><span class="sxs-lookup"><span data-stu-id="a6985-120">[TableSort](/javascript/api/excel/excel.tablesort): Represents an object that manages sorting operations on a table.</span></span>

- <span data-ttu-id="a6985-121">[Chart](/javascript/api/excel/excel.chart): объект диаграммы на листе, который является визуальным представлением базовых данных.</span><span class="sxs-lookup"><span data-stu-id="a6985-121">[Chart](/javascript/api/excel/excel.chart): Represents a chart object in a worksheet, which is a visual representation of underlying data.</span></span>
  - <span data-ttu-id="a6985-122">[ChartCollection](/javascript/api/excel/excel.chartcollection): коллекция диаграмм на листе.</span><span class="sxs-lookup"><span data-stu-id="a6985-122">[ChartCollection](/javascript/api/excel/excel.chartcollection): A collection of charts in a worksheet.</span></span>

- <span data-ttu-id="a6985-123">[PivotTable](/javascript/api/excel/excel.pivottable): сводная таблица Excel, которая является иерархической группировкой и представлением данных.</span><span class="sxs-lookup"><span data-stu-id="a6985-123">[PivotTable](/javascript/api/excel/excel.pivottable): Represents an Excel PivotTable, which is a hierarchical grouping and presentation of data.</span></span>
  - <span data-ttu-id="a6985-124">[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): коллекция сводных таблиц на листе.</span><span class="sxs-lookup"><span data-stu-id="a6985-124">[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): A collection of PivotTables in a worksheet.</span></span>

- <span data-ttu-id="a6985-125">[Filter](/javascript/api/excel/excel.filter): объект, управляющий фильтрацией столбца таблицы.</span><span class="sxs-lookup"><span data-stu-id="a6985-125">[Filter](/javascript/api/excel/excel.filter): Represents an object that manages the filtering of a table's column.</span></span>

- <span data-ttu-id="a6985-126">[NamedItem](/javascript/api/excel/excel.nameditem): определенное имя для диапазона ячеек или значения.</span><span class="sxs-lookup"><span data-stu-id="a6985-126">[NamedItem](/javascript/api/excel/excel.nameditem): Represents a defined name for a range of cells or a value.</span></span>
  - <span data-ttu-id="a6985-127">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection). Коллекция объектов **NamedItem** в книге.</span><span class="sxs-lookup"><span data-stu-id="a6985-127">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): A collection of the **NamedItem** objects in a workbook.</span></span>

- <span data-ttu-id="a6985-128">[Binding](/javascript/api/excel/excel.binding): абстрактный класс, представляющий привязку к разделу книги.</span><span class="sxs-lookup"><span data-stu-id="a6985-128">[Binding](/javascript/api/excel/excel.binding): An abstract class that represents a binding to a section of the workbook.</span></span>
  - <span data-ttu-id="a6985-129">[BindingCollection](/javascript/api/excel/excel.bindingcollection): коллекция объектов **Binding** в книге.</span><span class="sxs-lookup"><span data-stu-id="a6985-129">[BindingCollection](/javascript/api/excel/excel.bindingcollection): A collection of the **Binding** objects in a workbook.</span></span>

## <a name="excel-javascript-api-requirement-sets"></a><span data-ttu-id="a6985-130">Наборы обязательных элементов API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a6985-130">Excel JavaScript API requirement sets</span></span>

<span data-ttu-id="a6985-131">Наборы обязательных элементов — именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="a6985-131">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="a6985-132">Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API.</span><span class="sxs-lookup"><span data-stu-id="a6985-132">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="a6985-133">Дополнительны сведения о наборах обязательных элементов API JavaScript для Excel см. в статье [Наборы требований API JavaScript для Excel](../requirement-sets/excel-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="a6985-133">For detailed information about Excel JavaScript API requirement sets, see the [Excel JavaScript API requirement sets](../requirement-sets/excel-api-requirement-sets.md) article.</span></span>

## <a name="excel-javascript-api-reference"></a><span data-ttu-id="a6985-134">Справочные материалы по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a6985-134">Excel JavaScript API reference</span></span>

<span data-ttu-id="a6985-135">Дополнительные сведения об API JavaScript для Excel см. в [справочной документации по API JavaScript для Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="a6985-135">For detailed information about the Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="see-also"></a><span data-ttu-id="a6985-136">См. также</span><span class="sxs-lookup"><span data-stu-id="a6985-136">See also</span></span>

- [<span data-ttu-id="a6985-137">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="a6985-137">Excel add-ins overview</span></span>](/office/dev/add-ins/excel/excel-add-ins-overview)
- [<span data-ttu-id="a6985-138">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a6985-138">Office Add-ins platform overview</span></span>](/office/dev/add-ins/overview/office-add-ins)
- [<span data-ttu-id="a6985-139">Примеры надстроек Excel на сайте GitHub</span><span class="sxs-lookup"><span data-stu-id="a6985-139">Excel add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
- [<span data-ttu-id="a6985-140">Открытые спецификации API</span><span class="sxs-lookup"><span data-stu-id="a6985-140">API open specifications</span></span>](../openspec/openspec.md)

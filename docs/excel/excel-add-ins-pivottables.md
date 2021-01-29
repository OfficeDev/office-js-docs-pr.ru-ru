---
title: Работа со сводными таблицами с помощью API JavaScript для Excel
description: Используйте API JavaScript для Excel, чтобы создавать с помощью севоварок и взаимодействовать с их компонентами.
ms.date: 01/26/2021
localization_priority: Normal
ms.openlocfilehash: 9832322d40bbeb247685ff2498bdce42975c0377
ms.sourcegitcommit: 3123b9819c5225ee45a5312f64be79e46cbd0e3c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/29/2021
ms.locfileid: "50043913"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Работа со сводными таблицами с помощью API JavaScript для Excel

СvotTables упрощают более крупные наборы данных. Они позволяют быстро манипулировать сгруппными данными. С помощью API JavaScript для Excel надстройка может создавать с помощью список и взаимодействовать с их компонентами. В этой статье описывается, как API JavaScript для Office представлена с помощью список, и представлены примеры кода для ключевых сценариев.

Если вы не знакомы с функциями с помощью список, рассмотрите возможность их изучения в качестве конечного пользователя.
Дополнительные сведения об этих [средствах](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) см. в подметке "Создание с помощью с помощью совверной таблицы для анализа данных на таблицах".

> [!IMPORTANT]
> В настоящее время с помощью OLAP не поддерживаются с помощью стеблей. Power Pivot также не поддерживается.

## <a name="object-model"></a>Объектная модель

[PivotTable](/javascript/api/excel/excel.pivottable) — это центральный объект для список в API JavaScript для Office.

- `Workbook.pivotTables` и `Worksheet.pivotTables` являются [pivotTableCollections,](/javascript/api/excel/excel.pivottablecollection) которые содержат [pivotTables](/javascript/api/excel/excel.pivottable) в книге и на электронных таблицах соответственно.
- [PivotTable](/javascript/api/excel/excel.pivottable) содержит [pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) с несколькими [pivotHierarchies](/javascript/api/excel/excel.pivothierarchy).
- Эти [pivotHierarchies](/javascript/api/excel/excel.pivothierarchy) можно добавить в определенные коллекции иерархии, чтобы определить, как данные pivotTable pivots (как поясняется в [следующем разделе).](#hierarchies)
- [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) содержит [pivotFieldCollection,](/javascript/api/excel/excel.pivotfieldcollection) который содержит только одно [pivotField](/javascript/api/excel/excel.pivotfield). Если проект расширяется и включает в себя севобли OLAP, это может измениться.
- К [pivotField](/javascript/api/excel/excel.pivotfield) может применяться один или несколько [pivotFilters,](/javascript/api/excel/excel.pivotfilters) если [pivotHierarchy](/javascript/api/excel/excel.pivothierarchy) поля назначена категории иерархии. 
- [PivotField](/javascript/api/excel/excel.pivotfield) содержит [pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) с несколькими [pivotItems](/javascript/api/excel/excel.pivotitem).
- [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.

Рассмотрим, как эти связи применяются к некоторым примерам данных. В следующих данных описаны продажи деревьев из различных ферм. Это будет пример в этой статье.

![Коллекция продаж разных видов в разных фермах.](../images/excel-pivots-raw-data.png)

Эти данные продаж фермы будет использоваться для список. Каждый столбец, например **"Типы",** имеет тип `PivotHierarchy` . **Иерархия** типов содержит поле **"Типы".** Поле **"Типы"** содержит элементы **Apple,** **Kiwi,** **Юли** **и** **Orange.**

### <a name="hierarchies"></a>Hierarchies

СvotTables организованы на основе четырех категорий иерархии: [строка,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [столбец,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [данные](/javascript/api/excel/excel.datapivothierarchy)и [фильтр](/javascript/api/excel/excel.filterpivothierarchy).

Данные фермы, показанные ранее, имеет пять иерархий: **фермы,** **тип,** **классификация,** ящики, проданные в ферме, и **crates Sold Farms.** Каждая иерархия может существовать только в одной из четырех категорий. Если **тип** добавляется в иерархии столбцов, он также не может быть в иерархиях строк, данных или фильтров. Если **тип** впоследствии добавляется в иерархии строк, он удаляется из иерархий столбцов. Это поведение одинаково для назначения иерархии с помощью пользовательского интерфейса Excel или API JavaScript для Excel.

Иерархии строк и столбцов определяют группировку данных. Например, иерархия строк  ферм объединяет все наборы данных из одной фермы. Выбор между иерархией строк и столбцов определяет ориентацию pivotTable.

Иерархии данных — это значения, которые необходимо агрегировать на основе иерархий строк и столбцов. С помощью список с иерархией  строк в фермах и иерархией данных "Crates **Sold Farms"** (По умолчанию) для каждой фермы показана общая сумма (по умолчанию) всех разных ферм.

Иерархии фильтров включают или исключают данные из pivot на основе значений этого отфильтрованного типа. Иерархия фильтров **классификации** с выбранным типом **"Органическая"** показывает данные только для органично выбранных видов.

Вот данные фермы еще раз вместе со списной. С помощью иерархии строк  "Ферма" и "Тип" в качестве иерархии используется "Ферма" и "Тип", "Crates **Sold at Farm"** и  **"Crates SoldМайл"** в качестве иерархий данных (с функцией суммирования по умолчанию) и **"Классификация** как иерархия фильтров" (с выбранным "органичным"). 

![Выбор данных о продажах вех рядом со совивей с иерархиями строк, данных и фильтров.](../images/excel-pivot-table-and-data.png)

Эта списоная таблица может быть сгенерирована с помощью API JavaScript или пользовательского интерфейса Excel. Оба варианта позволяют дальнейшее манипулирование с помощью надстройки.

## <a name="create-a-pivottable"></a>Создание севоtTable

СvotTables need a name, source, and destination. Источником может быть адрес диапазона или имя таблицы (переданное в виде `Range` , `string` или `Table` типа). Назначением является адрес диапазона (заданный как a `Range` или `string` ).
В следующих примерах демонстрируются различные методики создания с помощью с помощью список.

### <a name="create-a-pivottable-with-range-addresses"></a>Создание с помощью севоttable с адресами диапазона

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Создание с помощью с помощью объектов Range

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

### <a name="create-a-pivottable-at-the-workbook-level"></a>Создание севоtTable на уровне книги

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Использование существующей pivotTable

С помощью коллекции pivotTables, созданной вручную, можно также использовать коллекцию pivotTable книги или отдельных таблиц. Следующий код получает из книги севоtTable с именем **My Pivot.**

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Добавление строк и столбцов в списоную

Строки и столбцы своют данные вокруг значений этих полей.

Добавление **столбца "Ферма"** совокупно совокупные объемы продаж для каждой фермы. Добавление строк **"Тип"** и **"Классификация"** дополнительно разбивает данные в зависимости от того, какие именно продукты были проданы и были ли они органичными.

![PivotTable with a Farm column and Type and Classification rows.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

Кроме того, можно использовать с помощью с помощью севоводки только строк или столбцов.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Добавление иерархий данных в списоную

Иерархии данных заполняют с помощью совмещаемых данных с помощью строк и столбцов. При добавлении иерархий данных **"Crates Sold at Farm"** и **"Crates Sold SoldА"** суммы этих рисунков для каждой строки и столбца суммы.

В этом примере **"Ферма"** и **"Тип"** — это строки, данные о продажах в кавере.

![Сиветь, показывающая общий объем продаж разных видов деревьев в зависимости от фермы, из которой они поступили.](../images/excel-pivots-data-hierarchy.png)

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a>Макеты pivotTable и получение совописаемых данных

[PivotLayout](/javascript/api/excel/excel.pivotlayout) определяет размещение иерархий и их данных. Вы можете получить доступ к макету, чтобы определить диапазоны, в которых хранятся данные.

На следующей схеме показано, какие вызовы функции макета соответствуют диапазонам pivotTable.

![Схема, показывающая, какие разделы с помощью pivotTable возвращаются функциями получения диапазона макета.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>Get data from the PivotTable

Макет определяет, как отображается pivotTable на этом экране. Это `PivotLayout` означает, что объект управляет диапазонами, используемыми для элементов pivotTable. Используйте диапазоны, предоставляемые макетом, для получения данных, собранных и собранных сводной. В частности, используйте для доступа к данным, которые `PivotLayout.getDataBodyRange` создает с помощью pivotTable.

В следующем коде показано, как получить последнюю строку данных с помощью  макета (общая  сумма как суммы проданных в ферме, так и столбцов **"Сумма** проданных ящиков" в предыдущем примере). Затем эти значения суммются в итоговом итоговом значении, которое отображается в ячейке **E30** (за пределами списной).

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

### <a name="layout-types"></a>Типы макетов

Сивчатые таблицы имеют три стиля макета: Compact, Outline и Tabular. В предыдущих примерах мы видели компактный стиль.

В следующих примерах используются стили структур и таблиц соответственно. В примере кода показано, как цикли между различными макетами.

#### <a name="outline-layout"></a>Макет контура

![PivotTable using the outline layout.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>Табулярный макет

![PivotTable using the tabular layout.](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a>Удаление советнойtable

PivotTables are deleted by using their name.

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a>Фильтрация севоtTable

Основной метод фильтрации данных pivotTable — с помощью pivotFilters. Срезы предлагают альтернативный, менее гибкий метод фильтрации. 

[Фильтры pivotFilters](/javascript/api/excel/excel.pivotfilters) фильтруют данные на основе [](#hierarchies) четырех категорий иерархии (фильтров, столбцов, строк и значений) в списанной. Существует четыре типа pivotFilters, которые позволяют использовать фильтрацию на основе даты календаря, разбиение строк, сравнение номеров и фильтрацию на основе пользовательского ввода. 

[Срезы](/javascript/api/excel/excel.slicer) можно применять как к срезам, так и к обычным таблицам Excel. При применении к pivotTable срезы функционируют как [pivotManualFilter](#pivotmanualfilter) и позволяют фильтровать на основе пользовательского ввода. В отличие от PivotFilters, срезы имеют компонент [пользовательского интерфейса Excel.](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d) С помощью `Slicer` класса вы создаете этот компонент пользовательского интерфейса, управляете фильтрацией и управляете его внешним видом. 

### <a name="filter-with-pivotfilters"></a>Фильтрация с помощью pivotFilters

[Фильтры pivotFilters](/javascript/api/excel/excel.pivotfilters) позволяют фильтровать данные pivotTable на основе четырех категорий [иерархии](#hierarchies) (фильтров, столбцов, строк и значений). В объектной модели pivotTable применяются к `PivotFilters` [pivotField,](/javascript/api/excel/excel.pivotfield)и каждому из них может быть назначен один или `PivotField` `PivotFilters` несколько. Чтобы применить pivotFilters к pivotField, соответствующая [pivotHierarchy](/javascript/api/excel/excel.pivothierarchy) поля должна быть назначена категории иерархии. 

#### <a name="types-of-pivotfilters"></a>Типы pivotFilters

| Тип фильтра | Цель фильтрации | Справочные материалы по API JavaScript для Excel |
|:--- |:--- |:--- |
| DateFilter | Фильтрация на основе даты в календаре. | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | Фильтрация сравнения текста. | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | Настраиваемая фильтрация входных данных. | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | Фильтрация сравнения номеров. | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>Создание pivotFilter

Чтобы отфильтровать данные pivotTable с помощью (например, a), применим фильтр `Pivot*Filter` `PivotDateFilter` к [pivotField.](/javascript/api/excel/excel.pivotfield) В следующих четырех примерах кода покажем, как использовать каждый из четырех типов pivotFilters. 

##### <a name="pivotdatefilter"></a>PivotDateFilter

Первый пример кода применяет [pivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) к pivotField с обновлением даты, скрывая все данные до **2020-08-01**.  

> [!IMPORTANT] 
> A не может применяться к `Pivot*Filter` pivotField, если pivotHierarchy этого поля не назначена категории иерархии. В следующем примере кода необходимо добавить его в категорию pivotTable, прежде чем его можно будет использовать `dateHierarchy` `rowHierarchies` для фильтрации.

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
> В следующих трех фрагментах кода отображаются только фрагменты фильтра, а не полные `Excel.run` вызовы.

##### <a name="pivotlabelfilter"></a>PivotLabelFilter

Во втором фрагменте кода показано, как применить [pivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) к **Типу** PivotField, используя свойство для исключения меток, которые начинаются с буквы `LabelFilterCondition.beginsWith` **L**. 

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

##### <a name="pivotmanualfilter"></a>PivotManualFilter

В третьем фрагменте кода к полю классификации применяется ручной фильтр  [с pivotManualFilter,](/javascript/api/excel/excel.pivotmanualfilter) отфильтровывая данные, не включающие классификацию **"Органическая".** 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

Чтобы сравнить числа, используйте фильтр значений [с pivotValueFilter,](/javascript/api/excel/excel.pivotvaluefilter)как показано в фрагменте кода. The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Sold** PivotField, including only farms whose sum of crates sold exceeds the value **500**. 

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

#### <a name="remove-pivotfilters"></a>Удаление pivotFilters

Чтобы удалить все pivotFilters, примените метод к каждому pivotField, как показано в следующем `clearAllFilters` примере кода. 

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

### <a name="filter-with-slicers"></a>Фильтрация с срезами

[Срезы](/javascript/api/excel/excel.slicer) позволяют фильтровать данные из таблицы или таблицы Excel. Срез использует значения из указанного столбца или pivotField для фильтрации соответствующих строк. Эти значения хранятся в качестве объектов [SlicerItem](/javascript/api/excel/excel.sliceritem) в `Slicer` объекте . Надстройка может настраивать эти фильтры, как и пользователи[(с помощью пользовательского интерфейса Excel).](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d) Срез находится поверх таблицы на уровне рисования, как показано на следующем снимке экрана.

![Данные фильтрации срезов в списанной.](../images/excel-slicer.png)

> [!NOTE]
> Методы, описанные в этом разделе, посвящены использованию срезов, подключенных к списанной. Те же методы применяются и к использованию срезов, подключенных к таблицам.

#### <a name="create-a-slicer"></a>Создание среза

Вы можете создать срез в книге или на литейке с помощью `Workbook.slicers.add` метода или `Worksheet.slicers.add` метода. При этом к объекту [SlicerCollection](/javascript/api/excel/excel.slicercollection) указанного объекта добавляется `Workbook` `Worksheet` срез. Метод `SlicerCollection.add` имеет три параметра:

- `slicerSource`: источник данных, на котором основан новый срез. Это может быть `PivotTable` `Table` строка , или строка, представляющая имя или ИД `PivotTable` или `Table` .
- `sourceField`: поле в источнике данных, по которому необходимо отфильтровать данные. Это может быть `PivotField` `TableColumn` строка , или строка, представляющая имя или ИД `PivotField` или `TableColumn` .
- `slicerDestination`: таблица, на которой будет создан новый срез. Это может быть `Worksheet` объект, имя или ИД `Worksheet` объекта . Этот параметр не является нужным при `SlicerCollection` доступе к `Worksheet.slicers` нему. В этом случае в качестве места назначения используется таблица коллекции.

В следующем примере кода новый срез добавляется на таблицу **Pivot.** Источником среза является совивная область продаж фермы, которая фильтруется с использованием данных **Type.**  Срез также называется **срезом "Срезы"** для дальнейшей ссылки.

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

#### <a name="filter-items-with-a-slicer"></a>Фильтрация элементов с помощью среза

Срез фильтрует с помощью срезов с помощью элементов из `sourceField` срезов. Метод `Slicer.selectItems` задает элементы, которые остаются в срезе. Эти элементы передаются методу в качестве `string[]` ключей элементов. Все строки, содержащие эти элементы, остаются в агрегации pivotTable. Последующие `selectItems` вызовы, чтобы установить для списка ключи, указанные в этих вызовах.

> [!NOTE]
> Если передается элемент, который не находится в источнике `Slicer.selectItems` данных, будет `InvalidArgument` выброшена ошибка. Содержимое можно проверить с помощью свойства, которое является `Slicer.slicerItems` [SlicerItemCollection.](/javascript/api/excel/excel.sliceritemcollection)

В примере кода ниже показано, как выбрать три пункта для среза: **"Неугомя"** и **"Оранжевый".** 

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

Чтобы удалить все фильтры из среза, используйте метод, как `Slicer.clearFilters` показано в следующем примере.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>Стиль и форматирование среза

Надстройка может настраивать параметры отображения среза с помощью `Slicer` свойств. В следующем примере кода устанавливается стиль **SlicerStyleLight6,** замещется текст в верхней части среза **"Типы",** срез помещается в положение **(395, 15)** на уровне рисования и устанавливается размер среза **135x150** пикселей.

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

#### <a name="delete-a-slicer"></a>Удаление среза

Чтобы удалить срез, вызовите `Slicer.delete` метод. В следующем примере кода первый срез удаляется с текущего таблицы.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>Изменение функции агрегации

Иерархии данных объединяют свои значения. Для наборов данных чисел это сумма по умолчанию. Свойство определяет это поведение на основе типа `summarizeBy` [AggregationFunction.](/javascript/api/excel/excel.aggregationfunction)

Поддерживаемые в настоящее время типы функций агрегирования: `Sum` `Count` , и `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` (по умолчанию).

В следующих примерах кода агрегация изменяется на средние значения данных.

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

## <a name="change-calculations-with-a-showasrule"></a>Изменение вычислений с помощью ShowAsRule

По умолчанию сводные данные иерархий строк и столбцов объединяются независимо друг от друга. [ShowAsRule](/javascript/api/excel/excel.showasrule) изменяет иерархию данных на значения вывода на основе других элементов в pivotTable.

У `ShowAsRule` объекта есть три свойства:

- `calculation`: тип относительного вычисления, применяемого к иерархии данных (по `none` умолчанию).
- `baseField`: [PivotField](/javascript/api/excel/excel.pivotfield) в иерархии, содержащей базовые данные перед вычислением. Поскольку в список Excel есть сопоставление иерархии с полем "один к одному", для доступа к иерархии и полю используется одно и то же имя.
- `baseItem`: отдельный [pivotItem](/javascript/api/excel/excel.pivotitem) сравнивается со значениями базовых полей на основе типа вычисления. Это поле требуется не для всех вычислений.

В следующем примере вычисление в иерархии данных фермы **суммы** проданных ящиков устанавливается в процентах от общего числа столбцов.
Мы по-прежнему хотим, чтобы степень детализации была расширена  до уровня типа соков, поэтому мы будем использовать иерархию строк Type и его поле.
В примере  также в качестве первой строки иерархии показана ферма, поэтому в общем число записей фермы отображается процент, за создание которых отвечает и каждая ферма.

![СvotTable showing the percentages of the farms sales relative to the total total for both individual farms and individual farms types within each farm.](../images/excel-pivots-showas-percentage.png)

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

В предыдущем примере вычисление устанавливается для столбца относительно поля отдельной иерархии строк. Если вычисление относится к отдельному элементу, используйте `baseItem` свойство.

В следующем примере показано `differenceFrom` вычисление. Он отображает разницу в записях иерархии данных о продажах в кэш фермы относительно записей **A Farms.**
Это `baseField` **"Ферма",** поэтому мы видим различия между другими фермами, а также разбивку по каждому типу как кебайт **(Тип** также является иерархией строк в этом примере).

![Сиветь, показывающая различия в продажах деревьев между "A Farms" и другими. Это показывает разницу в общем продажах деревьев в фермах и в продажах типов. Если "A Farms" не продает определенный тип деревьев, отображается "#N/A".](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a>Изменение имен иерархии

Поля иерархии можно редактировать. В следующем коде показано, как изменить отображаемую имена двух иерархий данных.

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

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Справочник по API JavaScript для Excel](/javascript/api/excel)

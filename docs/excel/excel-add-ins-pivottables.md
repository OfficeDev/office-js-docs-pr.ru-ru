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
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Работать со сводными таблицами с помощью API JavaScript для Excel

Сводные таблицы упрощают работу с большими наборами данных. Они позволяют быстро управлять группированием данных. API JavaScript для Excel позволяет надстройке создавать сводные таблицы и взаимодействовать с их компонентами. В этой статье описывается, как сводные таблицы представлены с помощью API JavaScript для Office, а также приведены примеры кода для ключевых сценариев.

Если вы не знакомы с функциями сводных таблиц, рассмотрите возможность их изучения в качестве конечного пользователя.
Ознакомьтесь со статьей [Создание сводной таблицы, чтобы проанализировать данные листа](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) для хорошего учебника по этим средствам.

> [!IMPORTANT]
> Сводные таблицы, созданные с помощью OLAP, в настоящее время не поддерживаются. Кроме того, отсутствует поддержка Power Pivot.

## <a name="object-model"></a>Объектная модель

[Сводная таблица](/javascript/api/excel/excel.pivottable) является центральным объектом для сводных ТАБЛИЦ в API JavaScript для Office.

- `Workbook.pivotTables` и `Worksheet.pivotTables` — это [пивоттаблеколлектионс](/javascript/api/excel/excel.pivottablecollection) , которые содержат [Сводные таблицы](/javascript/api/excel/excel.pivottable) в книге и листе соответственно.
- [Сводная таблица](/javascript/api/excel/excel.pivottable) содержит [Пивосиерарчиколлектион](/javascript/api/excel/excel.pivothierarchycollection) с несколькими [пивосиерарчиес](/javascript/api/excel/excel.pivothierarchy).
- Эти [пивосиерарчиес](/javascript/api/excel/excel.pivothierarchy) можно добавить в конкретные коллекции иерархий, чтобы определить, как данные будут сведены в сводную таблицу (как описано в [следующем разделе](#hierarchies)).
- [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) содержит [пивотфиелдколлектион](/javascript/api/excel/excel.pivotfieldcollection) , в котором есть ровно один [PivotField](/javascript/api/excel/excel.pivotfield). Если проект разворачивается для включения сводных таблиц OLAP, это может измениться.
- К [PivotField](/javascript/api/excel/excel.pivotfield) может быть применено одно или несколько [PivotFilters](/javascript/api/excel/excel.pivotfilters) , если [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) поля назначено категории иерархии. 
- [PivotField](/javascript/api/excel/excel.pivotfield) содержит [Пивотитемколлектион](/javascript/api/excel/excel.pivotitemcollection) с несколькими [PivotItems](/javascript/api/excel/excel.pivotitem).
- [Сводная таблица](/javascript/api/excel/excel.pivottable) содержит объект [PivotLayout](/javascript/api/excel/excel.pivotlayout) , определяющий, где на листе отображаются [PivotFields](/javascript/api/excel/excel.pivotfield) и [PivotItems](/javascript/api/excel/excel.pivotitem) .

Рассмотрим, как эти отношения применяются к некоторым примерам данных. В приведенных ниже данных описываются продажи фруктов из различных ферм. Это будет пример во всей этой статье.

![Коллекция продаж фруктов различных типов из различных ферм.](../images/excel-pivots-raw-data.png)

Данные продаж фермы фруктов будут использоваться для создания сводной таблицы. Каждый столбец, например **types**, — это `PivotHierarchy` . Иерархия **types** содержит поле **типы** . Поле **типы** содержит элементы **Apple**, **киви**, **Лемон**, **травяные** и **оранжевые**.

### <a name="hierarchies"></a>Hierarchies

Сводные таблицы организованы в соответствии с четырьмя категориями иерархии: [строкой](/javascript/api/excel/excel.rowcolumnpivothierarchy), [столбцом](/javascript/api/excel/excel.rowcolumnpivothierarchy), [данными](/javascript/api/excel/excel.datapivothierarchy)и [фильтром](/javascript/api/excel/excel.filterpivothierarchy).

Приведенные выше данные фермы имеют пять иерархий: **фермы**, **типы**, **классификации**, **ящики**, проданные в ферме и **ящики, продаваемые оптовой торговлей**. Каждая иерархия может существовать только в одной из четырех категорий. Если **тип** добавляется к иерархиям столбцов, он также не может находиться в иерархиях "строка", "данные" или "Фильтрация". Если впоследствии **тип** добавляется к иерархиям строк, он удаляется из иерархий столбцов. Такое поведение аналогично тому, как выполняется назначение иерархии с помощью пользовательского интерфейса Excel или API JavaScript для Excel.

Иерархии строк и столбцов определяют, как группируются данные. Например, иерархия **ферм фермы** объединяет все наборы данных из одной фермы. Выбор между строкой и иерархией столбцов определяет ориентацию сводной таблицы.

Иерархии данных — это значения, которые должны быть объединены на основе иерархий строк и столбцов. Сводная таблица с иерархией **ферм** и иерархией данных для ящиков, проданных в **оптовой торговле** , показывает общую сумму (по умолчанию) всех различных Fruits для каждой фермы.

Иерархии фильтров включают или исключают данные из сводной таблицы на основе значений в этом типе фильтрации. Иерархия фильтров **классификации** **с типом "** не только выбранные" показывает только данные для придля себя фруктов.

Далее представлены данные фермы, вместе со сводной таблицей. В сводной таблице используется **ферма** и **тип** в качестве иерархий строк, **ящики** , проданные в ферме и ящики, проданные в ферме, а также **продаются по оптовой торговле** в виде иерархий данных (с использованием статистической функции по умолчанию Sum) и **классификации** в качестве иерархии фильтров ( **с выбранным** параметром "

![Выбор данных о продажах для фруктов рядом со сводной таблицей со строками, данными и иерархиями фильтров.](../images/excel-pivot-table-and-data.png)

Эту сводную таблицу можно создать с помощью API JavaScript или пользовательского интерфейса Excel. Оба варианта позволяют осуществлять дальнейшую обработку надстроек.

## <a name="create-a-pivottable"></a>Создание сводной таблицы

Для сводных таблиц требуются имя, источник и назначение. Источником может быть адрес диапазона или имя таблицы (передается как `Range` тип, `string` или `Table` тип). Назначение является адресом диапазона ( `Range` или `string` ).
В следующих примерах показаны различные методы создания сводных таблиц.

### <a name="create-a-pivottable-with-range-addresses"></a>Создание сводной таблицы с адресами диапазона

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Создание сводной таблицы с объектами Range

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

### <a name="create-a-pivottable-at-the-workbook-level"></a>Создание сводной таблицы на уровне книги

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Использование существующей сводной таблицы

Вы также можете получить доступ к сводным таблицам, созданным вручную, с помощью сводной таблицы книги или отдельных листов. В следующем коде показано получение сводной таблицы с именем **My Pivot** из книги.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Добавление строк и столбцов в сводную таблицу

Строки и столбцы поворачивают данные вокруг этих значений полей.

При добавлении столбца **фермы** все продажи для каждой фермы отворачиваются. Добавление строк **типа** и **классификации** дополнительно разделяет данные на основании того, сколько фруктов было продано, и не было ли оно согласовано.

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

Кроме того, можно создать сводную таблицу, используя только строки или столбцы.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Добавление иерархий данных в сводную таблицу

Иерархии данных заполняют сводную таблицу со сведениями, которые необходимо объединить в зависимости от строк и столбцов. Добавление иерархий данных ящиков, проданных **в ферме** и **ящиков, продаваемых в оптовой торговле** , приводит к суммированию этих значений для каждой строки и столбца.

В этом примере **ферма** и **тип** представляют собой строки, в которых продажи ящиков являются данными.

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a>Макеты сводных таблиц и извлечение сведенных данных

[PivotLayout](/javascript/api/excel/excel.pivotlayout) определяет размещение иерархий и их данных. Вы можете получить доступ к макету, чтобы определить диапазоны, в которых хранятся данные.

На следующей схеме показано, какие вызовы функций макета соответствуют какому диапазону сводной таблицы.

![Схема, на которой показано, какие разделы сводной таблицы возвращаются функциями диапазона получения в макете.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>Получение данных из сводной таблицы

Макет определяет способ отображения сводной таблицы на листе. Это означает, что `PivotLayout` объект управляет диапазонами, используемыми для элементов сводной таблицы. Используйте диапазоны, предоставленные макетом, для получения данных, собранных и агрегированных сводной таблицей. В частности, используйте `PivotLayout.getDataBodyRange` для доступа к тем, что делает Сводная таблица.

В приведенном ниже коде показано, как получить последнюю строку данных сводной таблицы, посвященную макету ( **общему** количеству **ящиков, проданных в ферме** , и **сумме ящиков, проданных** в одной колонке в предыдущем примере). Затем эти значения суммируются вместе для итогового итога, который отображается в ячейке **E30** (вне сводной таблицы).

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

В сводных таблицах есть три стиля макета: компактный, структурированный и табличный. В предыдущих примерах показан стиль "Компактный".

В приведенных ниже примерах используются структурированные и табличные стили соответственно. В примере кода показано, как циклически переключаться между различными макетами.

#### <a name="outline-layout"></a>Макет структуры

![Сводная таблица с использованием структуры.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>Табличный макет

![Сводная таблица с использованием табличного макета.](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a>Удаление сводной таблицы

Сводные таблицы удаляются с использованием их имени.

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a>Фильтрация сводной таблицы

Основным методом фильтрации данных сводной таблицы является PivotFilters. Срезы предоставляют альтернативный и менее гибкий метод фильтрации. 

[PivotFilters](/javascript/api/excel/excel.pivotfilters) фильтрация данных на основе четырех [иерархических категорий](#hierarchies) сводной таблицы (фильтров, столбцов, строк и значений). Существует четыре типа PivotFilters, позволяющие использовать фильтрацию на основе дат, анализ строк, сравнение чисел и фильтрацию на основе настраиваемого ввода. 

[Срезы](/javascript/api/excel/excel.slicer) можно применять как к сводным таблицам, так и к обычным таблицам Excel. При применении к сводной таблице срезы функционируют так же, как и [пивотмануалфилтер](#pivotmanualfilter) , и позволяют выполнять фильтрацию на основе настраиваемого ввода. В отличие от PivotFilters, срезы имеют [компонент пользовательского интерфейса Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d). С помощью `Slicer` класса вы создадите этот компонент пользовательского интерфейса, управляете фильтрацией и контролируйте его внешний вид. 

### <a name="filter-with-pivotfilters"></a>Фильтрация с помощью PivotFilters

[PivotFilters](/javascript/api/excel/excel.pivotfilters) позволяют фильтровать данные сводной таблицы на основе четырех [категорий иерархии](#hierarchies) (фильтров, столбцов, строк и значений). В объектной модели сводной таблицы `PivotFilters` применяются к [PivotField](/javascript/api/excel/excel.pivotfield), и у каждого из них `PivotField` может быть один или несколько назначенных `PivotFilters` . Чтобы применить PivotFilters к PivotField, соответствующему [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) поля необходимо назначить категории иерархии. 

#### <a name="types-of-pivotfilters"></a>Типы PivotFilters

| Тип фильтра | Назначение фильтра | Справочные материалы по API JavaScript для Excel |
|:--- |:--- |:--- |
| датефилтер | Фильтрация на основе даты в календаре. | [пивотдатефилтер](/javascript/api/excel/excel.pivotdatefilter) |
| лабелфилтер | Фильтрация по текстовому сравнению. | [пивотлабелфилтер](/javascript/api/excel/excel.pivotlabelfilter) |
| мануалфилтер | Настраиваемый фильтр ввода. | [пивотмануалфилтер](/javascript/api/excel/excel.pivotmanualfilter) |
| валуефилтер | Фильтрация сравнения номеров. | [пивотвалуефилтер](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>Создание PivotFilter

Чтобы отфильтровать данные сводной таблицы с помощью сводного фильтра (например, Пивотдатефилтер), примените фильтр к [PivotField](/javascript/api/excel/excel.pivotfield). В следующих четырех примерах кода показано, как использовать каждый из четырех типов PivotFilters. 

##### <a name="pivotdatefilter"></a>пивотдатефилтер

Первый пример кода применяет [пивотдатефилтер](/javascript/api/excel/excel.pivotdatefilter) к **дате обновления** PivotField, скрывая все данные до **2020-08-01**. 

> [!IMPORTANT] 
> Фильтр PIVOT нельзя применить к PivotField, если это поле PivotHierarchy не назначено категории иерархии. В следующем примере кода `dateHierarchy` необходимо добавить в категорию сводной таблицы, `rowHierarchies` прежде чем его можно будет использовать для фильтрации.

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
> В следующих трех фрагментах кода отображаются только отрывок, относящиеся к фильтрам, а не полные `Excel.run` вызовы.

##### <a name="pivotlabelfilter"></a>пивотлабелфилтер

Во втором фрагменте кода показано, как применить [пивотлабелфилтер](/javascript/api/excel/excel.pivotlabelfilter) к **типу** PivotField, используя свойство, `LabelFilterCondition.beginsWith` чтобы исключить метки, начинающиеся с буквы **L**. 

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

##### <a name="pivotmanualfilter"></a>пивотмануалфилтер

Третий фрагмент кода применяет к полю **классификации** вручную фильтр с [пивотмануалфилтер](/javascript/api/excel/excel.pivotmanualfilter) , отфильтровывая данные, которые не включают согласованности классификации **.** 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>пивотвалуефилтер

Чтобы сравнить числа, используйте фильтр значений с [пивотвалуефилтер](/javascript/api/excel/excel.pivotvaluefilter), как показано в последнем фрагменте кода. В этом `PivotValueFilter` разделе сравниваются данные в **ферме** PivotField с данными в рабочих ящиках, проданных в **оптовой торговле** PivotField, включая только те фермы, сумма которых проданных ящиков превышает значение **500**. 

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

#### <a name="remove-pivotfilters"></a>Удаление PivotFilters

Чтобы удалить все PivotFilters, примените `clearAllFilters` метод к каждому PivotField, как показано в следующем примере кода. 

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

### <a name="filter-with-slicers"></a>Фильтрация с помощью срезов

[Срезы](/javascript/api/excel/excel.slicer) позволяют фильтровать данные из сводной таблицы или таблицы Excel. Срез использует значения из указанного столбца или PivotField для фильтрации соответствующих строк. Эти значения хранятся в виде объектов [SlicerItem](/javascript/api/excel/excel.sliceritem) в `Slicer` . Надстройка может настраивать эти фильтры, как это могут делать пользователи ([через пользовательский интерфейс Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)). Срез располагается вверху листа в графическом слое, как показано на следующем снимке экрана.

![Фильтрация данных среза в сводной таблице.](../images/excel-slicer.png)

> [!NOTE]
> Методы, описанные в этом разделе, касаются использования срезов, подключенных к сводным таблицам. Те же методы применяются и для использования срезов, подключенных к таблицам.

#### <a name="create-a-slicer"></a>Создание среза

Вы можете создать срез в книге или листе с помощью `Workbook.slicers.add` метода или `Worksheet.slicers.add` метода. Это приведет к добавлению среза в [слицерколлектион](/javascript/api/excel/excel.slicercollection) указанного `Workbook` или `Worksheet` объекта. `SlicerCollection.add`Метод имеет три параметра:

- `slicerSource`: Источник данных, на котором основан новый срез. Это может быть `PivotTable` , `Table` или строка, представляющая имя или идентификатор или `PivotTable` `Table` .
- `sourceField`: Поле в источнике данных, с помощью которого выполняется фильтрация. Это может быть `PivotField` , `TableColumn` или строка, представляющая имя или идентификатор или `PivotField` `TableColumn` .
- `slicerDestination`: Лист, на котором будет создан новый срез. Это может быть `Worksheet` объект или имя или идентификатор объекта `Worksheet` . Этот параметр не является обязательным при `SlicerCollection` доступе к `Worksheet.slicers` . В этом случае лист коллекции используется в качестве назначения.

В приведенном ниже примере кода в **сводную** таблицу добавляется новый срез. Источник среза — это сводная таблица и фильтры **продаж фермы** с использованием данных **типа** . Срез также называется **срезом фруктов** для дальнейшего использования.

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

Срез фильтрует сводную таблицу с элементами из `sourceField` . `Slicer.selectItems`Метод задает элементы, остающиеся в срезе. Эти элементы передаются в метод как объект `string[]` , представляющий ключи элементов. Все строки, содержащие эти элементы, сохраняются в статистической обработке сводной таблицы. Последующие вызовы `selectItems` задают для списка ключи, указанные в этих вызовах.

> [!NOTE]
> Если `Slicer.selectItems` передается элемент, который не находится в источнике данных, `InvalidArgument` возникает ошибка. Содержимое можно проверить с помощью `Slicer.slicerItems` свойства, которое является [слицеритемколлектион](/javascript/api/excel/excel.sliceritemcollection).

В приведенном ниже примере кода показаны три выбранных для среза элементов: **Лемон**, **травяной** и **оранжевый**.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

Чтобы удалить все фильтры из среза, используйте `Slicer.clearFilters` метод, как показано в следующем примере.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>Стиль и форматирование среза

Надстройка может настраивать параметры отображения среза с помощью `Slicer` свойств. В приведенном ниже примере кода для стиля задается значение **SlicerStyleLight6**, в верхней части среза задается **Тип фруктов**, помещается срез в позицию **(395, 15)** на уровне рисунка и задается размер среза **135x150** пикселей.

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

Чтобы удалить срез, вызовите `Slicer.delete` метод. В примере кода ниже показано, как удалить первый срез из текущего листа.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>Изменение статистической функции

Иерархия данных содержит статистические значения. Для наборов данных Numbers это сумма по умолчанию. `summarizeBy`Свойство определяет это поведение на основе типа [аггрегатионфунктион](/javascript/api/excel/excel.aggregationfunction) .

В настоящее время поддерживаются типы статистической функции,,,,,,,,, `Sum` `Count` `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` и `Automatic` (значение по умолчанию).

В приведенных ниже примерах кода статистическая схема изменяется для средних значений данных.

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

## <a name="change-calculations-with-a-showasrule"></a>Изменение вычислений с помощью Шовасруле

Сводные таблицы по умолчанию объединяют данные иерархий строк и столбцов независимо друг от друга. [Шовасруле](/javascript/api/excel/excel.showasrule) изменяет иерархию данных на выходные значения на основе других элементов в сводной таблице.

`ShowAsRule`У объекта есть три свойства:

- `calculation`: Тип относительного вычисления, применяемого к иерархии данных (значение по умолчанию — `none` ).
- `baseField`: [PivotField](/javascript/api/excel/excel.pivotfield) в иерархии, содержащей базовые данные перед применением вычисления. Так как сводные таблицы Excel имеют сопоставление "один к одному" в поле "иерархия", для доступа к иерархии и полю используется то же имя.
- `baseItem`: Отдельные [PivotItem](/javascript/api/excel/excel.pivotitem) по сравнению со значениями базовых полей на основе типа вычисления. Для этого поля требуется не все вычисления.

В следующем примере показана настройка вычисления **суммы ящиков, проданных в** иерархии данных фермы, в процентах от общей суммы по столбцу.
Мы по-прежнему хотим, чтобы гранулярность была расширена до уровня типа фруктов, поэтому мы будем использовать иерархию **типов** строк и базовое поле.
В примере также используется **ферма** в качестве первой иерархии строк, поэтому записи итоговой фермы отображаются в процентах, ответственных за изготовление.

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

В предыдущем примере показано, как задать вычисление для столбца относительно поля отдельной иерархии строк. Когда расчет относится к отдельному элементу, используйте `baseItem` свойство.

В приведенном ниже примере показано `differenceFrom` вычисление. В нем отображается разность записей иерархии данных о продажах в ферме, относящихся к параметрам **ферм**.
`baseField` **Ферма** состоит в том, что мы видим различия между другими фермами, а также подразделение для каждого типа вроде фруктов (**тип** также является иерархией строк в данном примере).

![Сводная таблица, в которой показаны различия продаж фруктов между "фермами" и другими. В этом примере показана разница в общем объеме продаж фруктов ферм и продаж на различных типах фруктов. Если "фермы" не продают определенный тип фруктов, отображается "#N/A".](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a>Изменение имен иерархий

Поля иерархии можно редактировать. В приведенном ниже коде показано, как изменить отображаемые имена двух иерархий данных.

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

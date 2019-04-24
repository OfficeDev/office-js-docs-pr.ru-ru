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
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Работать со сводными таблицами с помощью API JavaScript для Excel

Сводные таблицы упрощают работу с большими наборами данных. Они позволяют быстро управлять группированием данных. API JavaScript для Excel позволяет надстройке создавать сводные таблицы и взаимодействовать с их компонентами.

Если вы не знакомы с функциями сводных таблиц, рассмотрите возможность их изучения в качестве конечного пользователя. Ознакомьтесь со статьей [Создание сводной таблицы, чтобы проанализировать данные листа](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) для хорошего учебника по этим средствам. 

В этой статье приведены примеры кода для распространенных сценариев. Подробнее об API сводных таблиц можно узнать в статье [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).

> [!IMPORTANT]
> Сводные таблицы, созданные с помощью OLAP, в настоящее время не поддерживаются. Кроме того, отсутствует поддержка Power Pivot.

## <a name="hierarchies"></a>Hierarchies

Сводные таблицы организованы в соответствии с четырьмя категориями иерархии: строкой, столбцом, данными и фильтром. В этой статье будут использоваться следующие данные, описывающие продажи фруктов из различных ферм.

![Коллекция продаж фруктов различных типов из различных ферм.](../images/excel-pivots-raw-data.png)

Эти данные имеют пять иерархий: **ферм**, **типов**, **классификаций**, ящиков, **проданных в ферме**, и ящики, продаваемые **оптовой торговлей**. Каждая иерархия может существовать только в одной из четырех категорий. Если **тип** добавляется к иерархиям столбцов и затем добавляется к иерархиям строк, он остается только последним.

Иерархии строк и столбцов определяют, как группируются данные. Например, иерархия **ферм фермы** объединяет все наборы данных из одной фермы. Выбор между строкой и иерархией столбцов определяет ориентацию сводной таблицы.

Иерархии данных — это значения, которые должны быть объединены на основе иерархий строк и столбцов. Сводная таблица с иерархией **ферм** и иерархией данных для ящиков, проданных в **оптовой торговле** , показывает общую сумму (по умолчанию) всех различных Fruits для каждой фермы.

Иерархии фильтров включают или исключают данные из сводной таблицы на основе значений в этом типе фильтрации. Иерархия фильтров **классификации** с типом "не **** только выбранные" показывает только данные для придля себя фруктов.

Далее представлены данные фермы, вместе со сводной таблицей. В сводной таблице используется **ферма** и **тип** в качестве иерархий строк, ящики, проданные **на ферме** и ящики, проданные по **оптовой торговле** в виде иерархий данных (с статистической функцией статистической обработки по умолчанию Sum), а **классификация** — как фильтр. иерархия ( **** с выделенным параметром). 

![Выбор данных о продажах для фруктов рядом со сводной таблицей со строками, данными и иерархиями фильтров.](../images/excel-pivot-table-and-data.png)

Эту сводную таблицу можно создать с помощью API JavaScript или ПОЛЬЗОВАТЕЛЬСКОГО интерфейса Excel. Оба варианта позволяют осуществлять дальнейшую обработку надстроек.

## <a name="create-a-pivottable"></a>Создание сводной таблицы

Для сводных таблиц требуются имя, источник и назначение. Источником может быть адрес диапазона или имя таблицы (передается как тип `Range`, `string`или `Table` тип). Назначение является адресом диапазона ( `Range` или `string`). В следующих примерах показаны различные методы создания сводных таблиц.

### <a name="create-a-pivottable-with-range-addresses"></a>Создание сводной таблицы с адресами диапазона

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Создание сводной таблицы с объектами Range

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

### <a name="create-a-pivottable-at-the-workbook-level"></a>Создание сводной таблицы на уровне книги

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Использование существующей сводной таблицы

Вы также можете получить доступ к сводным таблицам, созданным вручную, с помощью сводной таблицы книги или отдельных листов. 

Приведенный ниже код получает первую сводную таблицу в книге. Затем имя таблицы придается имени для упрощения справочных материалов.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Добавление строк и столбцов в сводную таблицу

Строки и столбцы поворачивают данные вокруг этих значений полей.

При добавлении столбца **фермы** все продажи для каждой фермы отворачиваются. Добавление строк **типа** и **классификации** дополнительно разделяет данные на основании того, сколько фруктов было продано, и не было ли оно согласовано.

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

Кроме того, можно создать сводную таблицу, используя только строки или столбцы.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Добавление иерархий данных в сводную таблицу

Иерархии данных заполняют сводную таблицу со сведениями, которые необходимо объединить в зависимости от строк и столбцов. Добавление иерархий данных ящиков, проданных **в ферме** и ящиков, продаваемых в **оптовой торговле** , приводит к суммированию этих значений для каждой строки и столбца. 

В этом примере **ферма** и **тип** представляют собой строки, в которых продажи ящиков являются данными. 

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

## <a name="change-aggregation-function"></a>Изменение статистической функции

Иерархия данных содержит статистические значения. Для наборов данных Numbers это сумма по умолчанию. `summarizeBy` Свойство определяет это поведение на основе типа [аггрегатионфунктион](/javascript/api/excel/excel.aggregationfunction) .

`Sum`В настоящее время поддерживаются типы статистической `Count`функции `Average`, `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP`,,,,,,,, и `Automatic` (значение по умолчанию).

В приведенных ниже примерах кода статистическая схема изменяется для средних значений данных.

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

## <a name="change-calculations-with-a-showasrule"></a>Изменение вычислений с помощью Шовасруле

Сводные таблицы по умолчанию объединяют данные иерархий строк и столбцов независимо друг от друга. [Шовасруле](/javascript/api/excel/excel.showasrule) изменяет иерархию данных на выходные значения на основе других элементов в сводной таблице.

У `ShowAsRule` объекта есть три свойства:

-   `calculation`: Тип относительного вычисления, применяемого к иерархии данных (значение по умолчанию — `none`).
-   `baseField`: Поле в иерархии, содержащее базовые данные перед применением вычисления. [PivotField](/javascript/api/excel/excel.pivotfield) обычно имеет то же имя, что и его родительская иерархия.
-   `baseItem`: Отдельные [PivotItem](/javascript/api/excel/excel.pivotitem) по сравнению со значениями базовых полей на основе типа вычисления. Для этого поля требуется не все вычисления.

В следующем примере показана настройка вычисления **суммы ящиков, проданных в** иерархии данных фермы, в процентах от общей суммы по столбцу. Мы по-прежнему хотим, чтобы гранулярность была расширена до уровня типа фруктов, поэтому мы будем использовать иерархию **типов** строк и базовое поле. В примере также используется **ферма** в качестве первой иерархии строк, поэтому записи итоговой фермы отображаются в процентах, ответственных за изготовление.

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

В предыдущем примере показано, как задать вычисление для столбца относительно иерархии отдельных строк. Когда расчет относится к отдельному элементу, используйте `baseItem` свойство.

В приведенном ниже примере `differenceFrom` показано вычисление. В нем отображается разность записей иерархии данных о продажах в ферме, относящихся к параметрам "фермы".
Ферма `baseField` состоит **** в том, что мы видим различия между другими фермами, а также подразделение для каждого типа вроде фруктов (**тип** также является иерархией строк в данном примере).

![Сводная таблица, в которой показаны различия продаж фруктов между "фермами" и другими. В этом примере показана разница в общем объеме продаж фруктов ферм и продаж на различных типах фруктов. Если "фермы" не продают определенный тип фруктов, отображается "#N/A".](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a>Макеты сводных таблиц

[PivotLayout](/javascript/api/excel/excel.pivotlayout) определяет размещение иерархий и их данных. Вы можете получить доступ к макету, чтобы определить диапазоны, в которых хранятся данные.

На следующей схеме показано, какие вызовы функций макета соответствуют какому диапазону сводной таблицы.

![Схема, на которой показано, какие разделы сводной таблицы возвращаются функциями диапазона получения в макете.](../images/excel-pivots-layout-breakdown.png)

В приведенном ниже коде показано, как получить последнюю строку данных сводной таблицы, прополнив макет. Затем эти значения суммируются вместе для общего итога.

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

В сводных таблицах есть три стиля макета: компактный, структурированный и табличный. В предыдущих примерах показан стиль "Компактный". 

В приведенных ниже примерах используются структурированные и табличные стили соответственно. В примере кода показано, как циклически переключаться между различными макетами.

### <a name="outline-layout"></a>Макет структуры

![Сводная таблица с использованием структуры.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>Табличный макет

![Сводная таблица с использованием табличного макета.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a>Изменение имен иерархий

Поля иерархии можно редактировать. В приведенном ниже коде показано, как изменить отображаемые имена двух иерархий данных.

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

## <a name="delete-a-pivottable"></a>Удаление сводной таблицы

Сводные таблицы удаляются с использованием их имени.

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Справочник по API JavaScript для Excel](/javascript/api/excel)

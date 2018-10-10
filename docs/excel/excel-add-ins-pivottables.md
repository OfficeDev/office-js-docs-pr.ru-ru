---
title: Работа со сводными таблицами с помощью API JavaScript для Excel
description: Использование Excel JavaScript API для создания сводных таблиц и взаимодействия с их компонентами.
ms.date: 09/21/2018
ms.openlocfilehash: 00dd982d4ba4de0db34277cd546b572d4394e258
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459282"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Работа со сводными таблицами с помощью API JavaScript для Excel

Сводные таблицы упрощают создание больших наборов данных. Они позволяют быстро манипулировать сгруппированными данными. Excel JavaScript API позволяет вашей надстройке создавать сводные таблицы и взаимодействовать со своими компонентами. 

Если вы не знакомы с функциональностью сводных таблиц, рекомендуем рассмотреть их в качестве конечного пользователя. В качестве учебника для начинающих по этим инструментам см. статью [Создание сводной таблицы для анализа данных листа](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables). 

В этой статье приведены примеры кода для распространенных сценариев. Для дальнейшего понимание API сводной таблицы, просмотрите [**Сводные таблицы**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) и [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).

> [!IMPORTANT]
> Сводные таблицы, созданные с помощью OLAP, в настоящее время не поддерживаются.

## <a name="hierarchies"></a>Иерархии

Сводные таблицы организованы на основе четырех категорий иерархии: строка, столбец, данные и фильтр. Следующие данные, описывающие продажи фруктов из разных ферм, будут использоваться в этой статье.

![Коллекция продажи фруктов различных типов из нескольких ферм.](../images/excel-pivots-raw-data.png)

Эти данные имеют пять иерархий: **Фермы**, **Тип**, **Классификация**, **Ящики, проданные на ферме** и **Ящики, проданные оптом**. Каждая иерархия может существовать только в одной из четырех категорий. Если **Тип** добавляется к иерархиям столбцов, а затем добавляется к иерархиям строк, он остается только в последних.

Иерархии строк и столбцов определяют, как будут группироваться данные. Например, иерархия строки **Фермы** объединит все наборы данных из одной фермы. Выбор между иерархией строк и столбцов определяет ориентацию сводной таблицы.

Иерархии данных - это агрегированные значения, основанные на иерархиях строк и столбцов. Сводная таблица с иерархией строк **Фермы** и иерархия данных **Ящики, проданные оптом** показывает общую сумму (по умолчанию) всех разных фруктов для каждой фермы.

Иерархии фильтров включают или исключают данные из сводного документа на основе значений в этом отфильтрованном типе. Иерархия фильтра **Классификация** с выбранным типом **Органика** отображает только данные для органических фруктов.

Вот опять данные фермы, вместе со сводной таблицей. Сводная таблица использует **Ферму** и **Тип** в качестве иерархий строк, **Ящики, проданные на ферме** и **Ящики, проданные оптом** в качестве иерархий данных (с функцией агрегации по умолчанию для суммы) и **Классификацию** в качестве иерархии фильтров (при выборе **Органики**). 

![Выбор данных о продажах фруктов рядом со сводной таблицей с иерархиями строки, данных и фильтра.](../images/excel-pivot-table-and-data.png)

Эта сводная таблица может быть сгенерирована через API JavaScript или через интерфейс Excel. Оба параметра позволяют осуществлять дальнейшие манипуляции с помощью надстроек.

## <a name="create-a-pivottable"></a>Создание сводной таблицы

Для сводных таблиц требуется имя, источник и место назначения. Источником может быть адрес диапазона или имя таблицы (передано как тип `Range`, `string` или `Table`). Адрес назначения - это адрес диапазона (заданный как либо `Range`, либо `string`). Следующие примеры показывают различные методы создания сводной таблицы.

### <a name="create-a-pivottable-with-range-addresses"></a>Создание сводной таблицы с помощью адресов диапазона

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Создание сводной таблицы с помощью объектов диапазона

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

### <a name="create-a-pivottable-at-the-workbook-level"></a>Создание сводной таблицы на уровне рабочей книги

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Использование существующей сводной таблицы

Созданные вручную сводные таблицы, также доступны через коллекцию сводной таблицы рабочей книги или отдельных листов. 

Следующий код получает первую сводную таблицу в книге. Затем он присваивает таблице имя для удобства ссылки в будущем.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Добавление строк и столбцов в сводную таблицу

Строки и столбцы сводят данные, применимые к тем значениям полей.

Добавление столбца **Ферма** выполняет сведение всех продаж, относящихся к каждой ферме. Добавление строк **Тип** и **Классификация** дополнительно разбивает данные на основе того, какие фрукты были проданы, и были ли они органическими или нет.

![Сводная таблица со столбцом Ферма и строками Тип и Классификация.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

Вы также можете иметь сводную таблицу только строк или столбцов.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Добавление иерархий данных сводным таблицам

Иерархии данных заполняют сводную таблицу информацией для объединения на основе строк и столбцов. Добавление иерархий данных **Ящики, проданные на ферме** и **Ящики, проданные оптом** дает суммы этих цифр для каждой строки и столбца. 

В примере, как **Ферма**, так и **Тип** являются строками с данными продаж ящиков. 

![Сводная таблица показывает сумму всех продаж разных фруктов в ферме, в зависимости от фермы их происхождения.](../images/excel-pivots-data-hierarchy.png)

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

Иерархии данных имеют свои агрегированные значения. Для наборов данных чисел это значение по умолчанию. Свойство `summarizeBy` определяет эту реакцию на событие на основе типа `AggregrationFunction`. 

В настоящее время поддерживаются следующие типы статистической функции `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, и `Automatic` (по умолчанию).

В следующих примерах кода изменяется агрегирование для средних значений данных.

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

## <a name="change-calculations-with-a-showasrule"></a>Изменение расчетов с помощью ShowAsRule

Сводные таблицы по умолчанию агрегируют иерархии данных строк и столбцов независимо друг от друга. Объект `ShowAsRule` изменяет иерархии данных для вывода значения на основании других элементов в сводной таблице.

У объекта  `ShowAsRule` три свойства:
-   `calculation`: тип относительного вычисления, которое будет применено к иерархии данных (значение по умолчанию — `none`).
-   `baseField`: поле в иерархии, содержащее базовые данные перед вычислением. `PivotField` обычно имеет то же имя, что и у его родительской иерархии.
-   `baseItem`: отдельный элемент, сравниваемый со значениями базовых полей на основе типа вычислений. Это поле требуется не для всех вычислений.

В следующем примере выполняется вычисление на иерархии данных **Sum of Crates Sold at Farm** (Сумма проданных на ферме ящиков) в процентах от итогового значения по столбцу. Мы по-прежнему хотим, чтобы степень детализации расширялась до уровня типа фруктов, поэтому мы будем использовать иерархию строки **Type** (Тип) и ее базовое поле. В примере также есть **Farm** (Ферма) в качестве иерархии первой строки, в результате чего итоговые записи для ферм показывают процент, который каждая ферма вкладывает в общее производство.

![Сводная таблица, отражающая процент продаж фруктов относительно общей суммы как для отдельных ферм, так и для отдельных типов фруктов в каждой ферме.](../images/excel-pivots-showas-percentage.png)

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

Предыдущий пример задает вычисление по столбцу относительно отдельной иерархии строк. Когда вычисления относятся к отдельному элементу, используйте свойство `baseItem`. 

Следующий пример демонстрирует вычисление `differenceFrom`. Он отображает разницу в иерархических позициях данных по продажам фермами ящиков по сравнению с соответствующими позициями ферм «A». `baseField` представляет **фермы**, поэтому мы видим различия между другими фермами, а также разбивку по типам, например фруктам (**Type** [Тип] также является иерархией строк в этом примере).

![Сводная таблица, демонстрирующая различия в продажах фруктов между фермами «A» и другими фермами. Она показывает как разницу в общем объеме продаж фруктов фермерских хозяйств, так и различия в продажах по отдельным типам  фруктов. В тех случаях, когда фермы «A» не продают определенный тип фруктов, то отображается «#N/A».](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a>Макеты сводной таблицы

Макет сводной таблицы определяет положение иерархий и их данных. Доступ к макету для определения диапазонов, где хранятся данные. 

На следующей диаграмме показано, какие функции макета соответствуют каким диапазонам из сводной таблицы.

![Диаграмма, показывающая, какие части сводной таблицы возвращают функции get диапазона макета.](../images/excel-pivots-layout-breakdown.png)

Приведенный ниже код показывает получение последней строки сводной таблицы данных через макет. Затем эти значения суммируются друг с другом для общего итога.

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

Сводные таблицы используются три стиля макета: сжатый, структурный и табличный. Мы видели сжатый стиль в предыдущих примерах. 

В приведенных ниже примерах используется, соответственнo, структурный и табличный стили, соответственно. В примере кода показано, как переключаться между различные макетами.

### <a name="outline-layout"></a>Структурный макет

![Использование макета структуры сводной таблицы.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>Табличный макет

![Использование макета таблицы сводной таблицы.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a>Изменение имен иерархий

Поля иерархии доступны для редактирования. Следующий код демонстрирует как изменить отображаемые имена двух иерархий данных.

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

Сводная таблица удаляется по имени.

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a>См. также

- [Фундаментальные понятия программирования с использованием интерфейса API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Ссылка по API JavaScript для Excel](https://docs.microsoft.com/javascript/api/excel)
